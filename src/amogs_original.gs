// ============================== amogs_original.gs ==============================
// Оригинальная рабочая версия из вашего кода

const config = {
  AMO_TOKEN: '',
  AMO_SUBDOMAIN: '',
  SHEET_NAME: ''
};

function getConfigFromSheet() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("БД");
  config.AMO_TOKEN = sheet.getRange("B2").getValue();
  
  // Нормализация поддомена
  let subdomain = sheet.getRange("B3").getValue();
  if (subdomain) {
    subdomain = String(subdomain).trim().replace(/^https?:\/\//i, '').replace(/\.amocrm\.(ru|eu)$/i, '');
  }
  config.AMO_SUBDOMAIN = subdomain;
  
  config.SHEET_NAME = sheet.getRange("B4").getValue();
  return config;
}

function readExcludedFields() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("БД");
  const range = sheet.getRange("A8:A100");
  const values = range.getValues();
  return values.filter(row => row[0]).map(row => row[0]);
}

function formatTimestamp(timestamp) {
  return new Date(timestamp * 1000).toLocaleString();
}

function isCallProcessed(callId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("АМО Звонки");
  if (!sheet) return false;
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return false;
  
  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return data.some(row => row[0] === callId);
}

function amoRequest(url, method = 'GET', body = null) {
  const options = {
    method: method,
    headers: { 'Authorization': `Bearer ${config.AMO_TOKEN}` },
    muteHttpExceptions: true,
    followRedirects: true
  };
  
  if (body) {
    options.payload = JSON.stringify(body);
    options.contentType = 'application/json';
  }
  
  const response = UrlFetchApp.fetch(url, options);
  const statusCode = response.getResponseCode();
  
  if (statusCode < 200 || statusCode >= 300) {
    console.log(`HTTP ${statusCode}: ${response.getContentText()}`);
    return null;
  }
  
  const responseText = response.getContentText();
  return responseText ? JSON.parse(responseText) : null;
}

function getAmoTasks(leadId) {
  try {
    const url = `https://${config.AMO_SUBDOMAIN}.amocrm.ru/api/v4/leads/${leadId}/tasks`;
    const response = amoRequest(url);
    return response ? response._embedded.tasks : [];
  } catch (error) {
    console.log('Ошибка получения задач:', error);
    return 'Ошибка';
  }
}

function getAllEntityFields(entityType, entityId) {
  try {
    const url = `https://${config.AMO_SUBDOMAIN}.amocrm.ru/api/v4/${entityType}/${entityId}`;
    const response = amoRequest(url);
    return response ? response : null;
  } catch (error) {
    console.log('Ошибка получения полей:', error);
    return null;
  }
}

function getAmoLead(leadId) {
  return getAllEntityFields('leads', leadId);
}

function getAmoContact(contactId) {
  return getAllEntityFields('contacts', contactId);
}

function getAmoCompany(companyId) {
  return getAllEntityFields('companies', companyId);
}

function getAmoNotes(entityType, entityId) {
  try {
    const url = `https://${config.AMO_SUBDOMAIN}.amocrm.ru/api/v4/${entityType}/${entityId}/notes`;
    const response = amoRequest(url);
    return response ? response._embedded.notes : [];
  } catch (error) {
    console.log('Ошибка получения заметок:', error);
    return [];
  }
}

function processEvent(eventData) {
  if (eventData.type !== 'outgoing_call' && eventData.type !== 'incoming_call') {
    return;
  }
  
  const callId = eventData.id;
  if (isCallProcessed(callId)) {
    return;
  }
  
  const leadId = eventData.entity_id;
  const lead = getAmoLead(leadId);
  if (!lead) return;
  
  const contact = lead._embedded.contacts[0] ? getAmoContact(lead._embedded.contacts[0].id) : null;
  const company = lead._embedded.companies[0] ? getAmoCompany(lead._embedded.companies[0].id) : null;
  
  const notes = getAmoNotes('leads', leadId);
  const callNote = notes.find(n => n.id === callId);
  
  if (!callNote) return;
  
  const callData = {
    callId: callId,
    leadId: leadId,
    leadName: lead.name,
    contactName: contact ? contact.name : '',
    companyName: company ? company.name : '',
    callType: eventData.type === 'outgoing_call' ? 'Исходящий' : 'Входящий',
    callDate: formatTimestamp(eventData.created_at),
    duration: callNote.params.duration || 0,
    audioUrl: callNote.params.link || '',
    callStatus: callNote.params.call_status || '',
    responsible: lead.responsible_user_id
  };
  
  appendToSheet(callData);
}

function appendToSheet(callData) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("АМО Звонки");
  if (!sheet) {
    SpreadsheetApp.getActive().insertSheet("АМО Звонки");
    const newSheet = SpreadsheetApp.getActive().getSheetByName("АМО Звонки");
    newSheet.getRange(1, 1, 1, 10).setValues([[
      'ID звонка', 'ID сделки', 'Название сделки', 'Контакт', 'Компания', 
      'Тип звонка', 'Дата звонка', 'Длительность', 'Ссылка на запись', 'Статус звонка'
    ]]);
  }
  
  const targetSheet = SpreadsheetApp.getActive().getSheetByName("АМО Звонки");
  targetSheet.appendRow([
    callData.callId,
    callData.leadId,
    callData.leadName,
    callData.contactName,
    callData.companyName,
    callData.callType,
    callData.callDate,
    callData.duration,
    callData.audioUrl,
    callData.callStatus
  ]);
}

function syncEventsToday() {
  getConfigFromSheet();
  
  const timeFrom = Math.floor((new Date().getTime() - 6 * 60 * 60 * 1000) / 1000);
  const url = `https://${config.AMO_SUBDOMAIN}.amocrm.ru/api/v4/events?filter[created_at][from]=${timeFrom}&limit=250`;
  
  let allEvents = [];
  let nextUrl = url;
  
  while (nextUrl) {
    const response = amoRequest(nextUrl);
    if (!response || !response._embedded) break;
    
    allEvents = allEvents.concat(response._embedded.events);
    nextUrl = response._links.next ? response._links.next.href : null;
  }
  
  const calls = allEvents.filter(e => 
    ['outgoing_call', 'incoming_call'].includes(e.type)
  );
  
  calls.forEach(event => processEvent(event));
  
  console.log(`Обработано ${calls.length} звонков`);
}

function logEvents() {
  getConfigFromSheet();
  
  const timeFrom = Math.floor((new Date().getTime() - 150 * 60 * 1000) / 1000);
  const url = `https://${config.AMO_SUBDOMAIN}.amocrm.ru/api/v4/events?filter[created_at][from]=${timeFrom}&limit=250`;
  
  const response = amoRequest(url);
  if (!response || !response._embedded) {
    console.log('Нет событий');
    return;
  }
  
  const events = response._embedded.events;
  console.log(`Найдено ${events.length} событий`);
  
  events.forEach(event => {
    console.log(`${event.type} - ${event.entity_type} - ${event.id}`);
  });
}