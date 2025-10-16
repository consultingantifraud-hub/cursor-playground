// ============================== amogs_clean.gs ==============================
// Чистая версия без проблем с кодировкой

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
  try {
    const options = {
      method: method,
      headers: { 
        'Authorization': `Bearer ${config.AMO_TOKEN}`,
        'Accept': 'application/hal+json'
      },
      muteHttpExceptions: true,
      followRedirects: true,
      timeout: 30000
    };
    
    if (body) {
      options.payload = JSON.stringify(body);
      options.contentType = 'application/json';
    }
    
    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();
    
    if (statusCode < 200 || statusCode >= 300) {
      const errorText = response.getContentText();
      console.log(`❌ HTTP ${statusCode}: ${errorText.slice(0, 200)}`);
      return null;
    }
    
    const responseText = response.getContentText();
    return responseText ? JSON.parse(responseText) : null;
    
  } catch (error) {
    console.log(`❌ Ошибка запроса к ${url}:`, error.message);
    return null;
  }
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

function getAmoNoteById(noteId) {
  try {
    const url = `https://${config.AMO_SUBDOMAIN}.amocrm.ru/api/v4/notes/${noteId}`;
    const response = amoRequest(url);
    return response;
  } catch (error) {
    console.log(`Ошибка получения заметки ${noteId}:`, error);
    return null;
  }
}

function isCallNote(note) {
  if (!note) return false;
  
  const noteType = String(note.note_type || '').toLowerCase();
  const params = note.params || {};
  
  // Стандартные типы звонков
  if (/^(call_in|call_out)$/i.test(noteType)) return true;
  
  // Гибридные звонки: common заметки с признаками звонка
  if (noteType === 'common') {
    const hasLink = params.link && /^https?:\/\//i.test(String(params.link));
    const hasDuration = params.duration && Number(params.duration) > 0;
    const hasCallStatus = params.call_status !== undefined;
    return hasLink || hasDuration || hasCallStatus;
  }
  
  return false;
}

function getCallType(note) {
  const noteType = String(note.note_type || '').toLowerCase();
  const params = note.params || {};
  
  if (noteType === 'call_out') return 'Исходящий';
  if (noteType === 'call_in') return 'Входящий';
  
  // Для common заметок определяем по содержимому
  if (noteType === 'common') {
    const text = String(note.params?.text || '').toLowerCase();
    if (text.includes('исходящий') || text.includes('outgoing')) return 'Исходящий';
    if (text.includes('входящий') || text.includes('incoming')) return 'Входящий';
  }
  
  return 'Неизвестно';
}

function processNoteEvent(eventData) {
  if (!['note_added', 'common_note_added', 'targeting_in_note_added', 'targeting_out_note_added'].includes(eventData.type)) {
    return false;
  }
  
  const noteId = eventData.value_after?.[0]?.note?.id;
  if (!noteId) {
    return false;
  }
  
  // Получаем полную заметку через API
  const fullNote = getAmoNoteById(noteId);
  if (!fullNote) {
    console.log(`❌ Не удалось получить заметку ${noteId}`);
    return false;
  }
  
  // Проверяем, является ли это звонком
  const isCall = isCallNote(fullNote);
  if (!isCall) {
    return false;
  }
  
  const callId = fullNote.id;
  if (isCallProcessed(callId)) {
    console.log(`⏭️ Звонок ${callId} уже обработан`);
    return false;
  }
  
  const leadId = eventData.entity_id;
  const lead = getAmoLead(leadId);
  if (!lead) {
    console.log(`❌ Не удалось получить сделку ${leadId}`);
    return false;
  }
  
  // Безопасное получение контакта
  let contact = null;
  if (lead._embedded && lead._embedded.contacts && lead._embedded.contacts.length > 0) {
    try {
      contact = getAmoContact(lead._embedded.contacts[0].id);
    } catch (error) {
      console.log(`⚠️ Ошибка получения контакта: ${error.message}`);
    }
  }
  
  // Безопасное получение компании
  let company = null;
  if (lead._embedded && lead._embedded.companies && lead._embedded.companies.length > 0) {
    try {
      company = getAmoCompany(lead._embedded.companies[0].id);
    } catch (error) {
      console.log(`⚠️ Ошибка получения компании: ${error.message}`);
    }
  }
  
  const callData = {
    callId: callId,
    leadId: leadId,
    leadName: lead.name || 'Без названия',
    contactName: contact ? contact.name : 'Без контакта',
    companyName: company ? company.name : 'Без компании',
    callType: getCallType(fullNote),
    callDate: formatTimestamp(eventData.created_at),
    duration: fullNote.params ? (fullNote.params.duration || 0) : 0,
    audioUrl: fullNote.params ? (fullNote.params.link || '') : '',
    callStatus: fullNote.params ? (fullNote.params.call_status || '') : '',
    responsible: lead.responsible_user_id || 'Неизвестно'
  };
  
  console.log(`✅ Обрабатываем звонок ${callId} для сделки ${leadId}`);
  appendToSheet(callData);
  return true;
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
  let pageCount = 0;
  const maxPages = 5;
  
  console.log(`🔄 Начало синхронизации событий с ${new Date(timeFrom * 1000).toLocaleString()}`);
  
  while (nextUrl && pageCount < maxPages) {
    try {
      pageCount++;
      console.log(`📄 Обрабатываем страницу ${pageCount}`);
      
      const response = amoRequest(nextUrl);
      if (!response || !response._embedded) {
        console.log(`⚠️ Нет данных на странице ${pageCount}`);
        break;
      }
      
      allEvents = allEvents.concat(response._embedded.events);
      console.log(`✅ Страница ${pageCount}: ${response._embedded.events.length} событий`);
      
      nextUrl = response._links.next ? response._links.next.href : null;
      Utilities.sleep(100);
      
    } catch (error) {
      console.log(`❌ Ошибка на странице ${pageCount}:`, error.message);
      break;
    }
  }
  
  console.log(`📊 Всего получено ${allEvents.length} событий за ${pageCount} страниц`);
  
  // Диагностика типов событий
  const eventTypes = {};
  allEvents.forEach(e => {
    eventTypes[e.type] = (eventTypes[e.type] || 0) + 1;
  });
  console.log(`📈 Типы событий:`, eventTypes);
  
  // Ищем события с заметками (common_note_added, note_added)
  const noteEvents = allEvents.filter(e => 
    e.type === 'note_added' || 
    e.type === 'common_note_added' ||
    e.type === 'targeting_in_note_added' ||
    e.type === 'targeting_out_note_added'
  );
  console.log(`📝 Найдено ${noteEvents.length} событий с заметками`);
  
  // Показываем примеры событий
  if (allEvents.length > 0) {
    console.log(`📄 Примеры событий (первые 5):`);
    allEvents.slice(0, 5).forEach((event, i) => {
      console.log(`  ${i+1}. ${event.type} - ${event.entity_type} - ${event.id}`);
    });
  }
  
  let processed = 0;
  noteEvents.forEach(event => {
    try {
      if (processNoteEvent(event)) {
        processed++;
      }
    } catch (error) {
      console.log(`❌ Ошибка обработки события ${event.id}:`, error.message);
    }
  });
  
  console.log(`✅ Обработано ${processed} звонков`);
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

function debugEventsDetailed() {
  getConfigFromSheet();
  
  // Учитываем часовой пояс - смотрим на 6 часов назад
  const timeFrom = Math.floor((new Date().getTime() - 6 * 60 * 60 * 1000) / 1000);
  const url = `https://${config.AMO_SUBDOMAIN}.amocrm.ru/api/v4/events?filter[created_at][from]=${timeFrom}&limit=50`;
  
  console.log('🔍 Детальная диагностика событий');
  console.log(`Время: ${new Date(timeFrom * 1000).toLocaleString()} (6 часов назад)`);
  
  const response = amoRequest(url);
  if (!response || !response._embedded) {
    console.log('❌ Нет событий');
    return;
  }
  
  const events = response._embedded.events;
  console.log(`📊 Всего событий: ${events.length}`);
  
  // Группируем по типам
  const eventTypes = {};
  events.forEach(e => {
    eventTypes[e.type] = (eventTypes[e.type] || 0) + 1;
  });
  
  console.log('📈 Типы событий:');
  Object.entries(eventTypes).forEach(([type, count]) => {
    console.log(`  ${type}: ${count}`);
  });
  
  // Показываем все события
  console.log('📄 Все события:');
  events.forEach((event, i) => {
    console.log(`  ${i+1}. ${event.type} - ${event.entity_type} - ${event.id} - ${new Date(event.created_at * 1000).toLocaleString()}`);
  });
  
  // Ищем события с заметками (все типы)
  const noteEvents = events.filter(e => 
    e.type === 'note_added' || 
    e.type === 'common_note_added' ||
    e.type === 'targeting_in_note_added' ||
    e.type === 'targeting_out_note_added'
  );
  console.log(`📝 События с заметками: ${noteEvents.length}`);
  
  if (noteEvents.length > 0) {
    console.log('📄 Примеры событий с заметками:');
    noteEvents.slice(0, 3).forEach((event, i) => {
      const note = event.value_after?.[0]?.note;
      console.log(`  ${i+1}. ID: ${event.id}, Note ID: ${note?.id}, Type: ${note?.note_type}`);
      console.log(`     Event type: ${event.type}`);
      console.log(`     Entity type: ${event.entity_type}`);
      console.log(`     Entity ID: ${event.entity_id}`);
      
      // Показываем полную структуру заметки
      console.log(`     Note structure:`, JSON.stringify(note, null, 2));
      
      if (note?.params) {
        console.log(`     Params: ${JSON.stringify(note.params)}`);
      }
      
      // Проверяем, есть ли текст заметки
      if (note?.text) {
        console.log(`     Text: ${String(note.text).slice(0, 200)}...`);
      }
    });
  }
}

function testConnection() {
  getConfigFromSheet();
  
  console.log('🔍 Тестирование подключения к amoCRM');
  console.log(`Поддомен: ${config.AMO_SUBDOMAIN}`);
  console.log(`Токен: ${config.AMO_TOKEN ? 'есть' : 'отсутствует'}`);
  
  try {
    const testUrl = `https://${config.AMO_SUBDOMAIN}.amocrm.ru/api/v4/account`;
    console.log(`Тестируем URL: ${testUrl}`);
    
    const response = amoRequest(testUrl);
    if (response) {
      console.log('✅ Подключение успешно');
      console.log(`Аккаунт: ${response.name || 'неизвестно'}`);
    } else {
      console.log('❌ Подключение не удалось');
    }
  } catch (error) {
    console.log('❌ Ошибка подключения:', error.message);
  }
}

function syncEventsLastHours(hours = 2) {
  getConfigFromSheet();
  
  const timeFrom = Math.floor((new Date().getTime() - hours * 60 * 60 * 1000) / 1000);
  const url = `https://${config.AMO_SUBDOMAIN}.amocrm.ru/api/v4/events?filter[created_at][from]=${timeFrom}&limit=250`;
  
  console.log(`🔄 Синхронизация за последние ${hours} часов`);
  console.log(`Время начала: ${new Date(timeFrom * 1000).toLocaleString()}`);
  
  let allEvents = [];
  let nextUrl = url;
  let pageCount = 0;
  const maxPages = 3;
  
  while (nextUrl && pageCount < maxPages) {
    try {
      pageCount++;
      console.log(`📄 Страница ${pageCount}`);
      
      const response = amoRequest(nextUrl);
      if (!response || !response._embedded) {
        console.log(`⚠️ Нет данных на странице ${pageCount}`);
        break;
      }
      
      allEvents = allEvents.concat(response._embedded.events);
      console.log(`✅ Получено ${response._embedded.events.length} событий`);
      
      nextUrl = response._links.next ? response._links.next.href : null;
      Utilities.sleep(200);
      
    } catch (error) {
      console.log(`❌ Ошибка на странице ${pageCount}:`, error.message);
      break;
    }
  }
  
  console.log(`📊 Всего событий: ${allEvents.length}`);
  
  const noteEvents = allEvents.filter(e => 
    e.type === 'note_added' || 
    e.type === 'common_note_added' ||
    e.type === 'targeting_in_note_added' ||
    e.type === 'targeting_out_note_added'
  );
  
  console.log(`📞 Событий с заметками: ${noteEvents.length}`);
  
  let processed = 0;
  noteEvents.forEach(event => {
    try {
      if (processNoteEvent(event)) {
        processed++;
      }
    } catch (error) {
      console.log(`❌ Ошибка обработки ${event.id}:`, error.message);
    }
  });
  
  console.log(`✅ Обработано звонков: ${processed}`);
}

function massExportCalls(days = 7) {
  getConfigFromSheet();
  
  const timeFrom = Math.floor((new Date().getTime() - days * 24 * 60 * 60 * 1000) / 1000);
  const url = `https://${config.AMO_SUBDOMAIN}.amocrm.ru/api/v4/events?filter[created_at][from]=${timeFrom}&limit=250`;
  
  console.log(`🚀 МАССОВАЯ ВЫГРУЗКА ЗВОНКОВ ЗА ${days} ДНЕЙ`);
  console.log(`Время начала: ${new Date(timeFrom * 1000).toLocaleString()}`);
  console.log(`Время окончания: ${new Date().toLocaleString()}`);
  
  let allEvents = [];
  let nextUrl = url;
  let pageCount = 0;
  const maxPages = 2; // Только 2 страницы = 500 событий
  const maxEvents = 500; // Лимит событий
  
  while (nextUrl && pageCount < maxPages && allEvents.length < maxEvents) {
    try {
      pageCount++;
      console.log(`📄 Обрабатываем страницу ${pageCount}`);
      
      const response = amoRequest(nextUrl);
      if (!response || !response._embedded) {
        console.log(`⚠️ Нет данных на странице ${pageCount}`);
        break;
      }
      
      allEvents = allEvents.concat(response._embedded.events);
      console.log(`✅ Страница ${pageCount}: ${response._embedded.events.length} событий`);
      
      // Проверяем лимит событий
      if (allEvents.length >= maxEvents) {
        console.log(`🛑 Достигнут лимит ${maxEvents} событий`);
        break;
      }
      
      nextUrl = response._links.next ? response._links.next.href : null;
      Utilities.sleep(200); // Быстрая задержка
      
    } catch (error) {
      console.log(`❌ Ошибка на странице ${pageCount}:`, error.message);
      break;
    }
  }
  
  console.log(`📊 ВСЕГО СОБРАНО ${allEvents.length} СОБЫТИЙ ЗА ${pageCount} СТРАНИЦ (лимит: ${maxEvents})`);
  
  // Диагностика типов событий
  const eventTypes = {};
  allEvents.forEach(e => {
    eventTypes[e.type] = (eventTypes[e.type] || 0) + 1;
  });
  console.log(`📈 ТИПЫ СОБЫТИЙ:`, eventTypes);
  
  // Ищем ВСЕ события с заметками
  const noteEvents = allEvents.filter(e => 
    e.type === 'note_added' || 
    e.type === 'common_note_added' ||
    e.type === 'targeting_in_note_added' ||
    e.type === 'targeting_out_note_added'
  );
  
  console.log(`📝 НАЙДЕНО ${noteEvents.length} СОБЫТИЙ С ЗАМЕТКАМИ`);
  
  // Показываем примеры событий с заметками
  if (noteEvents.length > 0) {
    console.log(`📄 ПРИМЕРЫ СОБЫТИЙ С ЗАМЕТКАМИ (первые 10):`);
    noteEvents.slice(0, 10).forEach((event, i) => {
      console.log(`  ${i+1}. ${event.type} - ${event.entity_type} - ${event.id} - ${new Date(event.created_at * 1000).toLocaleString()}`);
    });
  }
  
  let processed = 0;
  let errors = 0;
  
  console.log(`🔄 НАЧИНАЕМ ОБРАБОТКУ ${noteEvents.length} СОБЫТИЙ...`);
  
  noteEvents.forEach((event, index) => {
    try {
      console.log(`📞 Обрабатываем событие ${index + 1}/${noteEvents.length}: ${event.id}`);
      
      if (processNoteEvent(event)) {
        processed++;
        console.log(`✅ Звонок ${index + 1} обработан успешно`);
      } else {
        console.log(`⏭️ Событие ${index + 1} пропущено (не звонок или уже обработан)`);
      }
      
      // Пауза каждые 20 событий (реже)
      if ((index + 1) % 20 === 0) {
        console.log(`⏸️ Пауза после ${index + 1} событий...`);
        Utilities.sleep(500);
      }
      
    } catch (error) {
      errors++;
      console.log(`❌ Ошибка обработки события ${event.id}:`, error.message);
    }
  });
  
  console.log(`🎉 МАССОВАЯ ВЫГРУЗКА ЗАВЕРШЕНА!`);
  console.log(`📊 СТАТИСТИКА:`);
  console.log(`   - Всего событий: ${allEvents.length}`);
  console.log(`   - Событий с заметками: ${noteEvents.length}`);
  console.log(`   - Обработано звонков: ${processed}`);
  console.log(`   - Ошибок: ${errors}`);
  console.log(`   - Страниц обработано: ${pageCount}`);
  
  return {
    totalEvents: allEvents.length,
    noteEvents: noteEvents.length,
    processed: processed,
    errors: errors,
    pages: pageCount
  };
}

function quickExport() {
  console.log('🚀 БЫСТРАЯ ВЫГРУЗКА ЗВОНКОВ ЗА ПОСЛЕДНИЕ 3 ДНЯ');
  return massExportCalls(3);
}

function fullExport() {
  console.log('🚀 ПОЛНАЯ ВЫГРУЗКА ЗВОНКОВ ЗА ПОСЛЕДНИЕ 30 ДНЕЙ');
  return massExportCalls(30);
}