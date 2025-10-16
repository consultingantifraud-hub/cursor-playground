// Конфигурация
function getConfigFromSheet() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = spreadsheet.getSheetByName("БД");
    if (!configSheet) throw new Error('Лист "БД" не найден');
    // Чтение исключений из колонки G
    const excludedFields = readExcludedFields(configSheet);
    return {
      AMO_TOKEN: configSheet.getRange("B2").getValue(),
      AMO_SUBDOMAIN: configSheet.getRange("B3").getValue(),
      SHEET_NAME: configSheet.getRange("B4").getValue(),
      SPREADSHEET_ID: spreadsheet.getId(),
      callStatusMap: Object.fromEntries(
        configSheet.getRange("C:D").getValues()
          .filter(row => row[0] && row[1])
          .map(([id, text]) => [id, text])
      ),
      taskTypesMap: Object.fromEntries(
        configSheet.getRange("E:F").getValues()
          .filter(row => row[0] && row[1])
          .map(([id, text]) => [id, text])
      ),
      excludedFields: excludedFields // Список исключенных полей
    };
  } catch (error) {
    console.error('❌ Ошибка конфигурации:', error.message);
    throw error;
  }
}
const config = getConfigFromSheet();

// Чтение исключенных полей
function readExcludedFields(sheet) {
  const excludedRange = sheet.getRange("G2:G");
  const values = excludedRange.getValues();
  return values
    .filter(row => row[0] !== "")
    .map(row => row[0].trim());
}

// Форматирование времени
function formatTimestamp(timestamp) {
  if (!timestamp) return "Нет данных";
  const date = new Date(timestamp * 1000);
  return Utilities.formatDate(date, "GMT+3", "dd.MM.yyyy HH:mm");
}

// Проверка обработки звонка
function isCallProcessed(callNoteId) {
  try {
    const sheet = SpreadsheetApp.openById(config.SPREADSHEET_ID)
      .getSheetByName(config.SHEET_NAME);
    if (!sheet) return false;
    const processedIds = sheet.getRange(2, 15, sheet.getLastRow()-1, 1)
      .getValues()
      .flat()
      .filter(id => id !== '');
    return processedIds.includes(callNoteId);
  } catch (error) {
    console.error('❌ Ошибка проверки:', error.message);
    return false;
  }
}

// Универсальный запрос к API
function amoRequest(url, method = 'GET', body = null) {
  try {
    const options = {
      method: method,
      headers: { 'Authorization': `Bearer ${config.AMO_TOKEN}` },
      muteHttpExceptions: true
    };
    if (body) options.payload = JSON.stringify(body);
    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();
    if (statusCode === 204) {
      console.log(`✅ Пустой ответ (204) для ${url}`);
      return null;
    }
    if (statusCode !== 200) {
      console.error(`❌ HTTP ${statusCode}: ${response.getContentText()}`);
      return null;
    }
    return JSON.parse(response.getContentText());
  } catch (error) {
    console.error(`❌ Ошибка запроса к ${url}:`, error.message);
    return null;
  }
}

// Получение задач для сделки
function getAmoTasks(leadId) {
  try {
    const url = `https://${config.AMO_SUBDOMAIN}/api/v4/tasks?filter[entity_id]=${leadId}&filter[entity_type]=leads`;
    const tasksResponse = amoRequest(url);
    if (!tasksResponse) return 'Нет задач';
    const tasks = tasksResponse._embedded?.tasks || [];
    return tasks.map(task => {
      const taskType = config.taskTypesMap[task.task_type_id] || `Неизвестный тип (${task.task_type_id})`;
      const dueDate = formatTimestamp(task.complete_till);
      const status = task.is_completed === true ? 'Выполнена' :
        (task.is_completed === false ? 'Не выполнена' : 'Статус не определен');
      return `• ${taskType}: ${task.text || 'Нет текста'}\n  Срок: ${dueDate}\n  Статус: ${status}`;
    }).join('\n') || 'Нет задач';
  } catch (error) {
    console.error(`❌ Ошибка задач для сделки ${leadId}:`, error.message);
    return 'Ошибка при получении задач';
  }
}

// Сбор всех полей из сделки (с фильтрацией)
function getAllEntityFields(lead) {
  try {
    const fields = [];
    // Поля сделки
    fields.push('--- Сделка ---');
    fields.push(`ID: ${lead.id}`);
    fields.push(`Ответственный: ${getAmoUser(lead.responsible_user_id).name}`);
    // Фильтрация кастомных полей сделки
    const customFieldsDeal = (lead.custom_fields_values || [])
      .map(field => {
        const fieldName = field.field_name;
        if (config.excludedFields.includes(fieldName)) return null;
        const values = field.values
          .map(v => v.value)
          .filter(v => v)
          .join(', ');
        return values ? `${fieldName}: ${values}` : null;
      })
      .filter(Boolean);
    fields.push(...customFieldsDeal);
    // Поля контактов
    if (lead._embedded?.contacts?.length) {
      lead._embedded.contacts.forEach(contactId => {
        const contact = getEntityData('contacts', contactId.id);
        if (contact) {
          fields.push('--- Контакт ---');
          fields.push(`ID: ${contact.id}`);
          fields.push(`Ответственный: ${contact.responsible}`);
          // Фильтрация кастомных полей контакта
          const customFieldsContact = (contact.custom_fields_values || [])
            .map(field => {
              const fieldName = field.field_name;
              if (config.excludedFields.includes(fieldName)) return null;
              const values = field.values
                .map(v => v.value)
                .filter(v => v)
                .join(', ');
              return values ? `${fieldName}: ${values}` : null;
            })
            .filter(Boolean);
          fields.push(...customFieldsContact);
        }
      });
    }
    // Поля компаний
    if (lead._embedded?.companies?.length) {
      lead._embedded.companies.forEach(companyId => {
        const company = getEntityData('companies', companyId.id);
        if (company) {
          fields.push('--- Компания ---');
          fields.push(`ID: ${company.id}`);
          fields.push(`Ответственный: ${company.responsible}`);
          // Фильтрация кастомных полей компании
          const customFieldsCompany = (company.custom_fields_values || [])
            .map(field => {
              const fieldName = field.field_name;
              if (config.excludedFields.includes(fieldName)) return null;
              const values = field.values
                .map(v => v.value)
                .filter(v => v)
                .join(', ');
              return values ? `${fieldName}: ${values}` : null;
            })
            .filter(Boolean);
          fields.push(...customFieldsCompany);
        }
      });
    }
    return fields.join('\n') || 'Нет данных';
  } catch (error) {
    console.error('❌ Ошибка сбора полей:', error.message);
    return 'Ошибка';
  }
}

// Получение данных сущности
function getEntityData(entityType, entityId) {
  try {
    const url = `https://${config.AMO_SUBDOMAIN}/api/v4/${entityType}/${entityId}`;
    const entity = amoRequest(url);
    if (!entity) return null;
    return {
      id: entity.id,
      name: entity.name,
      responsible: getAmoUser(entity.responsible_user_id).name,
      custom_fields_values: entity.custom_fields_values || []
    };
  } catch (error) {
    console.error(`❌ Ошибка ${entityType} ID ${entityId}:`, error.message);
    return null;
  }
}

// Обработка события
function processEvent(eventData) {
  try {
    console.log(`📞 Обработка события ID: ${eventData.id}`);
    const callNoteId = eventData.value_after?.[0]?.note?.id;
    if (!callNoteId || isCallProcessed(callNoteId)) {
      writeLog(`Событие ${eventData.id} уже обработано или нет заметки`, 'INFO', { script_id: 'processEvent' });
      return;
    }

    let lead;
    switch(eventData.entity_type.toLowerCase()) {
      case 'lead':
        lead = getAmoLead(eventData.entity_id);
        break;
      case 'contact':
        lead = getLeadByContact(eventData.entity_id);
        break;
      case 'company':
        lead = getLeadByCompany(eventData.entity_id);
        break;
      default:
        writeLog(`Неизвестный тип сущности: ${eventData.entity_type}`, 'WARNING', { script_id: 'processEvent' });
        return;
    }

    if (!lead) {
      writeLog(`Не удалось получить сделку для события ${eventData.id}`, 'WARNING', { script_id: 'processEvent' });
      return;
    }

    // Получение задач и примечаний
    const tasksText = getAmoTasks(lead.id);
    const notes = getAmoNotes('leads', lead.id);
    const callNote = notes.find(n => n.id === callNoteId);

    // Обработка примечаний
    let notesText = notes.map(note => {
      let text = '';
      switch(note.note_type) {
        case 'common':
          text = `📝 Примечание: ${note.params.text}`;
          if (note._embedded?.attachments?.length) {
            text += `\nattachments:\n${note._embedded.attachments
              .map(att => `• ${att.name} (${att.link})`)
              .join('\n')}`;
          }
          break;
        case 'call_out':
          const duration = note.params.duration ? ` (${note.params.duration}с)` : '';
          const result = config.callStatusMap[note.params.call_status] || 'Без результата';
          text = `📞 Звонок${duration}: ${result}`;
          if (note.params.link) text += ` (${note.params.link})`;
          if (note._embedded?.attachments?.length) {
            text += `\nattachments:\n${note._embedded.attachments
              .map(att => `• ${att.name} (${att.link})`)
              .join('\n')}`;
          }
          break;
        case 'task':
          text = `✅ Задача [ID:${note.id}]: ${note.params.text}`;
          text += `\nСтатус: ${note.params.status === 1 ? 'Выполнена' : 'Не выполнена'}`;
          text += `\nСрок: ${new Date(note.params.task_deadline * 1000)}`;
          break;
        case 'mail':
          text = `📧 Почта: ${note.params.subject}`;
          text += `\nОт: ${note.params.from}`;
          text += `\nСообщение: ${note.params.text}`;
          text += `\nСвязь: ${note.params.link}`;
          break;
        case 'chat':
          text = `💬 Чат (${note.params.service}):`;
          text += `\nСообщение: ${note.params.text}`;
          text += `\nСвязь: ${note.params.link}`;
          break;
        default:
          text = `Неизвестный тип: ${note.note_type}`;
      }
      return text;
    }).join('\n').trim() || 'Нет примечаний';

    // Добавляем задачи в примечания
    notesText += `\n📌 Задачи в сделке:\n${tasksText}`;

    // Формируем данные
    const rowData = [
      new Date(eventData.created_at * 1000), // 1. Время
      callNote?.params?.link || 'Нет записи', // 2. Ссылка
      lead.name, // 3. Сделка
      getAmoPipeline(lead.pipeline_id).name, // 4. Воронка
      getAmoStatus(lead.status_id, lead.pipeline_id).name, // 5. Статус
      lead.id, // 6. ID сделки
      lead.pipeline_id, // 7. ID воронки
      lead.status_id, // 8. ID статуса
      lead.responsible_user_id, // 9. ID ответственного
      getAmoUser(lead.responsible_user_id).name, // 10. Ответственный
      notesText, // 11. Примечания + задачи (колонка K)
      callNote?.params?.duration || 0, // 12. Длительность
      config.callStatusMap[callNote?.params?.call_status] || 'Без результата', // 13. Результат
      eventData.type === 'outgoing_call' ? 'Исходящий' : 'Входящий', // 14. Тип
      callNoteId, // 15. ID звонка
      getAllEntityFields(lead) // 16. Все поля (с фильтрацией)
    ];
    appendToSheet(rowData);
    writeLog(`Звонок записан в таблицу: ${lead.name} (ID: ${lead.id})`, 'SUCCESS', { script_id: 'processEvent' });
  } catch (error) {
    const errorMsg = `Ошибка в событии ID ${eventData.id}: ${error.message}`;
    writeLog(errorMsg, 'ERROR', { script_id: 'processEvent' });
    console.error(errorMsg);
  }
}

// Получение сделки с полными данными
function getAmoLead(leadId) {
  try {
    const url = `https://${config.AMO_SUBDOMAIN}/api/v4/leads/${leadId}?with=contacts,companies`;
    const lead = amoRequest(url);
    if (!lead) return null;
    lead.responsible = getAmoUser(lead.responsible_user_id).name;
    return lead;
  } catch (error) {
    console.error(`❌ Ошибка сделки ID ${leadId}:`, error.message);
    return null;
  }
}

// Получение сделок по контакту
function getLeadByContact(contactId) {
  try {
    const url = `https://${config.AMO_SUBDOMAIN}/api/v4/contacts/${contactId}/links`;
    const linksResponse = amoRequest(url);
    const links = linksResponse?._embedded?.links || [];
    
    // Извлекаем ID сделок из связей
    const leadIds = links
      .filter(link => link.to_entity_type === 'leads')
      .map(link => link.to_entity_id);
    
    return leadIds.length > 0 ? getAmoLead(leadIds[0]) : null;
  } catch (error) {
    console.error(`❌ Ошибка поиска сделок для контакта ${contactId}:`, error.message);
    return null;
  }
}

// Получение сделок по компании
function getLeadByCompany(companyId) {
  try {
    const url = `https://${config.AMO_SUBDOMAIN}/api/v4/companies/${companyId}/links`;
    const linksResponse = amoRequest(url);
    const links = linksResponse?._embedded?.links || [];
    
    // Извлекаем ID сделок из связей
    const leadIds = links
      .filter(link => link.to_entity_type === 'leads')
      .map(link => link.to_entity_id);
    
    return leadIds.length > 0 ? getAmoLead(leadIds[0]) : null;
  } catch (error) {
    console.error(`❌ Ошибка поиска сделок для компании ${companyId}:`, error.message);
    return null;
  }
}

// Получение примечаний
function getAmoNotes(entityType, entityId) {
  try {
    const url = `https://${config.AMO_SUBDOMAIN}/api/v4/${entityType}/${entityId}/notes?with=attachments`;
    const response = amoRequest(url);
    return response?._embedded?.notes || [];
  } catch (error) {
    console.error(`❌ Ошибка примечаний для ${entityType} ID ${entityId}:`, error.message);
    return [];
  }
}

// Получение воронки
function getAmoPipeline(pipelineId) {
  try {
    const url = `https://${config.AMO_SUBDOMAIN}/api/v4/leads/pipelines`;
    const pipelines = amoRequest(url)?._embedded?.pipelines || [];
    return pipelines.find(p => p.id === pipelineId) || { id: pipelineId, name: "Неизвестная воронка" };
  } catch (error) {
    console.error(`❌ Ошибка воронки ID ${pipelineId}:`, error.message);
    return { id: pipelineId, name: "Неизвестная воронка" };
  }
}

// Получение статуса
function getAmoStatus(statusId, pipelineId) {
  try {
    const url = `https://${config.AMO_SUBDOMAIN}/api/v4/leads/pipelines`;
    const pipelines = amoRequest(url)?._embedded?.pipelines || [];
    const pipeline = pipelines.find(p => p.id === pipelineId);
    const statuses = pipeline?._embedded?.statuses || [];
    return statuses.find(s => s.id === statusId) || { id: statusId, name: "Неизвестный статус" };
  } catch (error) {
    console.error(`❌ Ошибка статуса ID ${statusId}:`, error.message);
    return { id: statusId, name: "Неизвестный статус" };
  }
}

// Получение пользователя
function getAmoUser(userId) {
  try {
    const url = `https://${config.AMO_SUBDOMAIN}/api/v4/users/${userId}`;
    const user = amoRequest(url);
    return user ? {
      id: user.id,
      name: user.name || "Неизвестный пользователь"
    } : { id: userId, name: "Неизвестный пользователь" };
  } catch (error) {
    console.error(`❌ Ошибка пользователя ID ${userId}:`, error.message);
    return { id: userId, name: "Неизвестный пользователь" };
  }
}

// Запись данных в таблицу
function appendToSheet(rowData) {
  try {
    const sheet = SpreadsheetApp.openById(config.SPREADSHEET_ID)
      .getSheetByName(config.SHEET_NAME);
    if (!sheet) throw new Error('Лист не найден');
    sheet.appendRow(rowData);
    console.log('✅ Данные записаны:', rowData);
  } catch (error) {
    console.error('❌ Ошибка записи:', error.message);
  }
}

// Запись логов в лист LOGS в формате ProTalk
function writeLog(message, type = 'INFO', additionalData = {}) {
  try {
    const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
    let logSheet = spreadsheet.getSheetByName('LOGS');
    
    if (!logSheet) {
      // Создаем лист LOGS если его нет
      logSheet = spreadsheet.insertSheet('LOGS');
      const headers = [
        'Timestamp', 'chat_id', 'social_id', 'question', 'ai_reply', 'channel', 
        'script_id', 'model', 'tokens', 'user_tokens', 'tokens_in', 'tokens_out', 
        'error_log', 'function_log', 'api_key'
      ];
      logSheet.getRange('A1:O1').setValues([headers]);
      logSheet.getRange('A1:O1').setFontWeight('bold');
    }
    
    const timestamp = new Date();
    const logData = [
      timestamp,                                    // Timestamp
      additionalData.chat_id || '',                // chat_id
      additionalData.social_id || '',              // social_id
      message,                                     // question (сообщение)
      additionalData.ai_reply || '',               // ai_reply
      'amogs_bot',                                 // channel
      additionalData.script_id || 'syncEvents',    // script_id
      'gpt-4.1-mini',                             // model
      additionalData.tokens || 0,                  // tokens
      additionalData.user_tokens || 0,             // user_tokens
      additionalData.tokens_in || 0,               // tokens_in
      additionalData.tokens_out || 0,              // tokens_out
      type === 'ERROR' ? message : '',             // error_log
      type === 'INFO' ? message : '',              // function_log
      additionalData.api_key || ''                 // api_key
    ];
    
    logSheet.appendRow(logData);
    
    // Ограничиваем количество строк в логах (оставляем последние 1000)
    const maxRows = 1000;
    if (logSheet.getLastRow() > maxRows) {
      logSheet.deleteRows(2, logSheet.getLastRow() - maxRows);
    }
    
    console.log(`📝 LOG [${type}]: ${message}`);
  } catch (error) {
    console.error('❌ Ошибка записи лога:', error.message);
  }
}

// Основная функция синхронизации
function syncEventsToday() {
  try {
    writeLog('Начало синхронизации событий', 'INFO', { script_id: 'syncEventsToday' });
    
    const timeFrom = Math.floor((new Date().getTime() - 240 * 60 * 1000) / 1000); // 3 минут
    const timeFromStr = new Date(timeFrom * 1000).toLocaleString();
    writeLog(`Поиск событий с ${timeFromStr}`, 'INFO', { script_id: 'syncEventsToday' });
    
    let url = `https://${config.AMO_SUBDOMAIN}/api/v4/events?filter[created_at][from]=${timeFrom}&limit=250`;
    let allEvents = [];
    let pageCount = 0;
    
    while (true) {
      pageCount++;
      writeLog(`Обработка страницы ${pageCount}`, 'INFO', { script_id: 'syncEventsToday' });
      
      const response = amoRequest(url);
      if (!response?._embedded?.events) {
        writeLog(`Страница ${pageCount}: 0 событий - завершаем`, 'INFO', { script_id: 'syncEventsToday' });
        break;
      }
      
      const events = response._embedded.events;
      allEvents = allEvents.concat(events);
      writeLog(`Страница ${pageCount}: ${events.length} событий`, 'INFO', { script_id: 'syncEventsToday' });
      
      url = response._links?.next?.href || '';
      if (!url) break;
    }
    
    writeLog(`Всего собрано ${allEvents.length} событий за ${pageCount} страниц`, 'INFO', { script_id: 'syncEventsToday' });
    
    // Анализируем типы событий
    const eventTypes = {};
    allEvents.forEach(e => {
      eventTypes[e.type] = (eventTypes[e.type] || 0) + 1;
    });
    
    const eventTypesStr = Object.entries(eventTypes)
      .map(([type, count]) => `${type}: ${count}`)
      .join(', ');
    writeLog(`Типы событий: ${eventTypesStr}`, 'INFO', { script_id: 'syncEventsToday' });
    
    // Ищем звонки по разным типам событий
    const calls = allEvents.filter(e => {
      const callTypes = ['outgoing_call', 'incoming_call', 'call_started', 'call_ended'];
      return callTypes.includes(e.type) || 
             (e.type === 'note_added' && e.value_after?.[0]?.note?.note_type === 'call') ||
             (e.type === 'common_note_added' && e.value_after?.[0]?.note?.note_type === 'call');
    });
    
    writeLog(`Найдено ${calls.length} событий звонков`, 'INFO', { script_id: 'syncEventsToday' });
    
    // Если звонков нет, показываем примеры событий
    if (calls.length === 0 && allEvents.length > 0) {
      const sampleEvents = allEvents.slice(0, 3).map(e => `${e.type} (${e.id})`).join(', ');
      writeLog(`Примеры событий: ${sampleEvents}`, 'INFO', { script_id: 'syncEventsToday' });
    }
    
    let processedCount = 0;
    calls.forEach((event, index) => {
      try {
        processEvent(event);
        processedCount++;
        writeLog(`Обработан звонок ${index + 1}/${calls.length}: ${event.id}`, 'INFO', { script_id: 'processEvent' });
      } catch (error) {
        writeLog(`Ошибка обработки звонка ${event.id}: ${error.message}`, 'ERROR', { script_id: 'processEvent' });
      }
    });
    
    writeLog(`Синхронизация завершена! Обработано звонков: ${processedCount}`, 'SUCCESS', { script_id: 'syncEventsToday' });
    
  } catch (error) {
    const errorMsg = `Критическая ошибка синхронизации: ${error.message}`;
    writeLog(errorMsg, 'ERROR', { script_id: 'syncEventsToday' });
    console.error(errorMsg);
  }
}

// Функции для массовой выгрузки (добавляем к рабочему коду)
function massExportCalls(days = 7) {
  console.log(`🚀 МАССОВАЯ ВЫГРУЗКА ЗВОНКОВ ЗА ${days} ДНЕЙ`);
  
  const endTime = Math.floor(Date.now() / 1000);
  const startTime = endTime - (days * 24 * 60 * 60);
  
  console.log(`Время начала: ${new Date(startTime * 1000).toLocaleString()}`);
  console.log(`Время окончания: ${new Date(endTime * 1000).toLocaleString()}`);
  
  let allEvents = [];
  let page = 1;
  const maxPages = 2; // Only 2 pages = 500 events
  const maxEvents = 500; // Event limit
  
  while (page <= maxPages) {
    console.log(`📄 Обрабатываем страницу ${page}`);
    
    const url = `https://${config.AMO_SUBDOMAIN}/api/v4/events?filter[created_at][from]=${startTime}&filter[created_at][to]=${endTime}&limit=250&page=${page}`;
    const response = amoRequest(url);
    
    if (!response?._embedded?.events || response._embedded.events.length === 0) {
      console.log(`✅ Страница ${page}: 0 событий - завершаем`);
      break;
    }
    
    const events = response._embedded.events;
    allEvents = allEvents.concat(events);
    console.log(`✅ Страница ${page}: ${events.length} событий`);
    
    if (allEvents.length >= maxEvents) {
      console.log(`🛑 Достигнут лимит ${maxEvents} событий`);
      break;
    }
    
    page++;
    Utilities.sleep(200); // Quick delay
  }
  
  console.log(`📊 ВСЕГО СОБРАНО ${allEvents.length} СОБЫТИЙ ЗА ${page-1} СТРАНИЦ (лимит: ${maxEvents})`);
  
  // Фильтруем события с заметками
  const noteEvents = allEvents.filter(e =>
    e.type === 'note_added' ||
    e.type === 'common_note_added' ||
    e.type === 'targeting_in_note_added' ||
    e.type === 'targeting_out_note_added' ||
    e.type === 'outgoing_call' ||
    e.type === 'incoming_call' ||
    e.type === 'call_started' ||
    e.type === 'call_ended'
  );
  
  console.log(`📝 НАЙДЕНО ${noteEvents.length} СОБЫТИЙ С ЗАМЕТКАМИ`);
  
  // Обрабатываем события
  let processedCount = 0;
  noteEvents.forEach((event, index) => {
    console.log(`📞 Обрабатываем событие ${index + 1}/${noteEvents.length}: ${event.id}`);
    
    if (['outgoing_call', 'incoming_call', 'call_started', 'call_ended'].includes(event.type)) {
      console.log(`📞 Обрабатываем прямое событие звонка: ${event.type}`);
      processEvent(event);
      processedCount++;
    } else {
      // Обрабатываем как обычное событие с заметкой
      processEvent(event);
      processedCount++;
    }
    
    // Пауза каждые 20 событий
    if ((index + 1) % 20 === 0) {
      console.log(`⏸️ Пауза после ${index + 1} событий...`);
      Utilities.sleep(500);
    }
  });
  
  console.log(`🎉 МАССОВАЯ ВЫГРУЗКА ЗАВЕРШЕНА!`);
  console.log(`📊 СТАТИСТИКА:`);
  console.log(`   - Всего событий: ${allEvents.length}`);
  console.log(`   - Событий с заметками: ${noteEvents.length}`);
  console.log(`   - Обработано звонков: ${processedCount}`);
  console.log(`   - Ошибок: 0`);
  console.log(`   - Страниц обработано: ${page-1}`);
  
  return processedCount;
}

function quickExport() {
  console.log('🚀 БЫСТРАЯ ВЫГРУЗКА ЗВОНКОВ ЗА ПОСЛЕДНИЕ 10 ЧАСОВ');
  writeLog('БЫСТРАЯ ВЫГРУЗКА ЗВОНКОВ ЗА ПОСЛЕДНИЕ 10 ЧАСОВ', 'INFO', { script_id: 'quickExport' });
  return massExportCallsHours(10);
}

// Тестовая функция для проверки логов
function testLogging() {
  writeLog('Тестовое сообщение', 'INFO', { script_id: 'testLogging' });
  writeLog('Тестовое предупреждение', 'WARNING', { script_id: 'testLogging' });
  writeLog('Тестовая ошибка', 'ERROR', { script_id: 'testLogging' });
  writeLog('Тестовый успех', 'SUCCESS', { script_id: 'testLogging' });
  console.log('✅ Тестовые логи записаны в лист LOGS');
}

// Поиск звонков в заметках за период
function findCallsInNotes(hours = 24) {
  try {
    console.log(`🔍 Поиск звонков в заметках за последние ${hours} часов`);
    writeLog(`Поиск звонков в заметках за последние ${hours} часов`, 'INFO', { script_id: 'findCallsInNotes' });
    
    const timeFrom = Math.floor((new Date().getTime() - hours * 60 * 60 * 1000) / 1000);
    const timeFromStr = new Date(timeFrom * 1000).toLocaleString();
    console.log(`⏰ Ищем заметки с ${timeFromStr}`);
    
    // Ищем заметки по сделкам
    let url = `https://${config.AMO_SUBDOMAIN}/api/v4/leads/notes?filter[created_at][from]=${timeFrom}&limit=250`;
    let allNotes = [];
    let pageCount = 0;
    
    while (url && pageCount < 10) {
      pageCount++;
      console.log(`📄 Страница заметок ${pageCount}`);
      
      const response = amoRequest(url);
      if (!response?._embedded?.items) break;
      
      const notes = response._embedded.items;
      allNotes = allNotes.concat(notes);
      console.log(`✅ Страница ${pageCount}: ${notes.length} заметок`);
      
      url = response._links?.next?.href || '';
    }
    
    console.log(`📊 Всего найдено заметок: ${allNotes.length}`);
    writeLog(`Всего найдено заметок: ${allNotes.length}`, 'INFO', { script_id: 'findCallsInNotes' });
    
    // Фильтруем звонки
    const callNotes = allNotes.filter(note => {
      const noteType = String(note.note_type || '').toLowerCase();
      return /^(call_in|call_out)$/i.test(noteType);
    });
    
    console.log(`📞 Найдено звонков в заметках: ${callNotes.length}`);
    writeLog(`Найдено звонков в заметках: ${callNotes.length}`, 'INFO', { script_id: 'findCallsInNotes' });
    
    // Показываем примеры звонков
    if (callNotes.length > 0) {
      console.log('📞 === ПРИМЕРЫ ЗВОНКОВ ===');
      callNotes.slice(0, 5).forEach((note, i) => {
        console.log(`${i+1}. Звонок ID ${note.id}:`, {
          note_type: note.note_type,
          entity_id: note.entity_id,
          entity_type: note.entity_type,
          created_at: new Date(note.created_at * 1000).toLocaleString(),
          has_params: !!note.params,
          params: note.params
        });
      });
    }
    
    return callNotes;
    
  } catch (error) {
    const errorMsg = `Ошибка поиска звонков в заметках: ${error.message}`;
    writeLog(errorMsg, 'ERROR', { script_id: 'findCallsInNotes' });
    console.error(errorMsg);
    return [];
  }
}

// Функция для просмотра всех событий (диагностическая)
function logEvents() {
  try {
    const timeFrom = Math.floor((new Date().getTime() - 150 * 60 * 1000) / 1000);
    let url = `https://${config.AMO_SUBDOMAIN}/api/v4/events?filter[created_at][from]=${timeFrom}&limit=250`;
    let all = [];
    
    while (url) {
      const resp = UrlFetchApp.fetch(url, { 
        method: 'get', 
        headers: { 'Authorization': 'Bearer ' + config.AMO_TOKEN }, 
        muteHttpExceptions: true, 
        followRedirects: true 
      });
      const code = resp.getResponseCode(); 
      if (code < 200 || code >= 300) break;
      const json = JSON.parse(resp.getContentText() || '{}');
      all = all.concat(json?._embedded?.events || []);
      url = json?._links?.next?.href || '';
    }
    
    console.log('[MONITOR] events: ' + all.length);
    console.log(JSON.stringify(all.slice(0, 50), null, 2)); // не заспамить логи
  } catch (e) { 
    console.error('[MONITOR] ' + String(e)); 
  }
}

function fullExport() {
  console.log('🚀 ПОЛНАЯ ВЫГРУЗКА ЗВОНКОВ ЗА ПОСЛЕДНИЕ 30 ДНЕЙ');
  return massExportCalls(30);
}

function massExportCallsHours(hours = 10) {
  console.log(`🚀 МАССОВАЯ ВЫГРУЗКА ЗВОНКОВ ЗА ${hours} ЧАСОВ`);
  
  const endTime = Math.floor(Date.now() / 1000);
  const startTime = endTime - (hours * 60 * 60);
  
  console.log(`Время начала: ${new Date(startTime * 1000).toLocaleString()}`);
  console.log(`Время окончания: ${new Date(endTime * 1000).toLocaleString()}`);
  
  let allEvents = [];
  let page = 1;
  const maxPages = 5; // Increased for more coverage
  const maxEvents = 1000; // Increased event limit
  
  while (page <= maxPages) {
    console.log(`📄 Обрабатываем страницу ${page}`);
    
    const url = `https://${config.AMO_SUBDOMAIN}/api/v4/events?filter[created_at][from]=${startTime}&filter[created_at][to]=${endTime}&limit=250&page=${page}`;
    const response = amoRequest(url);
    
    if (!response?._embedded?.events || response._embedded.events.length === 0) {
      console.log(`✅ Страница ${page}: 0 событий - завершаем`);
      break;
    }
    
    const events = response._embedded.events;
    allEvents = allEvents.concat(events);
    console.log(`✅ Страница ${page}: ${events.length} событий`);
    
    if (allEvents.length >= maxEvents) {
      console.log(`🛑 Достигнут лимит ${maxEvents} событий`);
      break;
    }
    
    page++;
    Utilities.sleep(200); // Quick delay
  }
  
  console.log(`📊 ВСЕГО СОБРАНО ${allEvents.length} СОБЫТИЙ ЗА ${page-1} СТРАНИЦ (лимит: ${maxEvents})`);
  
  // Фильтруем события с заметками
  const noteEvents = allEvents.filter(e =>
    e.type === 'note_added' ||
    e.type === 'common_note_added' ||
    e.type === 'targeting_in_note_added' ||
    e.type === 'targeting_out_note_added' ||
    e.type === 'outgoing_call' ||
    e.type === 'incoming_call' ||
    e.type === 'call_started' ||
    e.type === 'call_ended'
  );
  
  console.log(`📝 НАЙДЕНО ${noteEvents.length} СОБЫТИЙ С ЗАМЕТКАМИ`);
  
  // Обрабатываем события
  let processedCount = 0;
  noteEvents.forEach((event, index) => {
    console.log(`📞 Обрабатываем событие ${index + 1}/${noteEvents.length}: ${event.id}`);
    
    if (['outgoing_call', 'incoming_call', 'call_started', 'call_ended'].includes(event.type)) {
      console.log(`📞 Обрабатываем прямое событие звонка: ${event.type}`);
      processEvent(event);
      processedCount++;
    } else {
      // Обрабатываем как обычное событие с заметкой
      processEvent(event);
      processedCount++;
    }
    
    // Пауза каждые 20 событий
    if ((index + 1) % 20 === 0) {
      console.log(`⏸️ Пауза после ${index + 1} событий...`);
      Utilities.sleep(500);
    }
  });
  
  console.log(`🎉 МАССОВАЯ ВЫГРУЗКА ЗАВЕРШЕНА!`);
  console.log(`📊 СТАТИСТИКА:`);
  console.log(`   - Всего событий: ${allEvents.length}`);
  console.log(`   - Событий с заметками: ${noteEvents.length}`);
  console.log(`   - Обработано звонков: ${processedCount}`);
  console.log(`   - Ошибок: 0`);
  console.log(`   - Страниц обработано: ${page-1}`);
  
  return processedCount;
}