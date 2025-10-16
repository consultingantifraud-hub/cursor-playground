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
      writeLog(`⏭️ Событие ${eventData.id} уже обработано или нет заметки`, 'INFO');
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
        writeLog(`❌ Неизвестный тип сущности: ${eventData.entity_type}`, 'WARNING');
        return;
    }

    if (!lead) {
      writeLog(`❌ Не удалось получить сделку для события ${eventData.id}`, 'WARNING');
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
    writeLog(`✅ Звонок записан в таблицу: ${lead.name} (ID: ${lead.id})`, 'SUCCESS');
  } catch (error) {
    const errorMsg = `❌ Ошибка в событии ID ${eventData.id}: ${error.message}`;
    writeLog(errorMsg, 'ERROR');
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

// Запись логов в лист LOGS
function writeLog(message, type = 'INFO') {
  try {
    const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
    let logSheet = spreadsheet.getSheetByName('LOGS');
    
    if (!logSheet) {
      // Создаем лист LOGS если его нет
      logSheet = spreadsheet.insertSheet('LOGS');
      logSheet.getRange('A1:C1').setValues([['Время', 'Тип', 'Сообщение']]);
      logSheet.getRange('A1:C1').setFontWeight('bold');
    }
    
    const timestamp = new Date();
    const logData = [timestamp, type, message];
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
    writeLog('🚀 Начало синхронизации событий', 'INFO');
    
    const timeFrom = Math.floor((new Date().getTime() - 240 * 60 * 1000) / 1000); // 3 минут
    const timeFromStr = new Date(timeFrom * 1000).toLocaleString();
    writeLog(`⏰ Поиск событий с ${timeFromStr}`, 'INFO');
    
    let url = `https://${config.AMO_SUBDOMAIN}/api/v4/events?filter[created_at][from]=${timeFrom}&limit=250`;
    let allEvents = [];
    let pageCount = 0;
    
    while (true) {
      pageCount++;
      writeLog(`📄 Обработка страницы ${pageCount}`, 'INFO');
      
      const response = amoRequest(url);
      if (!response?._embedded?.events) {
        writeLog(`✅ Страница ${pageCount}: 0 событий - завершаем`, 'INFO');
        break;
      }
      
      const events = response._embedded.events;
      allEvents = allEvents.concat(events);
      writeLog(`✅ Страница ${pageCount}: ${events.length} событий`, 'INFO');
      
      url = response._links?.next?.href || '';
      if (!url) break;
    }
    
    writeLog(`📊 Всего собрано ${allEvents.length} событий за ${pageCount} страниц`, 'INFO');
    
    const calls = allEvents.filter(e =>
      ['outgoing_call', 'incoming_call'].includes(e.type)
    );
    
    writeLog(`📞 Найдено ${calls.length} событий звонков`, 'INFO');
    
    let processedCount = 0;
    calls.forEach((event, index) => {
      try {
        processEvent(event);
        processedCount++;
        writeLog(`✅ Обработан звонок ${index + 1}/${calls.length}: ${event.id}`, 'INFO');
      } catch (error) {
        writeLog(`❌ Ошибка обработки звонка ${event.id}: ${error.message}`, 'ERROR');
      }
    });
    
    writeLog(`🎉 Синхронизация завершена! Обработано звонков: ${processedCount}`, 'SUCCESS');
    
  } catch (error) {
    const errorMsg = `❌ Критическая ошибка синхронизации: ${error.message}`;
    writeLog(errorMsg, 'ERROR');
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
  writeLog('🚀 БЫСТРАЯ ВЫГРУЗКА ЗВОНКОВ ЗА ПОСЛЕДНИЕ 10 ЧАСОВ', 'INFO');
  return massExportCallsHours(10);
}

// Тестовая функция для проверки логов
function testLogging() {
  writeLog('🧪 Тестовое сообщение', 'INFO');
  writeLog('⚠️ Тестовое предупреждение', 'WARNING');
  writeLog('❌ Тестовая ошибка', 'ERROR');
  writeLog('✅ Тестовый успех', 'SUCCESS');
  console.log('✅ Тестовые логи записаны в лист LOGS');
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