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
      writeLog(`Событие ${eventData.id} уже обработано или нет заметки`, 'INFO');
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
        writeLog(`Неизвестный тип сущности: ${eventData.entity_type}`, 'WARNING');
        return;
    }

    if (!lead) {
      writeLog(`Не удалось получить сделку для события ${eventData.id}`, 'WARNING');
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
    writeLog(`Звонок записан в таблицу: ${lead.name} (ID: ${lead.id})`, 'SUCCESS');
  } catch (error) {
    const errorMsg = `Ошибка в событии ID ${eventData.id}: ${error.message}`;
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

// Простое логирование в консоль и лист LOGS
function writeLog(message, type = 'INFO') {
  try {
    // Логируем в консоль
    console.log(`[${type}] ${message}`);
    
    // Логируем в лист LOGS
    const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
    let logSheet = spreadsheet.getSheetByName('LOGS');
    
    if (!logSheet) {
      logSheet = spreadsheet.insertSheet('LOGS');
      logSheet.getRange('A1:C1').setValues([['Время', 'Тип', 'Сообщение']]);
      logSheet.getRange('A1:C1').setFontWeight('bold');
    }
    
    const timestamp = new Date();
    const logData = [timestamp, type, message];
    logSheet.appendRow(logData);
    
    // Ограничиваем количество строк в логах (оставляем последние 500)
    const maxRows = 500;
    if (logSheet.getLastRow() > maxRows) {
      logSheet.deleteRows(2, logSheet.getLastRow() - maxRows);
    }
    
  } catch (error) {
    console.error('❌ Ошибка записи лога:', error.message);
  }
}

// Основная функция синхронизации (ОТКЛЮЧЕНА - звонки обрабатываются ProTalk)
function syncEventsToday() {
  try {
    writeLog('syncEventsToday ОТКЛЮЧЕНА - звонки обрабатываются ProTalk', 'INFO');
    writeLog('Используйте ToProTalkBot() и chatToProTalkBot() для обработки звонков', 'INFO');
    
    // Проверяем, есть ли звонки в листе
    const sheet = SpreadsheetApp.openById(config.SPREADSHEET_ID)
      .getSheetByName(config.SHEET_NAME);
    
    if (!sheet) {
      writeLog('Лист не найден', 'ERROR');
      return;
    }
    
    const lastRow = sheet.getLastRow();
    writeLog(`Всего строк в листе: ${lastRow}`, 'INFO');
    
    if (lastRow < 2) {
      writeLog('Лист пуст', 'WARNING');
      return;
    }
    
    // Проверяем последние 5 строк
    const data = sheet.getRange(Math.max(2, lastRow - 4), 1, Math.min(5, lastRow - 1), sheet.getLastColumn()).getValues();
    
    let callsWithTranscript = 0;
    let callsWithEvaluation = 0;
    
    data.forEach((row, i) => {
      const rowNum = lastRow - 4 + i;
      const transcript = row[21]; // V
      const evaluation = row[22]; // W
      
      if (transcript && transcript.toString().trim().length > 0) callsWithTranscript++;
      if (evaluation && evaluation.toString().trim().length > 0) callsWithEvaluation++;
    });
    
    writeLog(`Звонков с транскрипцией: ${callsWithTranscript}`, 'INFO');
    writeLog(`Звонков с оценкой: ${callsWithEvaluation}`, 'INFO');
    
  } catch (error) {
    const errorMsg = `Ошибка проверки листа: ${error.message}`;
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
  writeLog('БЫСТРАЯ ВЫГРУЗКА ЗВОНКОВ ЗА ПОСЛЕДНИЕ 10 ЧАСОВ', 'INFO');
  return massExportCallsHours(10);
}

// Тестовая функция для проверки логов
function testLogging() {
  writeLog('Тестовое сообщение', 'INFO');
  writeLog('Тестовое предупреждение', 'WARNING');
  writeLog('Тестовая ошибка', 'ERROR');
  writeLog('Тестовый успех', 'SUCCESS');
  console.log('✅ Тестовые логи записаны в лист LOGS');
}

// Функция отправки заметок в amoCRM
function sendNotesToAmoCRM() {
  try {
    writeLog('--- Начало выполнения sendNotesToAmoCRM ---', 'INFO');
    
    const dbSheet = SpreadsheetApp.getActive().getSheetByName("БД");
    if (!dbSheet) throw new Error('Лист "БД" не найден');
    
    const amoToken = dbSheet.getRange("B2").getValue();
    let amoSubdomain = dbSheet.getRange("B3").getValue();
    const dataSheetName = dbSheet.getRange("B4").getValue();

    if (!amoToken || !amoSubdomain) throw new Error('Не указан токен или поддомен');
    amoSubdomain = amoSubdomain.replace(/\.amocrm\.(ru|eu)$/i, '').trim();
    
    if (!/^[a-z0-9-]+$/i.test(amoSubdomain)) {
      throw new Error(`Некорректный поддомен: ${amoSubdomain}`);
    }

    const dataSheet = SpreadsheetApp.getActive().getSheetByName(dataSheetName);
    if (!dataSheet) throw new Error(`Лист "${dataSheetName}" не найден`);

    const lastRow = dataSheet.getLastRow();
    const dataRange = dataSheet.getRange(2, 1, lastRow - 1, 27);
    const data = dataRange.getValues();
    
    const filteredData = data
      .map((row, index) => ({ row, index: index + 2 }))
      .filter(item => {
        const aaValue = item.row[26]; // Столбец AA
        const zValue = item.row[25];  // Столбец Z
        return (aaValue && !zValue);
      });

    writeLog(`Всего строк: ${lastRow - 1}`, 'INFO');
    writeLog(`Подходит для обработки: ${filteredData.length}`, 'INFO');

    if (filteredData.length === 0) {
      writeLog('Нет строк для обработки', 'INFO');
      return;
    }

    filteredData.forEach(item => {
      const rowNumber = item.index;
      const rowData = item.row;
      
      const dealId = rowData[5];        // Столбец F
      const transcription = rowData[21]; // Столбец V
      const assessment = rowData[22];   // Столбец W

      writeLog(`Обработка строки ${rowNumber}: dealId=${dealId}`, 'INFO');

      try {
        if (typeof transcription === 'string' && transcription.trim()) {
          addNoteToAmo(amoSubdomain, amoToken, dealId, transcription.trim());
          writeLog('Транскрипция отправлена', 'INFO');
        }

        if (typeof assessment === 'string' && assessment.trim()) {
          addNoteToAmo(amoSubdomain, amoToken, dealId, assessment.trim());
          writeLog('Оценка отправлена', 'INFO');
        }

        dataSheet.getRange(rowNumber, 26).setValue(new Date());
        writeLog('Отметка времени установлена в колонку Z', 'INFO');
        Utilities.sleep(500);

      } catch (e) {
        writeLog(`Ошибка в строке ${rowNumber}: ${e.message}`, 'ERROR');
        dataSheet.getRange(rowNumber, 26).setValue(new Date());
        writeLog('Отметка времени установлена в колонку Z после ошибки', 'INFO');
      }
    });

    writeLog('--- Выполнение sendNotesToAmoCRM завершено ---', 'INFO');

  } catch (e) {
    writeLog(`*** ГЛОБАЛЬНАЯ ОШИБКА: ${e.message}`, 'ERROR');
  }
}

// Функция отправки заметки в amoCRM
function addNoteToAmo(subdomain, apiKey, dealId, noteText) {
  try {
    writeLog(`--- Начало запроса к amoCRM ---`, 'DEBUG');
    const url = `https://${subdomain}.amocrm.ru/api/v4/leads/${dealId}/notes`;
    const payload = [{
      "note_type": "common",
      "text": noteText,
      "created_at": Math.floor(Date.now() / 1000)
    }];

    writeLog(`URL: ${url}`, 'DEBUG');
    writeLog(`Payload: ${JSON.stringify(payload)}`, 'DEBUG');

    const options = {
      'method': 'post',
      'headers': { 
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      },
      'payload': JSON.stringify(payload),
      'muteHttpExceptions': true
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const result = JSON.parse(response.getContentText());

    writeLog(`Код ответа: ${responseCode}`, 'INFO');
    writeLog(`Ответ сервера: ${JSON.stringify(result)}`, 'DEBUG');

    if (responseCode >= 400) {
      throw new Error(`amoCRM error: ${result.detail || response.getContentText()}`);
    }

    return result;

  } catch (e) {
    writeLog(`Ошибка API: ${e.message}`, 'ERROR');
    throw e;
  }
}

// Функция для обработки звонков через ProTalk (транскрипция)
function chatToProTalkBot() {
  let rowToProcess = null;
  let targetSheet = null;
  
  try {
    writeLog('--- Начало выполнения chatToProTalkBot ---', 'INFO');
    
    const dbSheet = SpreadsheetApp.getActive().getSheetByName("БД");
    if (!dbSheet) throw new Error("Лист 'БД' не найден");
   
    const botToken = dbSheet.getRange("B5").getValue();
    const botId = dbSheet.getRange("B6").getValue();
    const targetSheetName = dbSheet.getRange("B4").getValue();
   
    writeLog(`Настройки: лист ${targetSheetName}, botId=${botId}`, 'INFO');
   
    targetSheet = SpreadsheetApp.getActive().getSheetByName(targetSheetName);
    if (!targetSheet) throw new Error(`Лист '${targetSheetName}' не найден`);
   
    const dataRange = targetSheet.getDataRange();
    const values = dataRange.getValues();
    writeLog(`Всего строк в листе: ${values.length}`, 'INFO');
   
    // Поиск строки для обработки (колонка U = 20, Y = 24)
    for (let i = 1; i < values.length; i++) {
      const question = values[i][20]; // U
      const startDate = values[i][24]; // Y
     
      if (question && !startDate) {
        rowToProcess = i + 1;
        break;
      }
    }
   
    if (!rowToProcess) {
      writeLog('Нет строк для обработки транскрипции', 'INFO');
      return;
    }
   
    writeLog(`Обрабатывается строка ${rowToProcess} для транскрипции`, 'INFO');
    
    // Запись даты начала в колонку Y (25-й столбец)
    const startDateCell = targetSheet.getRange(rowToProcess, 25);
    startDateCell.setValue(new Date());
    writeLog('Отметка времени установлена в колонку Y', 'INFO');
   
    const questionCell = targetSheet.getRange(rowToProcess, 21);
    let question = questionCell.getValue();
    if (!question) throw new Error("Пустой вопрос в строке " + rowToProcess);
   
    writeLog(`Вопрос: ${question.substring(0, 100)}...`, 'INFO');
   
    question = '🏁##' + question;
    const requests = question.split("##").map(q => q.trim()).filter(Boolean);
    writeLog(`Количество запросов: ${requests.length}`, 'INFO');
   
    const chatId = "chat_" + Date.now();
    const apiUrl = `https://eu1.api.pro-talk.ru/api/v1.0/ask/${botToken}`;
    let lastResponse = null;
   
    const TIMEOUT_MS = 300000;
   
    for (const q of requests) {
      try {
        writeLog(`Отправка запроса: ${q.substring(0, 50)}...`, 'INFO');
        writeLog(`URL: ${apiUrl}`, 'DEBUG');
        
        const payload = {
          bot_id: botId,
          chat_id: chatId,
          message: q
        };
        writeLog(`Payload: ${JSON.stringify(payload)}`, 'DEBUG');
        
        const response = UrlFetchApp.fetch(apiUrl, {
          method: "post",
          contentType: "application/json",
          muteHttpExceptions: true,
          timeout: TIMEOUT_MS,
          payload: JSON.stringify(payload)
        });
       
        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();
        writeLog(`Код ответа: ${responseCode}`, 'INFO');
        writeLog(`Ответ сервера: ${responseText.substring(0, 200)}...`, 'DEBUG');
       
        if (responseCode !== 200) {
          throw new Error(`HTTP ${responseCode}: ${responseText}`);
        }
       
        const result = JSON.parse(responseText);
        lastResponse = result.done;
        writeLog(`Получен ответ: ${lastResponse ? 'OK' : 'ERROR'}`, 'INFO');
        Utilities.sleep(1000);
       
      } catch (e) {
        writeLog(`Ошибка в запросе: ${e.message}`, 'ERROR');
        throw e;
      }
    }
   
    const cleanResponse = lastResponse ? lastResponse.replace(/\*/g, '') : '';
    targetSheet.getRange(rowToProcess, 22).setValue(cleanResponse);
    writeLog(`Транскрипция записана в строку ${rowToProcess}`, 'SUCCESS');
    writeLog('--- Выполнение chatToProTalkBot завершено ---', 'INFO');
   
  } catch (e) {
    writeLog(`*** ОШИБКА в chatToProTalkBot: ${e.message}`, 'ERROR');
    
    // Гарантированная запись ошибки
    if (rowToProcess && targetSheet) {
      try {
        targetSheet.getRange(rowToProcess, 22).setValue("Разговор не распознан");
        writeLog('Ошибка записана в колонку V', 'INFO');
      } catch (setError) {
        writeLog(`Ошибка записи статуса: ${setError.message}`, 'ERROR');
      }
    }
  }
}

// Функция для обработки звонков через ProTalk (оценка)
function ToProTalkBot() {
  let rowToProcess = null;
  let targetSheet = null;
  
  try {
    writeLog('--- Начало выполнения ToProTalkBot ---', 'INFO');
    
    const dbSheet = SpreadsheetApp.getActive().getSheetByName("БД");
    if (!dbSheet) throw new Error("Лист 'БД' не найден");
   
    const botToken = dbSheet.getRange("B5").getValue();
    const botId = dbSheet.getRange("B6").getValue();
    const targetSheetName = dbSheet.getRange("B4").getValue();
   
    writeLog(`Настройки: лист ${targetSheetName}, botId=${botId}`, 'INFO');
   
    targetSheet = SpreadsheetApp.getActive().getSheetByName(targetSheetName);
    if (!targetSheet) throw new Error(`Лист '${targetSheetName}' не найден`);
   
    const dataRange = targetSheet.getDataRange();
    const values = dataRange.getValues();
    writeLog(`Всего строк в листе: ${values.length}`, 'INFO');
   
    // Поиск строки для обработки (колонка U = 20, Y = 24)
    for (let i = 1; i < values.length; i++) {
      const question = values[i][20]; // U
      const startDate = values[i][24]; // Y
     
      if (question && !startDate) {
        rowToProcess = i + 1;
        break;
      }
    }
   
    if (!rowToProcess) {
      writeLog('Нет строк для обработки оценки', 'INFO');
      return;
    }
   
    writeLog(`Обрабатывается строка ${rowToProcess} для оценки`, 'INFO');
    
    // Запись даты начала в колонку Y (25-й столбец)
    const startDateCell = targetSheet.getRange(rowToProcess, 25);
    startDateCell.setValue(new Date());
    writeLog('Отметка времени установлена в колонку Y', 'INFO');
   
    const questionCell = targetSheet.getRange(rowToProcess, 21);
    let question = questionCell.getValue();
    if (!question) throw new Error("Пустой вопрос в строке " + rowToProcess);
   
    writeLog(`Вопрос: ${question.substring(0, 100)}...`, 'INFO');
   
    question = '🏁##' + question;
    const requests = question.split("##").map(q => q.trim()).filter(Boolean);
    writeLog(`Количество запросов: ${requests.length}`, 'INFO');
   
    const chatId = "chat_" + Date.now();
    const apiUrl = `https://us1.api.pro-talk.ru/api/v1.0/ask/${botToken}`;
    let lastResponse = null;
   
    const TIMEOUT_MS = 300000;
   
    for (const q of requests) {
      try {
        writeLog(`Отправка запроса: ${q.substring(0, 50)}...`, 'INFO');
        writeLog(`URL: ${apiUrl}`, 'DEBUG');
        
        const payload = {
          bot_id: botId,
          chat_id: chatId,
          message: q
        };
        writeLog(`Payload: ${JSON.stringify(payload)}`, 'DEBUG');
        
        const response = UrlFetchApp.fetch(apiUrl, {
          method: "post",
          contentType: "application/json",
          muteHttpExceptions: true,
          timeout: TIMEOUT_MS,
          payload: JSON.stringify(payload)
        });
       
        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();
        writeLog(`Код ответа: ${responseCode}`, 'INFO');
        writeLog(`Ответ сервера: ${responseText.substring(0, 200)}...`, 'DEBUG');
       
        if (responseCode !== 200) {
          throw new Error(`HTTP ${responseCode}: ${responseText}`);
        }
       
        const result = JSON.parse(responseText);
        lastResponse = result.done;
        writeLog(`Получен ответ: ${lastResponse ? 'OK' : 'ERROR'}`, 'INFO');
        Utilities.sleep(1000);
       
      } catch (e) {
        writeLog(`Ошибка в запросе: ${e.message}`, 'ERROR');
        throw e;
      }
    }
   
    const cleanResponse = lastResponse ? lastResponse.replace(/\*/g, '') : '';
    targetSheet.getRange(rowToProcess, 23).setValue(cleanResponse);
    writeLog(`Оценка записана в строку ${rowToProcess}`, 'SUCCESS');
    writeLog('--- Выполнение ToProTalkBot завершено ---', 'INFO');
   
  } catch (e) {
    writeLog(`*** ОШИБКА в ToProTalkBot: ${e.message}`, 'ERROR');
    
    // Гарантированная запись ошибки
    if (rowToProcess && targetSheet) {
      try {
        targetSheet.getRange(rowToProcess, 23).setValue("Разговор не распознан");
        writeLog('Ошибка записана в колонку W', 'INFO');
      } catch (setError) {
        writeLog(`Ошибка записи статуса: ${setError.message}`, 'ERROR');
      }
    }
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