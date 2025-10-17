// ============================== AMO.gs ===========================
// Основная логика интеграции с amoCRM
// Обработка событий, получение данных, синхронизация

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
    writeLog(`Обработка события ID: ${eventData.id}`, 'INFO');
    console.log(`📞 Обработка события ID: ${eventData.id}`);
    
    const callNoteId = eventData.value_after?.[0]?.note?.id;
    if (!callNoteId || isCallProcessed(callNoteId)) {
      writeLog(`Событие ${eventData.id} пропущено (уже обработано или нет ID)`, 'INFO');
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
        console.log(`Неизвестный тип сущности: ${eventData.entity_type}`);
        return;
    }

    if (!lead) return;

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
    writeLog(`Событие ${eventData.id} успешно обработано`, 'SUCCESS');
  } catch (error) {
    writeLog(`Ошибка в событии ID ${eventData.id}: ${error.message}`, 'ERROR');
    console.error(`❌ Ошибка в событии ID ${eventData.id}:`, error.message);
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

// Основная функция синхронизации
function syncEventsToday() {
  try {
    writeLog('--- Начало синхронизации событий ---', 'INFO');
    console.log('--- Начало синхронизации событий ---');
    
    const timeFrom = Math.floor((new Date().getTime() - 3 * 60 * 1000) / 1000); // 3 минуты
    let url = `https://${config.AMO_SUBDOMAIN}/api/v4/events?filter[created_at][from]=${timeFrom}&limit=250`;
    let allEvents = [];
    
    writeLog(`Поиск событий с ${new Date(timeFrom * 1000)}`, 'INFO');
    
    while (true) {
      const response = amoRequest(url);
      if (!response?._embedded?.events) break;
      allEvents = allEvents.concat(response._embedded.events);
      url = response._links?.next?.href || '';
      if (!url) break;
    }
    
    writeLog(`Всего событий: ${allEvents.length}`, 'INFO');
    console.log(`Всего событий: ${allEvents.length}`);
    
    const calls = allEvents.filter(e =>
      ['outgoing_call', 'incoming_call'].includes(e.type)
    );
    
    writeLog(`Найдено звонков: ${calls.length}`, 'INFO');
    console.log(`Найдено звонков: ${calls.length}`);
    
    calls.forEach(event => processEvent(event));
    
    writeLog(`Обработано звонков: ${calls.length}`, 'SUCCESS');
    writeLog('--- Синхронизация завершена ---', 'INFO');
    console.log('--- Синхронизация завершена ---');
  } catch (error) {
    writeLog(`Ошибка синхронизации: ${error.message}`, 'ERROR');
    console.error('❌ Ошибка синхронизации:', error.message);
  }
}

// Функция для массовой выгрузки звонков
function massExportCalls(days = 7) {
  try {
    console.log(`--- Массовая выгрузка звонков за ${days} дней ---`);
    
    const timeFrom = Math.floor((new Date().getTime() - days * 24 * 60 * 60 * 1000) / 1000);
    let url = `https://${config.AMO_SUBDOMAIN}/api/v4/events?filter[created_at][from]=${timeFrom}&limit=250`;
    let allEvents = [];
    let page = 1;
    
    while (true) {
      console.log(`Обработка страницы ${page}...`);
      const response = amoRequest(url);
      if (!response?._embedded?.events) break;
      
      allEvents = allEvents.concat(response._embedded.events);
      console.log(`Страница ${page}: ${response._embedded.events.length} событий`);
      
      url = response._links?.next?.href || '';
      if (!url) break;
      page++;
    }
    
    console.log(`Всего событий за ${days} дней: ${allEvents.length}`);
    
    const calls = allEvents.filter(e =>
      ['outgoing_call', 'incoming_call'].includes(e.type)
    );
    
    console.log(`Найдено звонков: ${calls.length}`);
    calls.forEach(event => processEvent(event));
    
    console.log('--- Массовая выгрузка завершена ---');
    return calls.length;
  } catch (error) {
    console.error('❌ Ошибка массовой выгрузки:', error.message);
    return 0;
  }
}

// Быстрая выгрузка за последние 24 часа
function quickExport() {
  return massExportCalls(1);
}

// Полная выгрузка за 30 дней
function fullExport() {
  return massExportCalls(30);
}

// Тестовая функция для проверки логов
function testLogging() {
  writeLog('Тестовое сообщение', 'INFO');
  writeLog('Тестовое предупреждение', 'WARNING');
  writeLog('Тестовая ошибка', 'ERROR');
  writeLog('Тестовый успех', 'SUCCESS');
  console.log('✅ Тестовые логи записаны в лист LOGS');
}