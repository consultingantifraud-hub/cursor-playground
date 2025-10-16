// ============================== AMO.gs (МОДЕРНИЗИРОВАННЫЙ) ==============================
// Основан на рабочем коде, но с исправлениями фильтрации событий и улучшенным логированием
// Ключевые исправления:
// 1. Правильная фильтрация событий note_added с note_type call_in/call_out
// 2. Обработка гибридных звонков (common заметки с аудио)
// 3. Улучшенное логирование и отладка
// 4. Синхронизация с другими модулями

// ---- Конфигурация (улучшенная) -----------------------------------------
function getConfigFromSheet() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = spreadsheet.getSheetByName("БД");
    if (!configSheet) throw new Error('Лист "БД" не найден');
    
    // Чтение исключений из колонки G
    const excludedFields = readExcludedFields(configSheet);
    
    // Нормализация поддомена
    let subdomain = configSheet.getRange("B3").getValue();
    if (subdomain) {
      subdomain = subdomain.toString().trim().replace(/^https?:\/\//i, '').replace(/\.amocrm\.(ru|eu)$/i, '');
      if (!/^[a-z0-9-]+$/i.test(subdomain)) {
        throw new Error(`Некорректный поддомен: ${subdomain}`);
      }
    }
    
    return {
      AMO_TOKEN: configSheet.getRange("B2").getValue(),
      AMO_SUBDOMAIN: subdomain,
      SHEET_NAME: configSheet.getRange("B4").getValue(),
      SPREADSHEET_ID: spreadsheet.getId(),
      callStatusMap: Object.fromEntries(
        configSheet.getRange("C:D").getValues()
          .filter(row => row[0] && row[1])
          .map(([id, text]) => [String(id), String(text)])
      ),
      taskTypesMap: Object.fromEntries(
        configSheet.getRange("E:F").getValues()
          .filter(row => row[0] && row[1])
          .map(([id, text]) => [String(id), String(text)])
      ),
      excludedFields: excludedFields
    };
  } catch (error) {
    console.error('❌ Ошибка конфигурации:', error.message);
    throw error;
  }
}

// Чтение исключенных полей
function readExcludedFields(sheet) {
  const excludedRange = sheet.getRange("G2:G");
  const values = excludedRange.getValues();
  return values
    .filter(row => row[0] !== "")
    .map(row => String(row[0]).trim());
}

// ---- Утилиты ------------------------------------------------------------
function formatTimestamp(timestamp) {
  if (!timestamp) return "Нет данных";
  const date = new Date(timestamp * 1000);
  return Utilities.formatDate(date, "GMT+3", "dd.MM.yyyy HH:mm");
}

// Проверка обработки звонка (улучшенная)
function isCallProcessed(callNoteId) {
  try {
    const config = getConfigFromSheet();
    const sheet = SpreadsheetApp.openById(config.SPREADSHEET_ID)
      .getSheetByName(config.SHEET_NAME);
    if (!sheet) return false;
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return false; // только заголовки или пустой лист
    
    const processedIds = sheet.getRange(2, 15, lastRow - 1, 1)
      .getValues()
      .flat()
      .map(id => String(id || '').trim())
      .filter(id => id !== '');
    
    return processedIds.includes(String(callNoteId));
  } catch (error) {
    console.error('❌ Ошибка проверки:', error.message);
    return false;
  }
}

// ---- API запросы (улучшенные) ------------------------------------------
function amoRequest(url, method = 'GET', body = null) {
  try {
    const options = {
      method: method,
      headers: { 
        'Authorization': `Bearer ${getConfigFromSheet().AMO_TOKEN}`,
        'Accept': 'application/hal+json'
      },
      muteHttpExceptions: true,
      followRedirects: true
    };
    
    if (body) {
      options.payload = JSON.stringify(body);
      options.contentType = 'application/json';
    }
    
    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();
    
    if (statusCode === 204) {
      console.log(`✅ Пустой ответ (204) для ${url}`);
      return null;
    }
    
    if (statusCode < 200 || statusCode >= 300) {
      const errorText = response.getContentText();
      console.error(`❌ HTTP ${statusCode}: ${errorText.slice(0, 200)}`);
      throw new Error(`HTTP ${statusCode}: ${errorText.slice(0, 200)}`);
    }
    
    const responseText = response.getContentText();
    return responseText ? JSON.parse(responseText) : null;
  } catch (error) {
    console.error(`❌ Ошибка запроса к ${url}:`, error.message);
    throw error;
  }
}

// ---- Получение данных сущностей ----------------------------------------
function getAmoLead(leadId) {
  try {
    const config = getConfigFromSheet();
    const url = `https://${config.AMO_SUBDOMAIN}.amocrm.ru/api/v4/leads/${leadId}?with=contacts,companies`;
    const lead = amoRequest(url);
    if (!lead) return null;
    lead.responsible = getAmoUser(lead.responsible_user_id).name;
    return lead;
  } catch (error) {
    console.error(`❌ Ошибка сделки ID ${leadId}:`, error.message);
    return null;
  }
}

function getLeadByContact(contactId) {
  try {
    const config = getConfigFromSheet();
    const url = `https://${config.AMO_SUBDOMAIN}.amocrm.ru/api/v4/contacts/${contactId}/links`;
    const linksResponse = amoRequest(url);
    const links = linksResponse?._embedded?.links || [];
    
    const leadIds = links
      .filter(link => link.to_entity_type === 'leads')
      .map(link => link.to_entity_id);
    
    return leadIds.length > 0 ? getAmoLead(leadIds[0]) : null;
  } catch (error) {
    console.error(`❌ Ошибка поиска сделок для контакта ${contactId}:`, error.message);
    return null;
  }
}

function getLeadByCompany(companyId) {
  try {
    const config = getConfigFromSheet();
    const url = `https://${config.AMO_SUBDOMAIN}.amocrm.ru/api/v4/companies/${companyId}/links`;
    const linksResponse = amoRequest(url);
    const links = linksResponse?._embedded?.links || [];
    
    const leadIds = links
      .filter(link => link.to_entity_type === 'leads')
      .map(link => link.to_entity_id);
    
    return leadIds.length > 0 ? getAmoLead(leadIds[0]) : null;
  } catch (error) {
    console.error(`❌ Ошибка поиска сделок для компании ${companyId}:`, error.message);
    return null;
  }
}

function getEntityData(entityType, entityId) {
  try {
    const config = getConfigFromSheet();
    const url = `https://${config.AMO_SUBDOMAIN}.amocrm.ru/api/v4/${entityType}/${entityId}`;
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

function getAmoNotes(entityType, entityId) {
  try {
    const config = getConfigFromSheet();
    const url = `https://${config.AMO_SUBDOMAIN}.amocrm.ru/api/v4/${entityType}/${entityId}/notes?with=attachments`;
    const response = amoRequest(url);
    return response?._embedded?.notes || [];
  } catch (error) {
    console.error(`❌ Ошибка примечаний для ${entityType} ID ${entityId}:`, error.message);
    return [];
  }
}

function getAmoPipeline(pipelineId) {
  try {
    const config = getConfigFromSheet();
    const url = `https://${config.AMO_SUBDOMAIN}.amocrm.ru/api/v4/leads/pipelines`;
    const pipelines = amoRequest(url)?._embedded?.pipelines || [];
    return pipelines.find(p => p.id === pipelineId) || { id: pipelineId, name: "Неизвестная воронка" };
  } catch (error) {
    console.error(`❌ Ошибка воронки ID ${pipelineId}:`, error.message);
    return { id: pipelineId, name: "Неизвестная воронка" };
  }
}

function getAmoStatus(statusId, pipelineId) {
  try {
    const config = getConfigFromSheet();
    const url = `https://${config.AMO_SUBDOMAIN}.amocrm.ru/api/v4/leads/pipelines`;
    const pipelines = amoRequest(url)?._embedded?.pipelines || [];
    const pipeline = pipelines.find(p => p.id === pipelineId);
    const statuses = pipeline?._embedded?.statuses || [];
    return statuses.find(s => s.id === statusId) || { id: statusId, name: "Неизвестный статус" };
  } catch (error) {
    console.error(`❌ Ошибка статуса ID ${statusId}:`, error.message);
    return { id: statusId, name: "Неизвестный статус" };
  }
}

function getAmoUser(userId) {
  try {
    const config = getConfigFromSheet();
    const url = `https://${config.AMO_SUBDOMAIN}.amocrm.ru/api/v4/users/${userId}`;
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

// ---- Проверка: является ли заметка звонком (КЛЮЧЕВОЕ ИСПРАВЛЕНИЕ) -----
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

// ---- Получение задач для сделки ----------------------------------------
function getAmoTasks(leadId) {
  try {
    const config = getConfigFromSheet();
    const url = `https://${config.AMO_SUBDOMAIN}.amocrm.ru/api/v4/tasks?filter[entity_id]=${leadId}&filter[entity_type]=leads`;
    const tasksResponse = amoRequest(url);
    if (!tasksResponse) return 'Нет задач';
    const tasks = tasksResponse._embedded?.tasks || [];
    return tasks.map(task => {
      const taskType = config.taskTypesMap[String(task.task_type_id)] || `Неизвестный тип (${task.task_type_id})`;
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

// ---- Сбор всех полей из сделки (с фильтрацией) ------------------------
function getAllEntityFields(lead) {
  try {
    const config = getConfigFromSheet();
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

// ---- Обработка события (КЛЮЧЕВОЕ ИСПРАВЛЕНИЕ) -------------------------
function processEvent(eventData) {
  try {
    console.log(`📞 Обработка события ID: ${eventData.id}, тип: ${eventData.type}`);
    
    // Проверяем, что это событие добавления заметки
    if (eventData.type !== 'note_added') {
      console.log(`⏭️ Пропускаем событие типа: ${eventData.type}`);
      return;
    }
    
    const callNoteId = eventData.value_after?.[0]?.note?.id;
    if (!callNoteId) {
      console.log('⏭️ Нет ID заметки в событии');
      return;
    }
    
    if (isCallProcessed(callNoteId)) {
      console.log(`⏭️ Звонок ${callNoteId} уже обработан`);
      return;
    }

    let lead;
    const entityType = String(eventData.entity_type || '').toLowerCase();
    
    switch(entityType) {
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
        console.log(`⏭️ Неизвестный тип сущности: ${entityType}`);
        return;
    }

    if (!lead) {
      console.log(`⏭️ Не найдена сделка для ${entityType} ID ${eventData.entity_id}`);
      return;
    }

    // Получение задач и примечаний
    const tasksText = getAmoTasks(lead.id);
    const notes = getAmoNotes('leads', lead.id);
    const callNote = notes.find(n => String(n.id) === String(callNoteId));

    // КЛЮЧЕВОЕ ИСПРАВЛЕНИЕ: проверяем, является ли заметка звонком
    if (!callNote || !isCallNote(callNote)) {
      console.log(`⏭️ Заметка ${callNoteId} не является звонком (тип: ${callNote?.note_type})`);
      return;
    }

    console.log(`✅ Обрабатываем звонок ${callNoteId} для сделки ${lead.id}`);

    // Обработка примечаний
    let notesText = notes.map(note => {
      let text = '';
      switch(note.note_type) {
        case 'common':
          text = `📝 Примечание: ${note.params?.text || ''}`;
          if (note._embedded?.attachments?.length) {
            text += `\nattachments:\n${note._embedded.attachments
              .map(att => `• ${att.name} (${att.link})`)
              .join('\n')}`;
          }
          break;
        case 'call_out':
        case 'call_in':
          const duration = note.params?.duration ? ` (${note.params.duration}с)` : '';
          const result = getConfigFromSheet().callStatusMap[String(note.params?.call_status)] || 'Без результата';
          text = `📞 Звонок${duration}: ${result}`;
          if (note.params?.link) text += ` (${note.params.link})`;
          if (note._embedded?.attachments?.length) {
            text += `\nattachments:\n${note._embedded.attachments
              .map(att => `• ${att.name} (${att.link})`)
              .join('\n')}`;
          }
          break;
        case 'task':
          text = `✅ Задача [ID:${note.id}]: ${note.params?.text || ''}`;
          text += `\nСтатус: ${note.params?.status === 1 ? 'Выполнена' : 'Не выполнена'}`;
          if (note.params?.task_deadline) {
            text += `\nСрок: ${new Date(note.params.task_deadline * 1000)}`;
          }
          break;
        case 'mail':
          text = `📧 Почта: ${note.params?.subject || ''}`;
          text += `\nОт: ${note.params?.from || ''}`;
          text += `\nСообщение: ${note.params?.text || ''}`;
          if (note.params?.link) text += `\nСвязь: ${note.params.link}`;
          break;
        case 'chat':
          text = `💬 Чат (${note.params?.service || ''}):`;
          text += `\nСообщение: ${note.params?.text || ''}`;
          if (note.params?.link) text += `\nСвязь: ${note.params.link}`;
          break;
        default:
          text = `Неизвестный тип: ${note.note_type}`;
      }
      return text;
    }).join('\n').trim() || 'Нет примечаний';

    // Добавляем задачи в примечания
    notesText += `\n📌 Задачи в сделке:\n${tasksText}`;

    // Формируем данные
    const isOutgoing = String(callNote.note_type || '').toLowerCase() === 'call_out';
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
      notesText, // 11. Примечания + задачи
      callNote?.params?.duration || 0, // 12. Длительность
      getConfigFromSheet().callStatusMap[String(callNote?.params?.call_status)] || 'Без результата', // 13. Результат
      isOutgoing ? 'Исходящий' : 'Входящий', // 14. Тип
      callNoteId, // 15. ID звонка
      getAllEntityFields(lead) // 16. Все поля
    ];
    
    appendToSheet(rowData);
    console.log(`✅ Звонок ${callNoteId} записан в таблицу`);
  } catch (error) {
    console.error(`❌ Ошибка в событии ID ${eventData.id}:`, error.message);
  }
}

// ---- Запись данных в таблицу -------------------------------------------
function appendToSheet(rowData) {
  try {
    const config = getConfigFromSheet();
    const sheet = SpreadsheetApp.openById(config.SPREADSHEET_ID)
      .getSheetByName(config.SHEET_NAME);
    if (!sheet) throw new Error('Лист не найден');
    sheet.appendRow(rowData);
    console.log('✅ Данные записаны в таблицу');
  } catch (error) {
    console.error('❌ Ошибка записи:', error.message);
  }
}

// ---- Основная функция синхронизации (ИСПРАВЛЕННАЯ) --------------------
function syncEventsToday() {
  try {
    console.log('🔄 Начало синхронизации событий');
    const config = getConfigFromSheet();
    const timeFrom = Math.floor((new Date().getTime() - 5 * 60 * 1000) / 1000); // последние 5 минут
    let url = `https://${config.AMO_SUBDOMAIN}.amocrm.ru/api/v4/events?filter[created_at][from]=${timeFrom}&limit=250`;
    let allEvents = [];
    let processed = 0;
    let written = 0;
    
    // Получаем все события
    while (url) {
      const response = amoRequest(url);
      if (!response?._embedded?.events) break;
      allEvents = allEvents.concat(response._embedded.events);
      url = response._links?.next?.href || '';
      if (allEvents.length > 3000) break; // предохранитель
    }
    
    console.log(`📊 Найдено событий: ${allEvents.length}`);
    
    // Фильтруем события добавления заметок для нужных сущностей
    const noteEvents = allEvents.filter(e => {
      if (e.type !== 'note_added') return false;
      if (!e.value_after?.[0]?.note?.id) return false;
      const entityType = String(e.entity_type || '').toLowerCase();
      return ['lead', 'contact', 'company'].includes(entityType);
    });
    
    console.log(`📝 События с заметками: ${noteEvents.length}`);
    
    // Обрабатываем каждое событие
    noteEvents.forEach(event => {
      const before = SpreadsheetApp.openById(config.SPREADSHEET_ID)
        .getSheetByName(config.SHEET_NAME).getLastRow();
      processEvent(event);
      processed++;
      const after = SpreadsheetApp.openById(config.SPREADSHEET_ID)
        .getSheetByName(config.SHEET_NAME).getLastRow();
      if (after > before) written++;
    });
    
    console.log(`✅ Обработано: ${processed}, записано: ${written}`);
  } catch (error) {
    console.error('❌ Ошибка синхронизации:', error.message);
  }
}

// ---- Разовый сбор за последние N часов --------------------------------
function syncEventsLastHours(hours) {
  try {
    hours = Number(hours) || 3;
    console.log(`🔄 Сбор событий за последние ${hours} часов`);
    const config = getConfigFromSheet();
    const timeFrom = Math.floor((new Date().getTime() - hours * 3600 * 1000) / 1000);
    let url = `https://${config.AMO_SUBDOMAIN}.amocrm.ru/api/v4/events?filter[created_at][from]=${timeFrom}&limit=250`;
    let allEvents = [];
    let processed = 0;
    let written = 0;
    
    // Получаем все события
    while (url) {
      const response = amoRequest(url);
      if (!response?._embedded?.events) break;
      allEvents = allEvents.concat(response._embedded.events);
      url = response._links?.next?.href || '';
      if (allEvents.length > 5000) break; // предохранитель
    }
    
    console.log(`📊 Найдено событий: ${allEvents.length}`);
    
    // Фильтруем события добавления заметок для нужных сущностей
    const noteEvents = allEvents.filter(e => {
      if (e.type !== 'note_added') return false;
      if (!e.value_after?.[0]?.note?.id) return false;
      const entityType = String(e.entity_type || '').toLowerCase();
      return ['lead', 'contact', 'company'].includes(entityType);
    });
    
    console.log(`📝 События с заметками: ${noteEvents.length}`);
    
    // Обрабатываем каждое событие
    noteEvents.forEach(event => {
      const before = SpreadsheetApp.openById(config.SPREADSHEET_ID)
        .getSheetByName(config.SHEET_NAME).getLastRow();
      processEvent(event);
      processed++;
      const after = SpreadsheetApp.openById(config.SPREADSHEET_ID)
        .getSheetByName(config.SHEET_NAME).getLastRow();
      if (after > before) written++;
    });
    
    console.log(`✅ Обработано: ${processed}, записано: ${written}`);
  } catch (error) {
    console.error('❌ Ошибка сбора событий:', error.message);
  }
}

// ---- Логирование событий для отладки ----------------------------------
function logEvents() {
  try {
    console.log('🔍 Логирование событий для отладки');
    const config = getConfigFromSheet();
    const timeFrom = Math.floor((new Date().getTime() - 150 * 60 * 1000) / 1000);
    let url = `https://${config.AMO_SUBDOMAIN}.amocrm.ru/api/v4/events?filter[created_at][from]=${timeFrom}&limit=250`;
    let allEvents = [];
    
    while (url) {
      const response = amoRequest(url);
      if (!response?._embedded?.events) break;
      allEvents = allEvents.concat(response._embedded.events);
      url = response._links?.next?.href || '';
      if (allEvents.length > 100) break; // ограничиваем для отладки
    }
    
    console.log(`📊 Всего событий: ${allEvents.length}`);
    
    // Показываем типы событий
    const eventTypes = {};
    allEvents.forEach(e => {
      eventTypes[e.type] = (eventTypes[e.type] || 0) + 1;
    });
    console.log('📈 Типы событий:', eventTypes);
    
    // Показываем события с заметками
    const noteEvents = allEvents.filter(e => e.type === 'note_added');
    console.log(`📝 События с заметками: ${noteEvents.length}`);
    
    // Показываем первые несколько событий с заметками
    noteEvents.slice(0, 5).forEach((event, i) => {
      console.log(`📄 Событие ${i+1}:`, {
        id: event.id,
        type: event.type,
        entity_type: event.entity_type,
        entity_id: event.entity_id,
        note_id: event.value_after?.[0]?.note?.id,
        note_type: event.value_after?.[0]?.note?.note_type
      });
    });
    
  } catch (error) {
    console.error('❌ Ошибка логирования событий:', error.message);
  }
}