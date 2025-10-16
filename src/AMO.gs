// ============================== AMO.gs (переписанный) ==============================
// Сбор событий из amoCRM в таблицу: сделки, заметки, задачи, аудио-ссылка,
// «Все поля» с фильтрацией. Без общего config-файла. Синглтон-сторож для синка.
//
// Ключевые отличия:
//  - processEvent_: обрабатываем ТОЛЬКО когда соответствующая note реально типа call_in/call_out.
//  - Сравнение ID заметок как строк (исправляет пропуски из-за типов).
//  - syncEventsToday(): тащит все события с value_after.note.id по нужным сущностям,
//    а фильтр «это звонок» делает processEvent_.
//  - Добавлен разовый сбор за часы: syncEventsLastHours(hours).
//
// Требуемые листы/ячейки в «БД»:
//  B2: AMO_LONG_TERM_TOKEN (Bearer)
//  B3: AMO_SUBDOMAIN/host (например, 100detaley.amocrm.ru)
//  B4: SHEET_NAME (например, "АМО Звонки")
//  Колонки C:D — соответствия статусов звонка (id → текст)
//  Колонки E:F — соответствия типов задач (id → текст)
//  Колонка G — список исключаемых полей (по именам)

// ---- Синглтон для синка -------------------------------------------------
var SYNC_RUN_KEY = 'RUN_SYNC_EVENTS';
var SYNC_WALL_MS = 25000; // 25 сек на один прогон

function _beginSingletonRun_(key, wallMs) {
  var props = PropertiesService.getScriptProperties();
  var now = Date.now();
  var raw = props.getProperty(key);
  if (raw) {
    var ts = Number(raw) || 0;
    if (ts && now - ts < wallMs) {
      console.log('[GUARD] skip ' + key + ' (already running)');
      return { ok:false, release:function(){} };
    }
  }
  props.setProperty(key, String(now));
  return { ok:true, release:function(){
    try { PropertiesService.getScriptProperties().deleteProperty(key); } catch(_) {}
  }};
}

// ---- Конфиг из листа «БД» ----------------------------------------------
function getCfg_() {
  var ss = SpreadsheetApp.getActive();
  var db = ss.getSheetByName('БД');
  if (!db) throw new Error('Лист "БД" не найден');

  function _host(v) {
    var h = String(v || '').trim().replace(/^https?:\/\//i,'').replace(/\/+$/,'');
    if (!/\.amocrm\.(ru|eu)$/i.test(h)) h += '.amocrm.ru';
    return h.toLowerCase();
  }

  return {
    AMO_TOKEN      : String(db.getRange('B2').getValue() || '').trim(),
    AMO_HOST       : _host(db.getRange('B3').getValue()),
    SHEET_NAME     : String(db.getRange('B4').getValue() || '').trim(),
    SPREADSHEET_ID : SpreadsheetApp.getActive().getId(),
    callStatusMap  : Object.fromEntries(
      db.getRange('C:D').getValues()
        .filter(function(r){ return r[0] && r[1]; })
        .map(function(r){ return [String(r[0]), String(r[1])]; })
    ),
    taskTypesMap   : Object.fromEntries(
      db.getRange('E:F').getValues()
        .filter(function(r){ return r[0] && r[1]; })
        .map(function(r){ return [String(r[0]), String(r[1])]; })
    ),
    excludedFields : db.getRange('G2:G').getValues()
      .map(function(r){ return String(r[0]||'').trim(); })
      .filter(Boolean)
  };
}

// ---- Универсальный запрос к amo ----------------------------------------
function amoRequest_(host, token, urlPath, opts) {
  var url = /^https?:\/\//i.test(urlPath) ? urlPath : ('https://' + host + urlPath);
  opts = opts || {};
  var options = {
    method : opts.method || 'get',
    headers: Object.assign(
      { 'Authorization': 'Bearer ' + token, 'Accept':'application/hal+json' },
      opts.headers || {}
    ),
    payload: opts.payload ? JSON.stringify(opts.payload) : undefined,
    contentType: opts.payload ? 'application/json' : undefined,
    muteHttpExceptions: true,
    followRedirects: true
  };
  var resp = UrlFetchApp.fetch(url, options);
  var code = resp.getResponseCode();
  if (code === 204) return null;
  var txt = resp.getContentText() || '';
  if (code < 200 || code >= 300) {
    throw new Error('HTTP ' + code + ' ' + url + ' | ' + txt.slice(0,600));
  }
  return txt ? JSON.parse(txt) : null;
}

// ---- Утилиты ------------------------------------------------------------
function formatTs_(ts) {
  if (!ts) return 'Нет данных';
  return Utilities.formatDate(new Date(ts*1000), 'GMT+3', 'dd.MM.yyyy HH:mm');
}

// Проверка: уже записывали ли звонок (по колонке O = 15-й столбец данных)
function isCallProcessed_(sheetName, spreadsheetId, callNoteId) {
  try {
    var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
    if (!sheet) return false;
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return false; // только заголовки или пустой лист
    
    var rng = sheet.getRange(2, 15, lastRow - 1, 1).getValues(); // O: 15
    var ids = [].concat.apply([], rng)
      .map(function(v){ return String(v||'').trim(); })
      .filter(Boolean);
    return ids.indexOf(String(callNoteId)) !== -1;
  } catch (e){
    console.warn('[CHECK] ' + e);
    return false;
  }
}

// ---- Получение сущностей ------------------------------------------------
function getAmoLead_(cfg, leadId) {
  try {
    var lead = amoRequest_(cfg.AMO_HOST, cfg.AMO_TOKEN, '/api/v4/leads/' + leadId + '?with=contacts,companies');
    if (!lead) return null;
    lead.responsible = getAmoUser_(cfg, lead.responsible_user_id).name;
    return lead;
  } catch(e){ console.warn('[LEAD] ' + e); return null; }
}

function getEntityData_(cfg, entityType, entityId) {
  try {
    var entity = amoRequest_(cfg.AMO_HOST, cfg.AMO_TOKEN, '/api/v4/' + entityType + '/' + entityId);
    if (!entity) return null;
    return {
      id: entity.id,
      name: entity.name,
      responsible: getAmoUser_(cfg, entity.responsible_user_id).name,
      custom_fields_values: entity.custom_fields_values || []
    };
  } catch(e){ console.warn('[ENTITY] ' + entityType + ' ' + entityId + ' ' + e); return null; }
}

function getLeadByContact_(cfg, contactId) {
  try {
    var links = amoRequest_(cfg.AMO_HOST, cfg.AMO_TOKEN, '/api/v4/contacts/' + contactId + '/links');
    var ids = (links?._embedded?.links || [])
      .filter(function(l){ return l.to_entity_type==='leads'; })
      .map(function(l){ return l.to_entity_id; });
    return ids.length ? getAmoLead_(cfg, ids[0]) : null;
  } catch(e){ console.warn('[LEADS_BY_CONTACT] ' + e); return null; }
}

function getLeadByCompany_(cfg, companyId) {
  try {
    var links = amoRequest_(cfg.AMO_HOST, cfg.AMO_TOKEN, '/api/v4/companies/' + companyId + '/links');
    var ids = (links?._embedded?.links || [])
      .filter(function(l){ return l.to_entity_type==='leads'; })
      .map(function(l){ return l.to_entity_id; });
    return ids.length ? getAmoLead_(cfg, ids[0]) : null;
  } catch(e){ console.warn('[LEADS_BY_COMPANY] ' + e); return null; }
}

function getAmoNotes_(cfg, entityType, entityId) {
  try {
    var resp = amoRequest_(cfg.AMO_HOST, cfg.AMO_TOKEN, '/api/v4/' + entityType + '/' + entityId + '/notes?with=attachments');
    return resp?._embedded?.notes || [];
  } catch(e){ console.warn('[NOTES] ' + e); return []; }
}

function getAmoPipeline_(cfg, pipelineId) {
  try {
    var resp = amoRequest_(cfg.AMO_HOST, cfg.AMO_TOKEN, '/api/v4/leads/pipelines');
    var p = (resp?._embedded?.pipelines || []).find(function(x){ return x.id === pipelineId; });
    return p || { id:pipelineId, name:'Неизвестная воронка' };
  } catch(e){ return { id:pipelineId, name:'Неизвестная воронка' }; }
}

function getAmoStatus_(cfg, statusId, pipelineId) {
  try {
    var resp = amoRequest_(cfg.AMO_HOST, cfg.AMO_TOKEN, '/api/v4/leads/pipelines');
    var p = (resp?._embedded?.pipelines || []).find(function(x){ return x.id === pipelineId; });
    var s = (p?._embedded?.statuses || []).find(function(st){ return st.id === statusId; });
    return s || { id:statusId, name:'Неизвестный статус' };
  } catch(e){ return { id:statusId, name:'Неизвестный статус' }; }
}

function getAmoUser_(cfg, userId) {
  try {
    var u = amoRequest_(cfg.AMO_HOST, cfg.AMO_TOKEN, '/api/v4/users/' + userId);
    return u ? { id:u.id, name: u.name || 'Неизвестный пользователь' } : { id:userId, name:'Неизвестный пользователь' };
  } catch(e){ return { id:userId, name:'Неизвестный пользователь' }; }
}

// ---- Сбор «всех полей» с фильтрацией -----------------------------------
function getAllEntityFields_(cfg, lead) {
  try {
    var out = [];
    out.push('--- Сделка ---');
    out.push('ID: ' + lead.id);
    out.push('Ответственный: ' + getAmoUser_(cfg, lead.responsible_user_id).name);

    var dealFields = (lead.custom_fields_values || []).map(function(field){
      var nm = field.field_name;
      if (cfg.excludedFields.indexOf(nm) !== -1) return null;
      var v = (field.values||[]).map(function(x){ return x.value; }).filter(Boolean).join(', ');
      return v ? (nm + ': ' + v) : null;
    }).filter(Boolean);
    Array.prototype.push.apply(out, dealFields);

    if (lead._embedded && lead._embedded.contacts && lead._embedded.contacts.length) {
      lead._embedded.contacts.forEach(function(c){
        var contact = getEntityData_(cfg,'contacts', c.id);
        if (contact) {
          out.push('--- Контакт ---');
          out.push('ID: ' + contact.id);
          out.push('Ответственный: ' + contact.responsible);
          var cf = (contact.custom_fields_values||[]).map(function(f){
            var nm = f.field_name;
            if (cfg.excludedFields.indexOf(nm) !== -1) return null;
            var v = (f.values||[]).map(function(x){ return x.value; }).filter(Boolean).join(', ');
            return v ? (nm + ': ' + v) : null;
          }).filter(Boolean);
          Array.prototype.push.apply(out, cf);
        }
      });
    }

    if (lead._embedded && lead._embedded.companies && lead._embedded.companies.length) {
      lead._embedded.companies.forEach(function(c){
        var company = getEntityData_(cfg,'companies', c.id);
        if (company) {
          out.push('--- Компания ---');
          out.push('ID: ' + company.id);
          out.push('Ответственный: ' + company.responsible);
          var cf = (company.custom_fields_values||[]).map(function(f){
            var nm = f.field_name;
            if (cfg.excludedFields.indexOf(nm) !== -1) return null;
            var v = (f.values||[]).map(function(x){ return x.value; }).filter(Boolean).join(', ');
            return v ? (nm + ': ' + v) : null;
          }).filter(Boolean);
          Array.prototype.push.apply(out, cf);
        }
      });
    }
    return out.join('\n') || 'Нет данных';
  } catch(e){ console.warn('[ALL_FIELDS] ' + e); return 'Ошибка'; }
}

// ---- Запись в таблицу ---------------------------------------------------
function appendToSheet_(cfg, rowData) {
  try {
    var sh = SpreadsheetApp.openById(cfg.SPREADSHEET_ID).getSheetByName(cfg.SHEET_NAME);
    if (!sh) throw new Error('Лист не найден');
    sh.appendRow(rowData);
  } catch(e){ console.error('[APPEND] ' + e); }
}

// ---- Задачи по сделке ---------------------------------------------------
function getAmoTasks_(cfg, leadId) {
  try {
    var url = '/api/v4/tasks?filter[entity_id]=' + leadId + '&filter[entity_type]=leads';
    var resp = amoRequest_(cfg.AMO_HOST, cfg.AMO_TOKEN, url);
    var tasks = resp?._embedded?.tasks || [];
    return tasks.map(function(task){
      var typ = cfg.taskTypesMap[String(task.task_type_id)] || ('Неизвестный тип (' + task.task_type_id + ')');
      var due = formatTs_(task.complete_till);
      var st  = task.is_completed === true ? 'Выполнена'
              : (task.is_completed === false ? 'Не выполнена' : 'Статус не определен');
      return '• ' + typ + ': ' + (task.text || 'Нет текста') + '\n  Срок: ' + due + '\n  Статус: ' + st;
    }).join('\n') || 'Нет задач';
  } catch(e){ console.warn('[TASKS] ' + e); return 'Ошибка при получении задач'; }
}

// ---- Проверка: является ли заметка звонком (расширенная) ----------------
function _isCallNote_(note) {
  if (!note) return false;
  var noteType = String(note.note_type || '').toLowerCase();
  var params = note.params || {};
  
  // Стандартные типы звонков
  if (/^(call_in|call_out)$/i.test(noteType)) return true;
  
  // Гибридные звонки: common заметки с признаками звонка
  if (noteType === 'common') {
    var hasLink = params.link && /^https?:\/\//i.test(String(params.link));
    var hasDuration = params.duration && Number(params.duration) > 0;
    var hasCallStatus = params.call_status !== undefined;
    return hasLink || hasDuration || hasCallStatus;
  }
  
  return false;
}

// ---- Обработка одного события (строго только звонок) --------------------
function processEvent_(cfg, eventData) {
  try {
    var callNoteId = eventData && eventData.value_after && eventData.value_after[0] &&
                     eventData.value_after[0].note && eventData.value_after[0].note.id;
    if (!callNoteId) return;

    // Сразу отсечём, если этот звонок уже занесён в таблицу
    if (isCallProcessed_(cfg.SHEET_NAME, cfg.SPREADSHEET_ID, callNoteId)) return;

    // Определяем сделку, к которой привязан объект события
    var lead = null;
    var et = String(eventData.entity_type||'').toLowerCase();
    if (et === 'lead') {
      lead = getAmoLead_(cfg, eventData.entity_id);
    } else if (et === 'contact') {
      lead = getLeadByContact_(cfg, eventData.entity_id);
    } else if (et === 'company') {
      lead = getLeadByCompany_(cfg, eventData.entity_id);
    } else {
      return; // не наш тип сущности
    }
    if (!lead) return;

    // Берём все заметки сделки и ищем ровно нашу note по id (сравниваем как строки!)
    var notes = getAmoNotes_(cfg, 'leads', lead.id);
    var callNote = notes.find(function(n){ return String(n && n.id) === String(callNoteId); });

    // Обрабатываем только если note — звонок (расширенная проверка)
    if (!callNote || !_isCallNote_(callNote)) return;

    // Текст примечаний (для контекста)
    var tasksText = getAmoTasks_(cfg, lead.id);
    var notesText = (notes.map(function(note){
      var t = '';
      switch(note.note_type){
        case 'common':
          t = '📝 Примечание: ' + (note.params && note.params.text || '');
          if (note._embedded && note._embedded.attachments && note._embedded.attachments.length) {
            t += '\nattachments:\n' + note._embedded.attachments.map(function(a){
              return '• ' + a.name + ' (' + a.link + ')';
            }).join('\n');
          }
          break;
        case 'call_out':
        case 'call_in':
          var dur = (note.params && note.params.duration) ? (' (' + note.params.duration + 'с)') : '';
          var res = cfg.callStatusMap[String(note.params && note.params.call_status)] || 'Без результата';
          t = '📞 Звонок' + dur + ': ' + res;
          if (note.params && note.params.link) t += ' (' + note.params.link + ')';
          if (note._embedded && note._embedded.attachments && note._embedded.attachments.length) {
            t += '\nattachments:\n' + note._embedded.attachments.map(function(a){
              return '• ' + a.name + ' (' + a.link + ')';
            }).join('\n');
          }
          break;
        case 'task':
          t = '✅ Задача [ID:' + note.id + ']: ' + (note.params && note.params.text || '');
          t += '\nСтатус: ' + ((note.params && note.params.status) === 1 ? 'Выполнена' : 'Не выполнена');
          if (note.params && note.params.task_deadline) t += '\nСрок: ' + new Date(note.params.task_deadline*1000);
          break;
        case 'mail':
          t = '📧 Почта: ' + (note.params && note.params.subject || '');
          t += '\nОт: ' + (note.params && note.params.from || '');
          t += '\nСообщение: ' + (note.params && note.params.text || '');
          if (note.params && note.params.link) t += '\nСвязь: ' + note.params.link;
          break;
        case 'chat':
          t = '💬 Чат (' + (note.params && note.params.service || '') + '):';
          t += '\nСообщение: ' + (note.params && note.params.text || '');
          if (note.params && note.params.link) t += '\nСвязь: ' + note.params.link;
          break;
        default:
          t = 'Неизвестный тип: ' + note.note_type;
      }
      return t;
    }).join('\n').trim()) || 'Нет примечаний';

    notesText += '\n📌 Задачи в сделке:\n' + tasksText;

    var isOut = String(callNote.note_type||'').toLowerCase() === 'call_out';
    var row = [
      new Date((eventData.created_at||Date.now()/1000) * 1000),                        // 1 Время
      (callNote.params && callNote.params.link) || 'Нет записи',                       // 2 Ссылка
      lead.name,                                                                       // 3 Сделка
      getAmoPipeline_(cfg, lead.pipeline_id).name,                                     // 4 Воронка
      getAmoStatus_(cfg, lead.status_id, lead.pipeline_id).name,                       // 5 Статус
      lead.id,                                                                         // 6 ID сделки
      lead.pipeline_id,                                                                // 7 ID воронки
      lead.status_id,                                                                  // 8 ID статуса
      lead.responsible_user_id,                                                        // 9 ID ответственного
      getAmoUser_(cfg, lead.responsible_user_id).name,                                 // 10 Ответственный
      notesText,                                                                       // 11 Примечания + задачи
      (callNote.params && callNote.params.duration) || 0,                              // 12 Длительность
      cfg.callStatusMap[String(callNote.params && callNote.params.call_status)] || 'Без результата', // 13 Результат
      isOut ? 'Исходящий' : 'Входящий',                                                // 14 Тип
      callNoteId,                                                                       // 15 ID звонка (note.id)
      getAllEntityFields_(cfg, lead)                                                   // 16 Все поля
    ];

    appendToSheet_(cfg, row);
  } catch(e){
    console.error('[EVENT] ' + e);
  }
}

// ---- Основной синк (окно 5 минут; фильтр звонков в processEvent_) -------
function syncEventsToday() {
  var run = _beginSingletonRun_(SYNC_RUN_KEY, SYNC_WALL_MS);
  if (!run.ok) return;
  try {
    var cfg = getCfg_();
    var sh = SpreadsheetApp.openById(cfg.SPREADSHEET_ID).getSheetByName(cfg.SHEET_NAME);
    var from = Math.floor((Date.now() - 5*60*1000)/1000); // последние 5 минут
    var url = '/api/v4/events?filter[created_at][from]=' + from + '&limit=250';
    var batch = [];
    while (url) {
      var resp = amoRequest_(cfg.AMO_HOST, cfg.AMO_TOKEN, url);
      var chunk = resp && resp._embedded && resp._embedded.events ? resp._embedded.events : [];
      // Берём события, где есть value_after[].note.id и сущность такая, откуда можно выйти на сделку
      var useful = chunk.filter(function(e){
        if (!e || !e.value_after || !e.value_after[0] || !e.value_after[0].note || !e.value_after[0].note.id) return false;
        var et = String(e.entity_type||'').toLowerCase();
        return et==='lead' || et==='contact' || et==='company';
      });
      batch = batch.concat(useful);
      url = resp && resp._links && resp._links.next ? (resp._links.next.href || '') : '';
      if (batch.length > 3000) break; // предохранитель по времени выполнения
    }
    console.log('[SYNC] events_with_note=', batch.length);
    var processed = 0, written = 0;
    batch.forEach(function(ev){ 
      var before = sh ? sh.getLastRow() : 0;
      processEvent_(cfg, ev); 
      processed++;
      if (sh && sh.getLastRow() > before) written++;
    });
    console.log('[SYNC] processed=', processed, 'written=', written);
  } catch(e){
    console.error('[SYNC] ' + e);
  } finally {
    run.release();
  }
}

// ---- Разовый сбор за последние N часов ---------------------------------
function syncEventsLastHours(hours){
  hours = Number(hours)||3;
  var run = _beginSingletonRun_(SYNC_RUN_KEY + '_H' + hours, SYNC_WALL_MS);
  if (!run.ok) return;
  try {
    var cfg = getCfg_();
    var from = Math.floor((Date.now() - hours*3600*1000)/1000);
    _syncEventsFromTs_(cfg, from);
  } catch(e){ console.error('[SYNC H] ' + e); }
  finally { run.release(); }
}

// Вспомогательный сканер от отметки времени (сохранит только звонки)
function _syncEventsFromTs_(cfg, fromTs){
  var url = '/api/v4/events?filter[created_at][from]=' + fromTs + '&limit=250';
  var batch = [];
  while (url) {
    var resp = amoRequest_(cfg.AMO_HOST, cfg.AMO_TOKEN, url);
    var chunk = resp && resp._embedded && resp._embedded.events ? resp._embedded.events : [];
    var useful = chunk.filter(function(e){
      if (!e || !e.value_after || !e.value_after[0] || !e.value_after[0].note || !e.value_after[0].note.id) return false;
      var et = String(e.entity_type||'').toLowerCase();
      return et==='lead' || et==='contact' || et==='company';
    });
    batch = batch.concat(useful);
    url = resp && resp._links && resp._links.next ? (resp._links.next.href || '') : '';
    if (batch.length > 5000) break; // предохранитель
  }
  console.log('[SYNC H] events_with_note=', batch.length);
  var processed = 0, written = 0;
  var sh = SpreadsheetApp.openById(cfg.SPREADSHEET_ID).getSheetByName(cfg.SHEET_NAME);
  batch.forEach(function(ev){ 
    var before = sh ? sh.getLastRow() : 0;
    processEvent_(cfg, ev); 
    processed++;
    if (sh && sh.getLastRow() > before) written++;
  });
  console.log('[SYNC H] processed=', processed, 'written=', written);
}