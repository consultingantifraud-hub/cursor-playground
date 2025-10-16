/** ================= Оценка.gs (стабильный воркер) =======================
 * Делает РОВНО одно действие за запуск (V или W).
 * Не трогаем триггеры. Добавлен мягкий логгер в лист LOGS.
 */

/* -------- Константы колонок (1-based) -------- */
var COL_B_LINK           = 2;   // B — ссылка на запись
var COL_J_RESPONSIBLE    = 10;  // J — ответственный
var COL_L_DURATION       = 12;  // L — длительность (сек)
var COL_P_ALLFIELDS      = 16;  // P — все поля (ищем STOP BOT)
var COL_V_TRANSCRIPT     = 22;  // V — транскрипция
var COL_W_EVAL           = 23;  // W — оценка
var COL_Y_STARTED_AT     = 25;  // Y — отметка старта оценки

/* -------- Синглтон -------- */
var RUN_KEY  = 'RUN_TO_PROTALK_BOT';
var RUN_WALL_MS = 20000;

/* -------- Настройки -------- */
var SAFE_EVAL = 'Статус: мало данных/короткий звонок. Следующий шаг: перезвонить клиенту сегодня и уточнить запрос и следующий контакт.';
var MAX_ROWS_SCAN = 400;

/* -------- Чтение конфигурации -------- */
function _cfg() {
  var ss = SpreadsheetApp.getActive();
  var db = ss.getSheetByName('БД');
  if (!db) throw new Error('Лист "БД" не найден');

  function N(cell, def){ var v = Number(db.getRange(cell).getValue()); return (isFinite(v)&&v>0)?v:def; }
  function S(cell){ return String(db.getRange(cell).getValue()||'').trim(); }

  var targetName = S('B4');
  var sh = ss.getSheetByName(targetName);
  if (!sh) throw new Error("Лист '" + targetName + "' не найден");

  var cfg = {
    SHEET:          sh,
    BOT_TOKEN:      S('B5'),
    BOT_ID:         S('B6'),
    PROMPT_TRANS:   S('H5'),
    PROMPT_EVAL:    S('I5'),
    MIN_SEC_TRANS:  N('H2', 60),
    MIN_SEC_EVAL:   N('I2', 60),
    MAX_TEXT_CHARS: N('K2', 29000)
  };
  if (!cfg.BOT_TOKEN || !cfg.BOT_ID) throw new Error('Пустые BOT_TOKEN/BOT_ID в «БД»');
  return cfg;
}

/* -------- Утилиты -------- */
function _hasStopBot(s){ return /\bstop\s*bot\s*:\s*true\b/i.test(String(s||'')); }
function _clip(s, max){ var t=String(s||''); return (max && t.length>max)?t.slice(0,max):t; }
function _isEmptyV(v){ var s=String(v||'').trim(); return !s || /^нужен\s*текст$/i.test(s) || /^без\s*текста$/i.test(s) || /^статус:\s*нужна\s*оценка/i.test(s); }
function _isEmptyW(w){ var s=String(w||'').trim(); return !s || /^нужна\s*оценка$/i.test(s) || /^без\s*оценки$/i.test(s); }
function _findUrlInRow(sh,row){
  var direct = String(sh.getRange(row, COL_B_LINK).getValue()||'').trim();
  if (/^https?:\/\//i.test(direct)) return direct;
  var lastCol = sh.getLastColumn();
  var vals = sh.getRange(row, 1, 1, lastCol).getValues()[0];
  for (var c=0;c<vals.length;c++){
    var cell = String(vals[c]||'');
    var m = cell.match(/https?:\/\/[^\s)"'>]+/i);
    if (m) return m[0];
  }
  return '';
}

/* -------- Логгер в LOGS -------- */
function _ensureLogsSheet_(){
  var ss=SpreadsheetApp.getActive(), sh=ss.getSheetByName('LOGS');
  if(!sh){ 
    try {
      sh = ss.insertSheet('LOGS'); 
      sh.appendRow(['time','chat_id','phase','text_preview','provider','bot_id','model','req_tokens','resp_tokens','latency_ms','error','endpoint']); 
      console.log('[LOGS] sheet created');
    } catch(e) {
      console.error('[LOGS] failed to create sheet:', String(e));
      return null;
    }
  }
  return sh;
}
function _logProtalk_(row){
  try{
    var sh = _ensureLogsSheet_();
    if (!sh) {
      console.warn('[LOGS] no sheet available');
      return;
    }
    sh.appendRow([ new Date(),
      String(row.chat_id||''), String(row.phase||''), String((row.text||'').slice(0,140)),
      String(row.provider||'Replica'), String(row.bot_id||''), String(row.model||''),
      String(row.req_tokens||''), String(row.resp_tokens||''), String(row.latency_ms||''),
      String(row.error||''), String(row.endpoint||'') ]);
    console.log('[LOGS] logged:', row.phase, 'for chat', row.chat_id);
  }catch(e){ console.warn('[LOGS] append error', String(e)); }
}

/* -------- Доставание текста из разных форм ответов ProTalk -------- */
function _pickDoneText(json){
  if (!json) return '';
  if (json.done && typeof json.done==='string' && json.done.trim()) return json.done.trim();
  if (json.text && typeof json.text==='string' && json.text.trim()) return json.text.trim();
  if (json.result && typeof json.result==='string' && json.result.trim()) return json.result.trim();
  if (json.result && json.result.text && typeof json.result.text==='string') return json.result.text.trim();
  if (json.message && typeof json.message==='string') return json.message.trim();
  return '';
}

/* -------- Вызов ProTalk с ретраями (использует protalkFetchAsk_ из QA) -------- */
function _askWithRetry(token, payload){
  var start = Date.now();
  var tries = 0;
  var chatId = (payload && payload.chat_id) || ('chat_'+Date.now());
  _logProtalk_({ chat_id: chatId, phase: 'STARTED', text: (payload && payload.message)||'', endpoint: protalkAskUrl_(token) });

  while (tries<=1){
    tries++;
    try{
      var resp = protalkFetchAsk_(token, payload, {
        method:'post',
        contentType:'application/json; charset=utf-8',
        followRedirects:true,
        muteHttpExceptions:true,
        headers:{'X-ProTalk-Client':'amo-gas/unified/2.0'}
      });
      var code = resp.getResponseCode();
      var body = resp.getContentText() || '';
      if (code===200){
        try{
          var j = JSON.parse(body);
          var txt = _pickDoneText(j);
          _logProtalk_({ chat_id: chatId, phase: 'DONE', text: txt||'', provider:'Replica',
                         model:(j && (j.model||j.engine||'')), latency_ms:(Date.now()-start),
                         endpoint: protalkAskUrl_(token) });
          return (txt && txt.trim()) ? j : null;
        }catch(e){
          _logProtalk_({ chat_id: chatId, phase: 'ERROR_JSON', error:String(e), endpoint: protalkAskUrl_(token) });
          return null;
        }
      }
      _logProtalk_({ chat_id: chatId, phase: 'HTTP_'+code, error: body.slice(0,180), endpoint: protalkAskUrl_(token) });
      return null;
    }catch(e){
      _logProtalk_({ chat_id: chatId, phase: 'THROW', error:String(e), endpoint: protalkAskUrl_(token) });
      if (Date.now()-start>30000) return null;
      Utilities.sleep(800);
    }
  }
  return null;
}

/* -------- Нормализация бренда (только для оценки) -------- */
var _COMPANY_RE = /(сту[дт][ие]тал[ияеьй]|студитали|сто\s*детал[еёеиы][йи]?|100\s*детал[еёеиы][йи]?)/gi;
function _normalizeBrandEval(t){
  if (!t) return t;
  return String(t)
    .replace(_COMPANY_RE,'Сто Деталей')
    .replace(/\bКомпания\s+Сто\s+Детал[еёеи][йи]?\b/gi,'Компания «Сто Деталей»');
}

/* -------- Синглтон -------- */
function _beginRun(){
  var props = PropertiesService.getScriptProperties();
  var now = Date.now();
  var raw = props.getProperty(RUN_KEY);
  if (raw){
    var ts = Number(raw)||0;
    if (ts && now-ts < RUN_WALL_MS){
      console.log('[BOT] skip: already running');
      return {ok:false, release:function(){}};
    }
  }
  props.setProperty(RUN_KEY, String(now));
  return {ok:true, release:function(){ try{ PropertiesService.getScriptProperties().deleteProperty(RUN_KEY); }catch(e){} }};
}

/* ================================= Основной воркер ================================= */
function ToProTalkBot(){
  var run = _beginRun(); if (!run.ok) return;
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)){ console.log('[BOT] busy by another run'); run.release(); return; }

  var summary = {scanned:0, skipped:0, transcribed:0, evaluated:0};

  try{
    var cfg = _cfg();
    var sh  = cfg.SHEET;
    var lastRow = sh.getLastRow();
    if (lastRow<2){ console.log('[BOT] sheet empty'); return; }

    var width = Math.max(COL_P_ALLFIELDS, sh.getLastColumn());
    var rowsToScan = Math.min(MAX_ROWS_SCAN, lastRow-1);
    var values = sh.getRange(2,1,rowsToScan,width).getValues(); // со 2-й строки

    // ---------- Шаг A: транскрипция ----------
    for (var i=0;i<values.length;i++){
      summary.scanned++;
      var row = i+2;

      var V = String(values[i][COL_V_TRANSCRIPT-1]||'').trim();
      if (!_isEmptyV(V)){ summary.skipped++; continue; }

      var all = values[i][COL_P_ALLFIELDS-1];
      if (_hasStopBot(all)){ summary.skipped++; continue; }

      var link = _findUrlInRow(sh,row);
      if (!/^https?:\/\//i.test(link)){ summary.skipped++; continue; }

      var sec = Number(values[i][COL_L_DURATION-1]||0);
      if (!(sec>=cfg.MIN_SEC_TRANS)){ summary.skipped++; continue; }

      var msg = cfg.PROMPT_TRANS ? (cfg.PROMPT_TRANS + '\n\n' + link) : link;
      var json = _askWithRetry(cfg.BOT_TOKEN, { bot_id: cfg.BOT_ID, chat_id: 'trans_'+Date.now(), message: msg });
      var out  = _clip(_pickDoneText(json) || 'СТАТУС: Нужна оценка (нет транскрипции)', cfg.MAX_TEXT_CHARS);

      sh.getRange(row, COL_V_TRANSCRIPT).setValue(out);
      SpreadsheetApp.flush();

      summary.transcribed = 1;
      console.log('[BOT] V written at row', row, 'len=', out.length);
      return; // одна операция за запуск
    }

    // ---------- Шаг Б: оценка ----------
    for (var j=0;j<values.length;j++){
      var row2 = j+2;

      var v = String(values[j][COL_V_TRANSCRIPT-1]||'').trim();
      var w = String(values[j][COL_W_EVAL-1]||'').trim();
      if (_isEmptyV(v) || !_isEmptyW(w)) { continue; }

      var all2 = values[j][COL_P_ALLFIELDS-1];
      if (_hasStopBot(all2)) { continue; }

      var sec2 = Number(values[j][COL_L_DURATION-1]||0);
      if (!(sec2>=cfg.MIN_SEC_EVAL)) { continue; }

      if (/^статус:\s*нужна\s*оценка/i.test(v)){
        sh.getRange(row2, COL_W_EVAL).setValue(SAFE_EVAL);
        SpreadsheetApp.flush();
        summary.evaluated = 1;
        console.log('[BOT] W SAFE by status at row', row2);
        return;
      }

      var cleanLen = v.replace(/\[неразборчиво\]/gi,'').replace(/\s+/g,' ').trim().length;
      if (cleanLen < 120){
        sh.getRange(row2, COL_W_EVAL).setValue(SAFE_EVAL);
        SpreadsheetApp.flush();
        summary.evaluated = 1;
        console.log('[BOT] W SAFE by short V at row', row2, 'len=', cleanLen);
        return;
      }

      try{ sh.getRange(row2, COL_Y_STARTED_AT).setValue(new Date()); }catch(e){}

      var prompt = (cfg.PROMPT_EVAL ? (cfg.PROMPT_EVAL.trim() + '\n\n') : '') + 'Транскрипция (без правок):\n' + v;
      var json2 = _askWithRetry(cfg.BOT_TOKEN, { bot_id: cfg.BOT_ID, chat_id: 'score_'+Date.now(), message: prompt });
      var out2  = _pickDoneText(json2);
      if (!out2 || /^статус\s*:/i.test(out2)) out2 = SAFE_EVAL;
      out2 = _clip(String(out2).replace(/\*/g,''), cfg.MAX_TEXT_CHARS);
      out2 = _normalizeBrandEval(out2);

      sh.getRange(row2, COL_W_EVAL).setValue(out2);
      SpreadsheetApp.flush();

      summary.evaluated = 1;
      console.log('[BOT] W written at row', row2, 'len=', out2.length);
      return;
    }

    console.log('[BOT] nothing to do; scanned=', summary.scanned, 'skipped=', summary.skipped);

  } catch(e){
    console.error('[BOT] ERROR:', String(e));
  } finally {
    try{ lock.releaseLock(); }catch(_){}
    run.release();
  }
}

/* -------- Совместимость старого имени -------- */
function chatToProTalkBot(){ return ToProTalkBot(); }

/* -------- Быстрые ручные раннеры -------- */
function qa_RunActiveRowForce(){
  var cfg=_cfg(), sh=cfg.SHEET, a=sh.getActiveRange(); if (!a) { console.log('[ROW] нет активной строки'); return; }
  qa_RunRowForce(a.getRow());
}
function qa_RunRowForce(row){ qa_RunRow(row, true); }
function qa_RunRow(row, force){
  row = Number(row); if (!row || row<=1){ console.log('[ROW] укажи номер строки (>1)'); return; }

  var cfg=_cfg(), sh=cfg.SHEET;
  var link = _findUrlInRow(sh,row);
  var sec  = Number(sh.getRange(row, COL_L_DURATION).getValue()||0);
  var all  = String(sh.getRange(row, COL_P_ALLFIELDS).getValue()||'');

  if (!force && _hasStopBot(all)){ console.log('[ROW] STOP BOT'); return; }
  if (!/^https?:\/\//i.test(link)){ console.log('[ROW] нет валидного URL'); return; }

  var V = String(sh.getRange(row, COL_V_TRANSCRIPT).getValue()||'').trim();
  var W = String(sh.getRange(row, COL_W_EVAL).getValue()||'').trim();

  if (_isEmptyV(V)){
    if (!force && !(sec>=cfg.MIN_SEC_TRANS)){ console.log('[ROW] сек < MIN_SEC_TRANS'); return; }
    var msg = cfg.PROMPT_TRANS ? (cfg.PROMPT_TRANS + '\n\n' + link) : link;
    var json = _askWithRetry(cfg.BOT_TOKEN, { bot_id: cfg.BOT_ID, chat_id:'trans_row_'+row+'_'+Date.now(), message: msg });
    var out  = _clip(_pickDoneText(json) || 'СТАТУС: Нужна оценка (нет транскрипции)', cfg.MAX_TEXT_CHARS);
    sh.getRange(row, COL_V_TRANSCRIPT).setValue(out); SpreadsheetApp.flush();
    console.log('[ROW] V written');
    return;
  }

  if (_isEmptyW(W)){
    if (!force && !(sec>=cfg.MIN_SEC_EVAL)){ console.log('[ROW] сек < MIN_SEC_EVAL'); return; }
    if (/^статус:\s*нужна\s*оценка/i.test(V)){ sh.getRange(row, COL_W_EVAL).setValue(SAFE_EVAL); SpreadsheetApp.flush(); console.log('[ROW] W SAFE'); return; }
    var cleanLen = V.replace(/\[неразборчиво\]/gi,'').replace(/\s+/g,' ').trim().length;
    if (cleanLen < 120){ sh.getRange(row, COL_W_EVAL).setValue(SAFE_EVAL); SpreadsheetApp.flush(); console.log('[ROW] W SAFE short'); return; }
    try{ sh.getRange(row, COL_Y_STARTED_AT).setValue(new Date()); }catch(e){}
    var prompt = (cfg.PROMPT_EVAL ? (cfg.PROMPT_EVAL.trim()+'\n\n'):'') + 'Транскрипция (без правок):\n' + V;
    var json2 = _askWithRetry(cfg.BOT_TOKEN, { bot_id: cfg.BOT_ID, chat_id:'score_row_'+row+'_'+Date.now(), message: prompt });
    var out2  = _pickDoneText(json2);
    if (!out2 || /^статус\s*:/i.test(out2)) out2 = SAFE_EVAL;
    out2 = _normalizeBrandEval(_clip(String(out2).replace(/\*/g,''), cfg.MAX_TEXT_CHARS));
    sh.getRange(row, COL_W_EVAL).setValue(out2); SpreadsheetApp.flush();
    console.log('[ROW] W written');
    return;
  }

  console.log('[ROW] ничего не делаем — V и W уже есть');
}