// ============================== Заметки.gs ==============================
// Отправляет V (транскрипт) и W (оценка) в amoCRM как text notes.
// Логирование в LOG_NOTES + небольшой dry-run дианостик.

const NOTES_TIME_BUDGET_MS   = 330000;
const MAX_ROWS_PER_RUN       = 25;
const NOTES_PAGES_MAX        = 2;
const DUP_CHECK_RECENT_LIMIT = 120;
const PAUSE_BETWEEN_OPS_MS   = 120;
const AMO_RETRY_TRIES        = 2;
const AMO_RETRY_BASE_MS      = 400;
const ENABLE_LOG_SHEET       = true;

function sendNotesToAmoCRM() {
  const t0 = Date.now();
  const stats = { rows_seen:0, rows_done:0, notes_sent:0, notes_dup:0, rows_error:0 };

  try {
    const cfg = getCfg_Notes_();
    const sh  = SpreadsheetApp.getActive().getSheetByName(cfg.SHEET_NAME);
    if (!sh) throw new Error(`Лист "${cfg.SHEET_NAME}" не найден`);
    const last = sh.getLastRow();
    if (last < 2) { _logSummary_(stats, t0); return; }

    const cols = Math.max(27, Math.min(60, sh.getLastColumn())); // ← чуть шире для надёжности
    const data = sh.getRange(2,1,last-1,cols).getValues();

    const queue = [];
    for (let i=0; i<data.length && queue.length < MAX_ROWS_PER_RUN; i++) {
      const R = i+2, r = data[i];
      const dealId    = r[5];
      const transcript= r[21];
      const evalText  = r[22];
      const aa        = String(r[26]||'').trim();

      if (!dealId) continue;
      if (!(transcript || evalText)) continue;
      if (aa === 'OK') continue;

      queue.push({ R, dealId, transcript:String(transcript||'').trim(), assessment:String(evalText||'').trim() });
    }
    stats.rows_seen = queue.length;
    if (!queue.length) { _logSummary_(stats, t0); return; }

    const grouped = new Map();
    for (const it of queue) {
      if (!grouped.has(it.dealId)) grouped.set(it.dealId, []);
      grouped.get(it.dealId).push(it);
    }

    for (const [dealId, items] of grouped.entries()) {
      if (Date.now() - t0 > NOTES_TIME_BUDGET_MS) { console.log('[NOTES] time budget reached'); break; }

      let existing = { set:new Set(), raw:[] };
      try { existing = getLeadCommonNotesTexts_Limited_(cfg, dealId, NOTES_PAGES_MAX, DUP_CHECK_RECENT_LIMIT); }
      catch(e){ console.warn('[NOTES] prefetch fail for ' + dealId + ': ' + String(e)); }

      for (const it of items) {
        if (Date.now() - t0 > NOTES_TIME_BUDGET_MS) break;

        const { R, transcript, assessment } = it;
        try {
          const resV = transcript ? _sendNoteIfNeed_(cfg, dealId, transcript, existing) : 'SKIP';
          const resW = assessment ? _sendNoteIfNeed_(cfg, dealId, assessment, existing) : 'SKIP';

          if (resV === 'OK') stats.notes_sent++; else if (resV === 'DUPLICATE') stats.notes_dup++;
          if (resW === 'OK') stats.notes_sent++; else if (resW === 'DUPLICATE') stats.notes_dup++;

          _logRow_(dealId, R, 'V', resV);
          _logRow_(dealId, R, 'W', resW);

          const ok = [resV,resW].every(s => s === 'OK' || s === 'DUPLICATE' || s === 'SKIP');
          if (ok) {
            sh.getRange(R, 26).setValue(new Date()); // Z
            sh.getRange(R, 27).setValue('OK');       // AA
            stats.rows_done++;
          } else {
            sh.getRange(R, 27).setValue('AMO_NOTE_ERROR');
            stats.rows_error++;
          }
        } catch(e){
          sh.getRange(R, 27).setValue('AMO_NOTE_ERROR');
          stats.rows_error++;
          _logRow_(dealId, R, '-', 'ERROR', String(e));
        }
        Utilities.sleep(PAUSE_BETWEEN_OPS_MS);
      }
    }
  } catch(e){ console.error('sendNotesToAmoCRM error: ' + String(e)); }
  finally { _logSummary_(stats, t0); }
}

// ---- Helpers ------------------------------------------------------------

function getCfg_Notes_() {
  const db = SpreadsheetApp.getActive().getSheetByName('БД');
  if (!db) throw new Error('Лист "БД" не найден');
  const host = String(db.getRange('B3').getValue() || '').trim().replace(/^https?:\/\//i,'').replace(/\/+$/,'');
  const amoHost = /\.amocrm\.(ru|eu)$/i.test(host) ? host.toLowerCase() : (host + '.amocrm.ru').toLowerCase();
  const sheetName = String(db.getRange('B4').getValue() || '').trim();
  if (!sheetName) throw new Error('SHEET_NAME пуст в БД!B4');
  return {
    AMO_TOKEN : String(db.getRange('B2').getValue() || '').trim(),
    AMO_HOST  : amoHost,
    SHEET_NAME: sheetName
  };
}

function _normalize_(s) {
  return String(s||'').replace(/\r\n/g,'\n').replace(/\r/g,'\n').replace(/[ \t]+/g,' ').replace(/\u00A0/g,' ').trim();
}

function getLeadCommonNotesTexts_Limited_(cfg, dealId, pagesLimit, recentLimit) {
  const keep = Math.max(1, Number(recentLimit||100));
  const maxPages = Math.max(1, Number(pagesLimit||1));
  let notes=[], url = `https://${cfg.AMO_HOST}/api/v4/leads/${dealId}/notes?with=attachments&limit=250`, pages=0;

  while (url && pages < maxPages) {
    const resp = UrlFetchApp.fetch(url, { method:'get', headers:{'Authorization':'Bearer '+cfg.AMO_TOKEN}, muteHttpExceptions:true, followRedirects:true });
    const code = resp.getResponseCode(); if (code<200||code>=300) break;
    const json = JSON.parse(resp.getContentText() || '{}');
    notes = notes.concat(json?._embedded?.notes || []);
    url = json?._links?.next?.href || ''; pages++;
  }

  const texts=[];
  for (const n of notes) {
    if (n.note_type === 'common') {
      const p = n.params || {};
      const t = (typeof p.text==='string' && p.text.trim()) ? p.text.trim() : (typeof n.text==='string' && n.text.trim() ? n.text.trim() : '');
      if (t) texts.push(t);
    }
    if (texts.length >= keep) break;
  }
  return { set: new Set(texts.map(_normalize_)), raw: texts };
}

function _sendNoteIfNeed_(cfg, dealId, text, existing) {
  const norm = _normalize_(text);
  if (existing.set.has(norm)) return 'DUPLICATE';
  const res = addNoteToAmo_(cfg, dealId, text);
  if (res === 'OK') existing.set.add(norm);
  return res;
}

function addNoteToAmo_(cfg, dealId, noteText) {
  const text = String(noteText || '');
  const body1 = [{ entity_id:Number(dealId), note_type:'common', params:{ text:text } }];
  const body2 = [{ note_type:'common', params:{ text:text } }];
  const headers = { 'Authorization':'Bearer '+cfg.AMO_TOKEN, 'Content-Type':'application/json', 'Accept':'application/hal+json' };

  return _withBackoff_(function(){
    // путь 1: коллекционный
    let resp = UrlFetchApp.fetch(`https://${cfg.AMO_HOST}/api/v4/leads/notes`, { method:'post', headers, muteHttpExceptions:true, followRedirects:true, payload:JSON.stringify(body1) });
    let code = resp.getResponseCode();
    if (code === 400) {
      const txt = resp.getContentText() || '';
      if (/code"\s*:\s*226|Error 226/i.test(txt)) return 'DUPLICATE';
      // fallback путь 2: по сделке
      resp = UrlFetchApp.fetch(`https://${cfg.AMO_HOST}/api/v4/leads/${dealId}/notes`, { method:'post', headers, muteHttpExceptions:true, followRedirects:true, payload:JSON.stringify(body2) });
      code = resp.getResponseCode();
      if (code === 400) {
        const txt2 = resp.getContentText() || '';
        if (/code"\s*:\s*226|Error 226/i.test(txt2)) return 'DUPLICATE';
        if (/length|too.*long/i.test(txt2)) { body2[0].params.text = text.slice(0,7999); resp = UrlFetchApp.fetch(`https://${cfg.AMO_HOST}/api/v4/leads/${dealId}/notes`, { method:'post', headers, muteHttpExceptions:true, followRedirects:true, payload:JSON.stringify(body2) }); code = resp.getResponseCode(); }
        if (code < 200 || code >= 300) throw new Error('amoCRM error ' + code + ': ' + (resp.getContentText()||'').slice(0,600));
        return 'OK';
      }
      if (code < 200 || code >= 300) throw new Error('amoCRM error ' + code + ': ' + (resp.getContentText()||'').slice(0,600));
      return 'OK';
    }
    if (code < 200 || code >= 300) throw new Error('amoCRM error ' + code + ': ' + (resp.getContentText()||'').slice(0,600));
    return 'OK';
  }, { tries: AMO_RETRY_TRIES, baseMs: AMO_RETRY_BASE_MS });
}

function _withBackoff_(fn, opt) {
  const tries = (opt && opt.tries) || 2;
  const base  = (opt && opt.baseMs) || 400;
  let a=0; for (; a<=tries; a++) {
    try { return fn(); } catch(e){ if (a>=tries) throw e; Utilities.sleep(base * Math.pow(2,a) + Math.floor(Math.random()*120)); }
  }
}

function _ensureLogSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('LOG_NOTES');
  if (!sh) { sh = ss.insertSheet('LOG_NOTES'); sh.getRange(1,1,1,5).setValues([['time','dealId','row','col','result']]); }
  return sh;
}
function _logRow_(dealId,row,col,res,extra){ if(!ENABLE_LOG_SHEET) return; _ensureLogSheet_().appendRow([new Date(), String(dealId), String(row), String(col), String(res||'') + (extra?(' | '+extra):'')]); }
function _logSummary_(stats, t0){ const ms=Date.now()-t0; const msg=`[NOTES] done in ${ms}ms: rows_seen=${stats.rows_seen}, rows_done=${stats.rows_done}, notes_sent=${stats.notes_sent}, dup=${stats.notes_dup}, errors=${stats.rows_error}`; console.log(msg); if(ENABLE_LOG_SHEET){ _ensureLogSheet_().appendRow([new Date(),'SUMMARY','','',msg]); } }

// ---- Диагностика причины «rows_seen=0» (ничего не шлёт) ----------------
function notes_DryRunWhyZero(){
  const cfg = getCfg_Notes_();
  const sh  = SpreadsheetApp.getActive().getSheetByName(cfg.SHEET_NAME);
  if (!sh) { console.log('Нет листа', cfg.SHEET_NAME); return; }
  const last = sh.getLastRow();
  if (last < 2) { console.log('Лист пуст'); return; }

  const cols = Math.max(27, Math.min(60, sh.getLastColumn()));
  const data = sh.getRange(2,1,last-1,cols).getValues();
  let seen=0, ok=0;
  for (let i=0;i<data.length;i++){
    const R=i+2, r=data[i];
    const dealId=r[5], V=r[21], W=r[22], AA=String(r[26]||'').trim();
    if (dealId) seen++;
    const why=[];
    if (!dealId) why.push('нет dealId(6)');
    if (!(V||W)) why.push('нет V/W(22/23)');
    if (AA==='OK') why.push('AA=OK');
    if (!why.length) { ok++; console.log('OK-кандидат R=',R,'dealId=',dealId,'V=',!!V,'W=',!!W); }
    else console.log('SKIP R=',R,'→', why.join('; '));
  }
  console.log('Итог: dealId-строк=',seen,'; кандидатов к отправке=',ok);
}