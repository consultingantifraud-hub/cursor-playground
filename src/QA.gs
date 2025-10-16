/** ================= QA SVC PATCHES =================
 * Версия: 2025-10-16.r3
 * Назначение: безопасные обвязки ProTalk + QA-проверки аудио URL.
 * ВАЖНО: это единственные рабочие функции getProtalkBase_/protalkAskUrl_/protalkFetchAsk_.
 */

/* -------------------------- Статусы/лог -------------------------- */
function QA_STATUS_NO_TRANSCRIPTION_() { return 'СТАТУС: Нужна оценка (нет транскрипции)'; }
function _qaLog_(lvl, msg, extra) {
  var tag = '[QA-SVC]';
  try {
    if (lvl === 'warn') console.warn(tag, msg, extra || '');
    else if (lvl === 'error') console.error(tag, msg, extra || '');
    else console.log(tag, msg, extra || '');
  } catch (_) {}
}
function getQaPatchesVersion_() { return '2025-10-16.r3'; }

/* -------------------------- Колонки/листы ------------------------ */
function _col_(name, fallback) {
  try { if (typeof this[name] === 'number' && isFinite(this[name]) && this[name] > 0) return this[name]; }
  catch (_) {}
  return fallback;
}
function _colLink_()     { return _col_('COL_B_LINK', 2); }
function _colDuration_() { return _col_('COL_L_DURATION', 12); }
function _getTargetSheet_() {
  var ss = SpreadsheetApp.getActive();
  var db = ss.getSheetByName('БД');
  if (!db) throw new Error('Лист "БД" не найден');
  var target = String(db.getRange('B4').getValue() || '').trim();
  var sh = target ? ss.getSheetByName(target) : null;
  if (!sh) throw new Error('Целевой лист не найден: ' + target);
  return sh;
}

/* -------------------------- ProTalk ------------------------------ */
function getProtalkBase_() { return 'https://eu1.api.pro-talk.ru'; } // единая точка
function protalkAskUrl_(botToken) { return getProtalkBase_() + '/api/v1.0/ask/' + encodeURIComponent(String(botToken || '')); }
function patchProtalkAskUrl_(ignoredOld, botToken) { return protalkAskUrl_(botToken); }

function protalkFetchAsk_(botToken, payload, fetchOpts) {
  var url = protalkAskUrl_(botToken);
  var options = Object.assign({
    method: 'post',
    contentType: 'application/json; charset=utf-8',
    muteHttpExceptions: true,
    followRedirects: true,
    payload: JSON.stringify(payload || {})
  }, fetchOpts || {});
  _qaLog_('info', 'protalkFetchAsk_()', { url: url, hasPayload: !!payload });

  var resp = UrlFetchApp.fetch(url, options);
  if (resp.getResponseCode() !== 200) {
    try {
      var alt = 'https://us1.api.pro-talk.ru/api/v1.0/ask/' + encodeURIComponent(String(botToken || ''));
      _qaLog_('warn', 'protalkFetchAsk_: EU1 non-200, fallback to US1', resp.getResponseCode());
      resp = UrlFetchApp.fetch(alt, options);
    } catch (_e) {}
  }
  return resp;
}

function protalkFetchAskJson_(botToken, payload, fetchOpts) {
  var resp = protalkFetchAsk_(botToken, payload, fetchOpts);
  var code = resp.getResponseCode();
  var body = resp.getContentText() || '';
  if (code !== 200) {
    _qaLog_('error', 'protalkFetchAskJson_: HTTP ' + code, body.slice(0, 600));
    throw new Error('Protalk ASK HTTP ' + code);
  }
  try { return JSON.parse(body); }
  catch (e) {
    _qaLog_('error', 'protalkFetchAskJson_: JSON parse error', body.slice(0, 600));
    throw e;
  }
}

/* -------------------------- Валидация/гард ----------------------- */
function validateAudioUrl_(url) {
  var res = { ok: false, reason: '', url: '' };
  try {
    if (!url || typeof url !== 'string') { res.reason = 'empty'; return res; }
    var u = url.trim();
    if (!u || u.toLowerCase() === 'нет записи') { res.reason = 'no_record'; return res; }
    if (!/^https?:\/\//i.test(u)) { res.reason = 'not_http_https'; return res; }
    if (/URL_аудио_файла_здесь/i.test(u) || /<.+>/.test(u)) { res.reason = 'placeholder'; return res; }

    var knownHosts = ['amocrm.mango-office.ru','mango-office.ru','mango-office.com','amocrm.ru'];
    var hostMatch = u.match(/^https?:\/\/([^\/]+)/i);
    var host = (hostMatch ? hostMatch[1] : '').toLowerCase();
    if (host && !knownHosts.some(function(h){ return host.indexOf(h) !== -1; })) {
      _qaLog_('warn', 'validateAudioUrl_: unknown host (не блокируем)', host);
    }

    res.ok = true; res.url = u; return res;
  } catch (e) { res.reason = 'exception:' + String(e); return res; }
}

function guardTranscribe_(audioUrl, transcribeFnName, opts) {
  var options = Object.assign({ retryCount: 1, retryDelayMs: 2500, minChars: 12 }, opts || {});
  var v = validateAudioUrl_(audioUrl);
  if (!v.ok) { _qaLog_('warn', 'guardTranscribe_: invalid audio URL', v); return QA_STATUS_NO_TRANSCRIPTION_(); }

  var fn = (typeof transcribeFnName === 'function') ? transcribeFnName : (this[transcribeFnName] || null);
  if (typeof fn !== 'function') { _qaLog_('error', 'guardTranscribe_: function not found', transcribeFnName); return QA_STATUS_NO_TRANSCRIPTION_(); }

  var attempt = 0;
  while (attempt <= options.retryCount) {
    try {
      attempt++;
      var text = String(fn(v.url) || '').trim();
      if (text && text.length >= options.minChars && !/^статус:/i.test(text)) return text;
      _qaLog_('warn', 'guardTranscribe_: empty/short/status-like text', { attempt: attempt, len: (text || '').length });
    } catch (e) {
      _qaLog_('warn', 'guardTranscribe_: transcribe threw', { attempt: attempt, err: String(e) });
    }
    if (attempt <= options.retryCount) Utilities.sleep(options.retryDelayMs);
  }
  return QA_STATUS_NO_TRANSCRIPTION_();
}

/* -------------------------- Пинг аудио --------------------------- */
function debugPingAudio_(url) {
  try {
    if (!url) { console.log('[PING] no url'); return; }
    var res = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { 'Range': 'bytes=0-0' },
      followRedirects: true,
      muteHttpExceptions: true
    });
    var headers = res.getHeaders() || {};
    var ct = headers['Content-Type'] || headers['content-type'] || '';
    var cl = headers['Content-Length'] || headers['content-length'] || '';
    console.log('[PING] code =', res.getResponseCode(), 'ctype =', ct, 'len(header) =', cl);
  } catch (e) { console.warn('[PING] exception', String(e)); }
}
function _findFirstAudioUrlInRow_(sh, row) {
  var colLink = _colLink_();
  var direct = String(sh.getRange(row, colLink).getValue() || '').trim();
  if (/^https?:\/\//i.test(direct)) return direct;

  var lastCol = Math.max(sh.getLastColumn(), colLink);
  var vals = sh.getRange(row, 1, 1, lastCol).getValues()[0].map(function(x){ return String(x||''); });
  var re = /(https?:\/\/[^\s)"'>]+)/ig;
  for (var i=0; i<vals.length; i++) {
    var cell = vals[i];
    var m;
    while ((m = re.exec(cell)) !== null) {
      if (m[1]) return m[1];
    }
  }
  return '';
}
function debugPingFirstAudioFromSheet_() {
  try {
    var sh = _getTargetSheet_();
    var linkCol = _colLink_();
    var durCol  = _colDuration_();

    var last = sh.getLastRow();
    if (last < 2) { console.log('[PING] sheet empty'); return; }

    var data = sh.getRange(2, 1, last-1, Math.max(sh.getLastColumn(), linkCol)).getValues();
    for (var i=0; i<data.length; i++) {
      var rowIdx = i + 2;
      var url = String(sh.getRange(rowIdx, linkCol).getValue() || '').trim();
      var sec = Number(sh.getRange(rowIdx, durCol).getValue() || 0);
      if (!/^https?:\/\//i.test(url)) url = _findFirstAudioUrlInRow_(sh, rowIdx);
      if (url && /^https?:\/\//i.test(url) && sec > 0) {
        console.log('[PING] Row', rowIdx, 'secs', sec, 'url', url);
        debugPingAudio_(url);
        break;
      }
    }
  } catch (e) { console.warn('[PING] sheet exception', String(e)); }
}
function qa_PingActiveRow() {
  try {
    var sh = _getTargetSheet_();
    var a = sh.getActiveRange();
    if (!a) { console.log('[PING] Нет активной строки'); return; }
    var row = a.getRow();
    if (row <= 1) { console.log('[PING] выбери строку данных (не заголовок)'); return; }

    var sec = Number(sh.getRange(row, _colDuration_()).getValue() || 0);
    var url = _findFirstAudioUrlInRow_(sh, row);

    console.log('[PING] row=', row, 'sec=', sec, 'url=', url || '(нет)');
    if (!url || !/^https?:\/\//i.test(url)) { console.log('[PING] в строке нет валидного URL'); return; }
    debugPingAudio_(url);
  } catch (e) { console.warn('[PING] qa_PingActiveRow exception', String(e)); }
}
function qa_PingRow(row) {
  try {
    row = Number(row);
    if (!row || row <= 1) { console.log('[PING] укажи корректный номер строки (>1)'); return; }
    var sh = _getTargetSheet_();

    var sec = Number(sh.getRange(row, _colDuration_()).getValue() || 0);
    var url = _findFirstAudioUrlInRow_(sh, row);

    console.log('[PING] row=', row, 'sec=', sec, 'url=', url || '(нет)');
    if (!url || !/^https?:\/\//i.test(url)) { console.log('[PING] в строке нет валидного URL'); return; }
    debugPingAudio_(url);
  } catch (e) { console.warn('[PING] qa_PingRow exception', String(e)); }
}

/* -------------------------- Selftests ---------------------------- */
function qa_Selftest() {
  try {
    var db = SpreadsheetApp.getActive().getSheetByName('БД');
    var token = db ? db.getRange('B5').getValue() : '';
    console.log('[SELFTEST] version =', getQaPatchesVersion_());
    console.log('[SELFTEST] base    =', getProtalkBase_());
    console.log('[SELFTEST] askUrl  =', protalkAskUrl_(token));
  } catch (e) { console.error('[SELFTEST] error', String(e)); }
}
function qa_PingFirstAudio() { debugPingFirstAudioFromSheet_(); }
function qa_GuardTranscribeSmokeTest() {
  var dummy = 'https://example.com/audio.wav';
  var res = validateAudioUrl_(dummy);
  console.log('[SMOKE] validate result:', JSON.stringify(res));
  var out = guardTranscribe_(dummy, 'nonExistingTranscribeFn_', { retryCount: 0 });
  console.log('[SMOKE] guardTranscribe_ ->', out);
}