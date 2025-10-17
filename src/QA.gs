// ============================== QA.gs (МОДЕРНИЗИРОВАННЫЙ) ==============================
// QA патчи и валидация на основе оригинального кода
// Ключевые улучшения:
// 1. Улучшенная валидация аудио URL
// 2. Более надежная работа с ProTalk API
// 3. Расширенная диагностика

// ---- Статусы и логирование --------------------------------------------
function QA_STATUS_NO_TRANSCRIPTION() {
  return 'СТАТУС: Нужна оценка (нет транскрипции)';
}

function qaLog(level, message, extra) {
  const tag = '[QA-SVC]';
  try {
    if (level === 'warn') console.warn(tag, message, extra || '');
    else if (level === 'error') console.error(tag, message, extra || '');
    else console.log(tag, message, extra || '');
  } catch (e) {
    console.error('[QA] Log error:', String(e));
  }
}

function getQaPatchesVersion() {
  return '2025-10-16.r4-modernized';
}

// ---- Конфигурация и утилиты -------------------------------------------
function getQaConfig() {
  try {
    const dbSheet = SpreadsheetApp.getActive().getSheetByName("БД");
    if (!dbSheet) throw new Error('Лист "БД" не найден');
    
    const targetSheetName = dbSheet.getRange("B4").getValue();
    const targetSheet = SpreadsheetApp.getActive().getSheetByName(targetSheetName);
    if (!targetSheet) throw new Error(`Лист "${targetSheetName}" не найден`);
    
    return {
      TARGET_SHEET: targetSheet,
      TARGET_SHEET_NAME: targetSheetName,
      BOT_TOKEN: dbSheet.getRange("B5").getValue(),
      BOT_ID: dbSheet.getRange("B6").getValue()
    };
  } catch (error) {
    console.error('[QA] Config error:', error.message);
    throw error;
  }
}

// ---- ProTalk API (улучшенный) -----------------------------------------
function getProtalkBase() {
  return 'https://eu1.api.pro-talk.ru';
}

function protalkAskUrl(botToken) {
  return getProtalkBase() + '/api/v1.0/ask/' + encodeURIComponent(String(botToken || ''));
}

function protalkFetchAsk(botToken, payload, fetchOpts) {
  const url = protalkAskUrl(botToken);
  const options = Object.assign({
    method: 'post',
    contentType: 'application/json; charset=utf-8',
    muteHttpExceptions: true,
    followRedirects: true,
    payload: JSON.stringify(payload || {})
  }, fetchOpts || {});
  
  qaLog('info', 'protalkFetchAsk()', { url: url, hasPayload: !!payload });

  try {
    const resp = UrlFetchApp.fetch(url, options);
    if (resp.getResponseCode() !== 200) {
      // Fallback на US1 сервер
      try {
        const alt = 'https://us1.api.pro-talk.ru/api/v1.0/ask/' + encodeURIComponent(String(botToken || ''));
        qaLog('warn', 'protalkFetchAsk: EU1 non-200, fallback to US1', resp.getResponseCode());
        const altResp = UrlFetchApp.fetch(alt, options);
        return altResp;
      } catch (altError) {
        qaLog('error', 'US1 fallback failed:', String(altError));
      }
    }
    return resp;
  } catch (error) {
    qaLog('error', 'protalkFetchAsk error:', String(error));
    throw error;
  }
}

function protalkFetchAskJson(botToken, payload, fetchOpts) {
  const resp = protalkFetchAsk(botToken, payload, fetchOpts);
  const code = resp.getResponseCode();
  const body = resp.getContentText() || '';
  
  if (code !== 200) {
    qaLog('error', 'protalkFetchAskJson: HTTP ' + code, body.slice(0, 600));
    throw new Error('Protalk ASK HTTP ' + code);
  }
  
  try {
    return JSON.parse(body);
  } catch (e) {
    qaLog('error', 'protalkFetchAskJson: JSON parse error', body.slice(0, 600));
    throw e;
  }
}

// ---- Валидация аудио URL (улучшенная) ---------------------------------
function validateAudioUrl(url) {
  const res = { ok: false, reason: '', url: '' };
  
  try {
    if (!url || typeof url !== 'string') {
      res.reason = 'empty';
      return res;
    }
    
    const u = url.trim();
    if (!u || u.toLowerCase() === 'нет записи') {
      res.reason = 'no_record';
      return res;
    }
    
    if (!/^https?:\/\//i.test(u)) {
      res.reason = 'not_http_https';
      return res;
    }
    
    if (/URL_аудио_файла_здесь/i.test(u) || /<.+>/.test(u)) {
      res.reason = 'placeholder';
      return res;
    }

    // Проверяем известные хосты
    const knownHosts = [
      'amocrm.mango-office.ru',
      'mango-office.ru', 
      'mango-office.com',
      'amocrm.ru',
      'pro-talk.ru'
    ];
    
    const hostMatch = u.match(/^https?:\/\/([^\/]+)/i);
    const host = (hostMatch ? hostMatch[1] : '').toLowerCase();
    
    if (host && !knownHosts.some(h => host.indexOf(h) !== -1)) {
      qaLog('warn', 'validateAudioUrl: unknown host (не блокируем)', host);
    }

    res.ok = true;
    res.url = u;
    return res;
  } catch (e) {
    res.reason = 'exception:' + String(e);
    return res;
  }
}

// ---- Guard для транскрипции (улучшенный) ------------------------------
function guardTranscribe(audioUrl, transcribeFnName, opts) {
  const options = Object.assign({
    retryCount: 1,
    retryDelayMs: 2500,
    minChars: 12
  }, opts || {});
  
  const v = validateAudioUrl(audioUrl);
  if (!v.ok) {
    qaLog('warn', 'guardTranscribe: invalid audio URL', v);
    return QA_STATUS_NO_TRANSCRIPTION();
  }

  const fn = (typeof transcribeFnName === 'function') ? transcribeFnName : (this[transcribeFnName] || null);
  if (typeof fn !== 'function') {
    qaLog('error', 'guardTranscribe: function not found', transcribeFnName);
    return QA_STATUS_NO_TRANSCRIPTION();
  }

  let attempt = 0;
  while (attempt <= options.retryCount) {
    try {
      attempt++;
      const text = String(fn(v.url) || '').trim();
      if (text && text.length >= options.minChars && !/^статус:/i.test(text)) {
        return text;
      }
      qaLog('warn', 'guardTranscribe: empty/short/status-like text', {
        attempt: attempt,
        len: (text || '').length
      });
    } catch (e) {
      qaLog('warn', 'guardTranscribe: transcribe threw', {
        attempt: attempt,
        err: String(e)
      });
    }
    if (attempt <= options.retryCount) {
      Utilities.sleep(options.retryDelayMs);
    }
  }
  
  return QA_STATUS_NO_TRANSCRIPTION();
}

// ---- Пинг аудио (улучшенный) ------------------------------------------
function debugPingAudio(url) {
  try {
    if (!url) {
      console.log('[PING] no url');
      return;
    }
    
    console.log(`[PING] Testing URL: ${url}`);
    
    const res = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { 'Range': 'bytes=0-0' },
      followRedirects: true,
      muteHttpExceptions: true
    });
    
    const headers = res.getHeaders() || {};
    const ct = headers['Content-Type'] || headers['content-type'] || '';
    const cl = headers['Content-Length'] || headers['content-length'] || '';
    const code = res.getResponseCode();
    
    console.log(`[PING] Result: code=${code}, ctype=${ct}, len=${cl}`);
    
    if (code === 200) {
      console.log('✅ URL доступен');
    } else if (code === 206) {
      console.log('✅ URL поддерживает Range запросы');
    } else {
      console.log(`⚠️ URL недоступен (код: ${code})`);
    }
  } catch (e) {
    console.warn('[PING] Exception:', String(e));
  }
}

function findFirstAudioUrlInRow(sheet, row) {
  try {
    // Проверяем колонку B (ссылка)
    const direct = String(sheet.getRange(row, 2).getValue() || '').trim();
    if (/^https?:\/\//i.test(direct)) return direct;

    // Ищем в других колонках
    const lastCol = Math.max(sheet.getLastColumn(), 20);
    const vals = sheet.getRange(row, 1, 1, lastCol).getValues()[0].map(x => String(x || ''));
    const re = /(https?:\/\/[^\s)"'>]+)/ig;
    
    for (let i = 0; i < vals.length; i++) {
      const cell = vals[i];
      let m;
      while ((m = re.exec(cell)) !== null) {
        if (m[1]) return m[1];
      }
    }
    
    return '';
  } catch (e) {
    qaLog('error', 'findFirstAudioUrlInRow:', String(e));
    return '';
  }
}

function debugPingFirstAudioFromSheet() {
  try {
    const config = getQaConfig();
    const sh = config.TARGET_SHEET;
    const last = sh.getLastRow();
    
    if (last < 2) {
      console.log('[PING] sheet empty');
      return;
    }

    const data = sh.getRange(2, 1, last - 1, Math.max(sh.getLastColumn(), 20)).getValues();
    let found = 0;
    
    for (let i = 0; i < data.length && found < 3; i++) {
      const rowIdx = i + 2;
      const url = findFirstAudioUrlInRow(sh, rowIdx);
      const sec = Number(sh.getRange(rowIdx, 12).getValue() || 0); // Колонка L
      
      if (url && /^https?:\/\//i.test(url) && sec > 0) {
        console.log(`[PING] Row ${rowIdx}: ${url} (${sec}s)`);
        debugPingAudio(url);
        found++;
      }
    }
    
    if (found === 0) {
      console.log('[PING] No audio URLs found');
    }
  } catch (e) {
    console.warn('[PING] Sheet exception:', String(e));
  }
}

// ---- Публичные функции для тестирования ------------------------------
function qa_PingActiveRow() {
  try {
    const config = getQaConfig();
    const sh = config.TARGET_SHEET;
    const a = sh.getActiveRange();
    
    if (!a) {
      console.log('[PING] Нет активной строки');
      return;
    }
    
    const row = a.getRow();
    if (row <= 1) {
      console.log('[PING] выбери строку данных (не заголовок)');
      return;
    }

    const sec = Number(sh.getRange(row, 12).getValue() || 0); // Колонка L
    const url = findFirstAudioUrlInRow(sh, row);

    console.log(`[PING] row=${row}, sec=${sec}, url=${url || '(нет)'}`);
    
    if (!url || !/^https?:\/\//i.test(url)) {
      console.log('[PING] в строке нет валидного URL');
      return;
    }
    
    debugPingAudio(url);
  } catch (e) {
    console.warn('[PING] qa_PingActiveRow exception:', String(e));
  }
}

function qa_PingRow(row) {
  try {
    row = Number(row);
    if (!row || row <= 1) {
      console.log('[PING] укажи корректный номер строки (>1)');
      return;
    }
    
    const config = getQaConfig();
    const sh = config.TARGET_SHEET;

    const sec = Number(sh.getRange(row, 12).getValue() || 0); // Колонка L
    const url = findFirstAudioUrlInRow(sh, row);

    console.log(`[PING] row=${row}, sec=${sec}, url=${url || '(нет)'}`);
    
    if (!url || !/^https?:\/\//i.test(url)) {
      console.log('[PING] в строке нет валидного URL');
      return;
    }
    
    debugPingAudio(url);
  } catch (e) {
    console.warn('[PING] qa_PingRow exception:', String(e));
  }
}

// ---- Selftests --------------------------------------------------------
function qa_Selftest() {
  try {
    const config = getQaConfig();
    console.log('[SELFTEST] version =', getQaPatchesVersion());
    console.log('[SELFTEST] base =', getProtalkBase());
    console.log('[SELFTEST] askUrl =', protalkAskUrl(config.BOT_TOKEN));
    console.log('[SELFTEST] target sheet =', config.TARGET_SHEET_NAME);
  } catch (e) {
    console.error('[SELFTEST] error:', String(e));
  }
}

function qa_PingFirstAudio() {
  debugPingFirstAudioFromSheet();
}

function qa_GuardTranscribeSmokeTest() {
  try {
    const dummy = 'https://example.com/audio.wav';
    const res = validateAudioUrl(dummy);
    console.log('[SMOKE] validate result:', JSON.stringify(res));
    
    const out = guardTranscribe(dummy, 'nonExistingTranscribeFn_', { retryCount: 0 });
    console.log('[SMOKE] guardTranscribe_ ->', out);
  } catch (e) {
    console.error('[SMOKE] error:', String(e));
  }
}

// ---- Диагностика системы ----------------------------------------------
function qa_FullDiagnostic() {
  try {
    console.log('[QA] ===========================================');
    console.log('[QA] ПОЛНАЯ ДИАГНОСТИКА QA СИСТЕМЫ');
    console.log('[QA] ===========================================');
    
    // 1. Selftest
    qa_Selftest();
    
    // 2. Тест валидации
    qa_GuardTranscribeSmokeTest();
    
    // 3. Тест пинга аудио
    qa_PingFirstAudio();
    
    // 4. Проверка конфигурации
    const config = getQaConfig();
    console.log('[QA] Config check:', {
      hasSheet: !!config.TARGET_SHEET,
      sheetName: config.TARGET_SHEET_NAME,
      hasBotToken: !!config.BOT_TOKEN,
      hasBotId: !!config.BOT_ID
    });
    
    console.log('[QA] ===========================================');
    console.log('[QA] ДИАГНОСТИКА ЗАВЕРШЕНА');
    console.log('[QA] ===========================================');
  } catch (e) {
    console.error('[QA] FullDiagnostic error:', String(e));
  }
}