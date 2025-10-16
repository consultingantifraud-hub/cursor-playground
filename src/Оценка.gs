// ============================== Оценка.gs (МОДЕРНИЗИРОВАННЫЙ) ==============================
// Основан на рабочем коде, но с улучшенной обработкой ошибок и логированием
// Ключевые улучшения:
// 1. Улучшенная обработка ошибок и логирование
// 2. Синхронизация с AMO.gs
// 3. Более надежная работа с ProTalk API
// 4. Логирование в отдельный лист LOGS

// ---- Константы колонок (1-based) --------------------------------------
const COL_B_LINK = 2;           // B — ссылка на запись
const COL_J_RESPONSIBLE = 10;   // J — ответственный
const COL_L_DURATION = 12;      // L — длительность (сек)
const COL_P_ALLFIELDS = 16;     // P — все поля (ищем STOP BOT)
const COL_V_TRANSCRIPT = 22;    // V — транскрипция
const COL_W_EVAL = 23;          // W — оценка
const COL_Y_STARTED_AT = 25;    // Y — отметка старта оценки

// ---- Настройки --------------------------------------------------------
const SAFE_EVAL = 'Статус: мало данных/короткий звонок. Следующий шаг: перезвонить клиенту сегодня и уточнить запрос и следующий контакт.';
const MAX_ROWS_SCAN = 400;
const TIMEOUT_MS = 300000; // 5 минут

// ---- Получение конфигурации -------------------------------------------
function getEvaluationConfig() {
  try {
    const dbSheet = SpreadsheetApp.getActive().getSheetByName("БД");
    if (!dbSheet) throw new Error("Лист 'БД' не найден");
    
    const targetSheetName = dbSheet.getRange("B4").getValue();
    const targetSheet = SpreadsheetApp.getActive().getSheetByName(targetSheetName);
    if (!targetSheet) throw new Error(`Лист '${targetSheetName}' не найден`);
    
    return {
      BOT_TOKEN: dbSheet.getRange("B5").getValue(),
      BOT_ID: dbSheet.getRange("B6").getValue(),
      TARGET_SHEET: targetSheet,
      TARGET_SHEET_NAME: targetSheetName,
      MIN_SEC_TRANS: Number(dbSheet.getRange("H2").getValue()) || 60,
      MIN_SEC_EVAL: Number(dbSheet.getRange("I2").getValue()) || 60,
      MAX_TEXT_CHARS: Number(dbSheet.getRange("K2").getValue()) || 29000,
      PROMPT_TRANS: dbSheet.getRange("H5").getValue() || '',
      PROMPT_EVAL: dbSheet.getRange("I5").getValue() || ''
    };
  } catch (error) {
    console.error('❌ Ошибка конфигурации оценки:', error.message);
    throw error;
  }
}

// ---- Утилиты ----------------------------------------------------------
function hasStopBot(text) {
  return /\bstop\s*bot\s*:\s*true\b/i.test(String(text || ''));
}

function clipText(text, maxLength) {
  const t = String(text || '');
  return (maxLength && t.length > maxLength) ? t.slice(0, maxLength) : t;
}

function isEmptyV(v) {
  const s = String(v || '').trim();
  return !s || /^нужен\s*текст$/i.test(s) || /^без\s*текста$/i.test(s) || /^статус:\s*нужна\s*оценка/i.test(s);
}

function isEmptyW(w) {
  const s = String(w || '').trim();
  return !s || /^нужна\s*оценка$/i.test(s) || /^без\s*оценки$/i.test(s);
}

function findFirstAudioUrlInRow(sheet, row) {
  const direct = String(sheet.getRange(row, COL_B_LINK).getValue() || '').trim();
  if (/^https?:\/\//i.test(direct)) return direct;
  
  const lastCol = Math.max(sheet.getLastColumn(), COL_B_LINK);
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
}

// ---- Логирование в LOGS -----------------------------------------------
function ensureLogsSheet() {
  try {
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName('LOGS');
    if (!sh) {
      sh = ss.insertSheet('LOGS');
      sh.appendRow(['time', 'chat_id', 'phase', 'text_preview', 'provider', 'bot_id', 'model', 'req_tokens', 'resp_tokens', 'latency_ms', 'error', 'endpoint']);
      console.log('✅ Создан лист LOGS');
    }
    return sh;
  } catch (error) {
    console.error('❌ Ошибка создания листа LOGS:', error.message);
    return null;
  }
}

function logProtalk(row) {
  try {
    const sh = ensureLogsSheet();
    if (!sh) {
      console.warn('⚠️ Лист LOGS недоступен');
      return;
    }
    
    sh.appendRow([
      new Date(),
      String(row.chat_id || ''),
      String(row.phase || ''),
      String((row.text || '').slice(0, 140)),
      String(row.provider || 'Replica'),
      String(row.bot_id || ''),
      String(row.model || ''),
      String(row.req_tokens || ''),
      String(row.resp_tokens || ''),
      String(row.latency_ms || ''),
      String(row.error || ''),
      String(row.endpoint || '')
    ]);
    console.log(`📝 Логировано: ${row.phase} для чата ${row.chat_id}`);
  } catch (error) {
    console.warn('⚠️ Ошибка логирования:', error.message);
  }
}

// ---- ProTalk API (улучшенный) -----------------------------------------
function getProtalkBase() {
  return 'https://eu1.api.pro-talk.ru';
}

function protalkAskUrl(botToken) {
  return getProtalkBase() + '/api/v1.0/ask/' + encodeURIComponent(String(botToken || ''));
}

function askWithRetry(botToken, payload) {
  const start = Date.now();
  const chatId = (payload && payload.chat_id) || ('chat_' + Date.now());
  
  logProtalk({
    chat_id: chatId,
    phase: 'STARTED',
    text: (payload && payload.message) || '',
    endpoint: protalkAskUrl(botToken)
  });

  let tries = 0;
  const maxTries = 2;
  
  while (tries <= maxTries) {
    tries++;
    try {
      const response = UrlFetchApp.fetch(protalkAskUrl(botToken), {
        method: 'post',
        contentType: 'application/json; charset=utf-8',
        muteHttpExceptions: true,
        followRedirects: true,
        timeout: TIMEOUT_MS,
        payload: JSON.stringify(payload || {})
      });
      
      const code = response.getResponseCode();
      const body = response.getContentText() || '';
      
      if (code === 200) {
        try {
          const json = JSON.parse(body);
          const text = pickDoneText(json);
          
          logProtalk({
            chat_id: chatId,
            phase: 'DONE',
            text: text || '',
            provider: 'Replica',
            model: (json && (json.model || json.engine || '')),
            latency_ms: (Date.now() - start),
            endpoint: protalkAskUrl(botToken)
          });
          
          return (text && text.trim()) ? json : null;
        } catch (parseError) {
          logProtalk({
            chat_id: chatId,
            phase: 'ERROR_JSON',
            error: String(parseError),
            endpoint: protalkAskUrl(botToken)
          });
          return null;
        }
      }
      
      logProtalk({
        chat_id: chatId,
        phase: 'HTTP_' + code,
        error: body.slice(0, 180),
        endpoint: protalkAskUrl(botToken)
      });
      
      // Fallback на US1 сервер
      if (tries === 1) {
        try {
          const usUrl = 'https://us1.api.pro-talk.ru/api/v1.0/ask/' + encodeURIComponent(String(botToken || ''));
          console.log('🔄 Fallback на US1 сервер');
          const usResponse = UrlFetchApp.fetch(usUrl, {
            method: 'post',
            contentType: 'application/json; charset=utf-8',
            muteHttpExceptions: true,
            followRedirects: true,
            timeout: TIMEOUT_MS,
            payload: JSON.stringify(payload || {})
          });
          
          if (usResponse.getResponseCode() === 200) {
            const usJson = JSON.parse(usResponse.getContentText());
            const usText = pickDoneText(usJson);
            
            logProtalk({
              chat_id: chatId,
              phase: 'DONE_US1',
              text: usText || '',
              provider: 'Replica',
              model: (usJson && (usJson.model || usJson.engine || '')),
              latency_ms: (Date.now() - start),
              endpoint: usUrl
            });
            
            return (usText && usText.trim()) ? usJson : null;
          }
        } catch (usError) {
          console.warn('⚠️ US1 fallback failed:', usError.message);
        }
      }
      
      return null;
    } catch (error) {
      logProtalk({
        chat_id: chatId,
        phase: 'THROW',
        error: String(error),
        endpoint: protalkAskUrl(botToken)
      });
      
      if (Date.now() - start > 30000) return null;
      Utilities.sleep(800);
    }
  }
  
  return null;
}

function pickDoneText(json) {
  if (!json) return '';
  if (json.done && typeof json.done === 'string' && json.done.trim()) return json.done.trim();
  if (json.text && typeof json.text === 'string' && json.text.trim()) return json.text.trim();
  if (json.result && typeof json.result === 'string' && json.result.trim()) return json.result.trim();
  if (json.result && json.result.text && typeof json.result.text === 'string') return json.result.text.trim();
  if (json.message && typeof json.message === 'string') return json.message.trim();
  return '';
}

// ---- Нормализация бренда (только для оценки) ---------------------------
const COMPANY_RE = /(сту[дт][ие]тал[ияеьй]|студитали|сто\s*детал[еёеиы][йи]?|100\s*детал[еёеиы][йи]?)/gi;

function normalizeBrandEval(text) {
  if (!text) return text;
  return String(text)
    .replace(COMPANY_RE, 'Сто Деталей')
    .replace(/\bКомпания\s+Сто\s+Детал[еёеи][йи]?\b/gi, 'Компания «Сто Деталей»');
}

// ---- Синглтон для предотвращения параллельных запусков ---------------
const RUN_KEY = 'RUN_TO_PROTALK_BOT';
const RUN_WALL_MS = 20000;

function beginRun() {
  const props = PropertiesService.getScriptProperties();
  const now = Date.now();
  const raw = props.getProperty(RUN_KEY);
  
  if (raw) {
    const ts = Number(raw) || 0;
    if (ts && now - ts < RUN_WALL_MS) {
      console.log('⏭️ Пропускаем: уже выполняется');
      return { ok: false, release: function() {} };
    }
  }
  
  props.setProperty(RUN_KEY, String(now));
  return {
    ok: true,
    release: function() {
      try {
        PropertiesService.getScriptProperties().deleteProperty(RUN_KEY);
      } catch (e) {}
    }
  };
}

// ---- Основной воркер (МОДЕРНИЗИРОВАННЫЙ) -----------------------------
function ToProTalkBot() {
  const run = beginRun();
  if (!run.ok) return;
  
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) {
    console.log('⏭️ Занято другим процессом');
    run.release();
    return;
  }

  const summary = { scanned: 0, skipped: 0, transcribed: 0, evaluated: 0 };

  try {
    const cfg = getEvaluationConfig();
    const sh = cfg.TARGET_SHEET;
    const lastRow = sh.getLastRow();
    
    if (lastRow < 2) {
      console.log('📊 Лист пуст');
      return;
    }

    const width = Math.max(COL_P_ALLFIELDS, sh.getLastColumn());
    const rowsToScan = Math.min(MAX_ROWS_SCAN, lastRow - 1);
    const values = sh.getRange(2, 1, rowsToScan, width).getValues();

    console.log(`🔍 Сканируем ${rowsToScan} строк`);

    // ---------- Шаг A: транскрипция ----------
    for (let i = 0; i < values.length; i++) {
      summary.scanned++;
      const row = i + 2;

      const V = String(values[i][COL_V_TRANSCRIPT - 1] || '').trim();
      if (!isEmptyV(V)) {
        summary.skipped++;
        continue;
      }

      const all = values[i][COL_P_ALLFIELDS - 1];
      if (hasStopBot(all)) {
        summary.skipped++;
        continue;
      }

      const link = findFirstAudioUrlInRow(sh, row);
      if (!/^https?:\/\//i.test(link)) {
        summary.skipped++;
        continue;
      }

      const sec = Number(values[i][COL_L_DURATION - 1] || 0);
      if (!(sec >= cfg.MIN_SEC_TRANS)) {
        summary.skipped++;
        continue;
      }

      console.log(`🎯 Обрабатываем строку ${row} для транскрипции`);

      const msg = cfg.PROMPT_TRANS ? (cfg.PROMPT_TRANS + '\n\n' + link) : link;
      const json = askWithRetry(cfg.BOT_TOKEN, {
        bot_id: cfg.BOT_ID,
        chat_id: 'trans_' + Date.now(),
        message: msg
      });
      
      const out = clipText(pickDoneText(json) || 'СТАТУС: Нужна оценка (нет транскрипции)', cfg.MAX_TEXT_CHARS);

      sh.getRange(row, COL_V_TRANSCRIPT).setValue(out);
      SpreadsheetApp.flush();

      summary.transcribed = 1;
      console.log(`✅ V записано в строку ${row}, длина: ${out.length}`);
      return; // одна операция за запуск
    }

    // ---------- Шаг Б: оценка ----------
    for (let j = 0; j < values.length; j++) {
      const row2 = j + 2;

      const v = String(values[j][COL_V_TRANSCRIPT - 1] || '').trim();
      const w = String(values[j][COL_W_EVAL - 1] || '').trim();
      if (isEmptyV(v) || !isEmptyW(w)) {
        continue;
      }

      const all2 = values[j][COL_P_ALLFIELDS - 1];
      if (hasStopBot(all2)) {
        continue;
      }

      const sec2 = Number(values[j][COL_L_DURATION - 1] || 0);
      if (!(sec2 >= cfg.MIN_SEC_EVAL)) {
        continue;
      }

      console.log(`🎯 Обрабатываем строку ${row2} для оценки`);

      if (/^статус:\s*нужна\s*оценка/i.test(v)) {
        sh.getRange(row2, COL_W_EVAL).setValue(SAFE_EVAL);
        SpreadsheetApp.flush();
        summary.evaluated = 1;
        console.log(`✅ W SAFE по статусу в строке ${row2}`);
        return;
      }

      const cleanLen = v.replace(/\[неразборчиво\]/gi, '').replace(/\s+/g, ' ').trim().length;
      if (cleanLen < 120) {
        sh.getRange(row2, COL_W_EVAL).setValue(SAFE_EVAL);
        SpreadsheetApp.flush();
        summary.evaluated = 1;
        console.log(`✅ W SAFE по короткому тексту в строке ${row2}, длина: ${cleanLen}`);
        return;
      }

      try {
        sh.getRange(row2, COL_Y_STARTED_AT).setValue(new Date());
      } catch (e) {}

      const prompt = (cfg.PROMPT_EVAL ? (cfg.PROMPT_EVAL.trim() + '\n\n') : '') + 'Транскрипция (без правок):\n' + v;
      const json2 = askWithRetry(cfg.BOT_TOKEN, {
        bot_id: cfg.BOT_ID,
        chat_id: 'score_' + Date.now(),
        message: prompt
      });
      
      let out2 = pickDoneText(json2);
      if (!out2 || /^статус\s*:/i.test(out2)) out2 = SAFE_EVAL;
      out2 = clipText(String(out2).replace(/\*/g, ''), cfg.MAX_TEXT_CHARS);
      out2 = normalizeBrandEval(out2);

      sh.getRange(row2, COL_W_EVAL).setValue(out2);
      SpreadsheetApp.flush();

      summary.evaluated = 1;
      console.log(`✅ W записано в строку ${row2}, длина: ${out2.length}`);
      return;
    }

    console.log(`📊 Ничего не найдено; сканировано: ${summary.scanned}, пропущено: ${summary.skipped}`);

  } catch (error) {
    console.error('❌ Ошибка в ToProTalkBot:', error.message);
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
    run.release();
  }
}

// ---- Совместимость со старым именем ------------------------------------
function chatToProTalkBot() {
  return ToProTalkBot();
}

// ---- Ручные раннеры для отладки ----------------------------------------
function qa_RunActiveRowForce() {
  try {
    const cfg = getEvaluationConfig();
    const sh = cfg.TARGET_SHEET;
    const a = sh.getActiveRange();
    if (!a) {
      console.log('❌ Нет активной строки');
      return;
    }
    qa_RunRowForce(a.getRow());
  } catch (error) {
    console.error('❌ Ошибка qa_RunActiveRowForce:', error.message);
  }
}

function qa_RunRowForce(row) {
  qa_RunRow(row, true);
}

function qa_RunRow(row, force) {
  try {
    row = Number(row);
    if (!row || row <= 1) {
      console.log('❌ Укажите корректный номер строки (>1)');
      return;
    }

    const cfg = getEvaluationConfig();
    const sh = cfg.TARGET_SHEET;

    const link = findFirstAudioUrlInRow(sh, row);
    const sec = Number(sh.getRange(row, COL_L_DURATION).getValue() || 0);
    const all = String(sh.getRange(row, COL_P_ALLFIELDS).getValue() || '');

    if (!force && hasStopBot(all)) {
      console.log('⏭️ STOP BOT');
      return;
    }
    if (!/^https?:\/\//i.test(link)) {
      console.log('❌ Нет валидного URL');
      return;
    }

    const V = String(sh.getRange(row, COL_V_TRANSCRIPT).getValue() || '').trim();
    const W = String(sh.getRange(row, COL_W_EVAL).getValue() || '').trim();

    if (isEmptyV(V)) {
      if (!force && !(sec >= cfg.MIN_SEC_TRANS)) {
        console.log('❌ Секунд < MIN_SEC_TRANS');
        return;
      }
      
      console.log(`🎯 Транскрипция для строки ${row}`);
      const msg = cfg.PROMPT_TRANS ? (cfg.PROMPT_TRANS + '\n\n' + link) : link;
      const json = askWithRetry(cfg.BOT_TOKEN, {
        bot_id: cfg.BOT_ID,
        chat_id: 'trans_row_' + row + '_' + Date.now(),
        message: msg
      });
      const out = clipText(pickDoneText(json) || 'СТАТУС: Нужна оценка (нет транскрипции)', cfg.MAX_TEXT_CHARS);
      sh.getRange(row, COL_V_TRANSCRIPT).setValue(out);
      SpreadsheetApp.flush();
      console.log('✅ V записано');
      return;
    }

    if (isEmptyW(W)) {
      if (!force && !(sec >= cfg.MIN_SEC_EVAL)) {
        console.log('❌ Секунд < MIN_SEC_EVAL');
        return;
      }
      
      if (/^статус:\s*нужна\s*оценка/i.test(V)) {
        sh.getRange(row, COL_W_EVAL).setValue(SAFE_EVAL);
        SpreadsheetApp.flush();
        console.log('✅ W SAFE');
        return;
      }
      
      const cleanLen = V.replace(/\[неразборчиво\]/gi, '').replace(/\s+/g, ' ').trim().length;
      if (cleanLen < 120) {
        sh.getRange(row, COL_W_EVAL).setValue(SAFE_EVAL);
        SpreadsheetApp.flush();
        console.log('✅ W SAFE короткий');
        return;
      }
      
      console.log(`🎯 Оценка для строки ${row}`);
      try {
        sh.getRange(row, COL_Y_STARTED_AT).setValue(new Date());
      } catch (e) {}
      
      const prompt = (cfg.PROMPT_EVAL ? (cfg.PROMPT_EVAL.trim() + '\n\n') : '') + 'Транскрипция (без правок):\n' + V;
      const json2 = askWithRetry(cfg.BOT_TOKEN, {
        bot_id: cfg.BOT_ID,
        chat_id: 'score_row_' + row + '_' + Date.now(),
        message: prompt
      });
      
      let out2 = pickDoneText(json2);
      if (!out2 || /^статус\s*:/i.test(out2)) out2 = SAFE_EVAL;
      out2 = normalizeBrandEval(clipText(String(out2).replace(/\*/g, ''), cfg.MAX_TEXT_CHARS));
      sh.getRange(row, COL_W_EVAL).setValue(out2);
      SpreadsheetApp.flush();
      console.log('✅ W записано');
      return;
    }

    console.log('ℹ️ Ничего не делаем — V и W уже есть');
  } catch (error) {
    console.error('❌ Ошибка qa_RunRow:', error.message);
  }
}