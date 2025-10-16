// ============================== Заметки.gs (МОДЕРНИЗИРОВАННЫЙ) ==============================
// Основан на рабочем коде, но с улучшенной обработкой ошибок и логированием
// Ключевые улучшения:
// 1. Улучшенная обработка ошибок и валидация
// 2. Синхронизация с AMO.gs
// 3. Детальное логирование в LOG_NOTES
// 4. Более надежная отправка заметок

// ---- Константы --------------------------------------------------------
const NOTES_TIME_BUDGET_MS = 330000; // 5.5 минут
const MAX_ROWS_PER_RUN = 25;
const NOTES_PAGES_MAX = 2;
const DUP_CHECK_RECENT_LIMIT = 120;
const PAUSE_BETWEEN_OPS_MS = 120;
const AMO_RETRY_TRIES = 2;
const AMO_RETRY_BASE_MS = 400;
const ENABLE_LOG_SHEET = true;

// ---- Получение конфигурации (синхронизировано с AMO.gs) --------------
function getNotesConfig() {
  try {
    const dbSheet = SpreadsheetApp.getActive().getSheetByName("БД");
    if (!dbSheet) throw new Error('Лист "БД" не найден');
    
    const amoToken = dbSheet.getRange("B2").getValue();
    let amoSubdomain = dbSheet.getRange("B3").getValue();
    const dataSheetName = dbSheet.getRange("B4").getValue();

    // Проверка обязательных параметров
    if (!amoToken || !amoSubdomain) {
      throw new Error('Не указан токен или поддомен');
    }
    
    // Нормализация поддомена
    amoSubdomain = String(amoSubdomain).trim().replace(/^https?:\/\//i, '').replace(/\.amocrm\.(ru|eu)$/i, '');
    
    if (!/^[a-z0-9-]+$/i.test(amoSubdomain)) {
      throw new Error(`Некорректный поддомен: ${amoSubdomain}`);
    }

    return {
      AMO_TOKEN: String(amoToken).trim(),
      AMO_SUBDOMAIN: amoSubdomain,
      SHEET_NAME: String(dataSheetName).trim()
    };
  } catch (error) {
    console.error('❌ Ошибка конфигурации заметок:', error.message);
    throw error;
  }
}

// ---- Утилиты ----------------------------------------------------------
function normalizeText(text) {
  return String(text || '')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/[ \t]+/g, ' ')
    .replace(/\u00A0/g, ' ')
    .trim();
}

// ---- Получение существующих заметок сделки ---------------------------
function getLeadCommonNotesTextsLimited(dealId, pagesLimit, recentLimit) {
  try {
    const config = getNotesConfig();
    const keep = Math.max(1, Number(recentLimit || 100));
    const maxPages = Math.max(1, Number(pagesLimit || 1));
    let notes = [];
    let url = `https://${config.AMO_SUBDOMAIN}.amocrm.ru/api/v4/leads/${dealId}/notes?with=attachments&limit=250`;
    let pages = 0;

    while (url && pages < maxPages) {
      const resp = UrlFetchApp.fetch(url, {
        method: 'get',
        headers: { 'Authorization': 'Bearer ' + config.AMO_TOKEN },
        muteHttpExceptions: true,
        followRedirects: true
      });
      
      const code = resp.getResponseCode();
      if (code < 200 || code >= 300) break;
      
      const json = JSON.parse(resp.getContentText() || '{}');
      notes = notes.concat(json?._embedded?.notes || []);
      url = json?._links?.next?.href || '';
      pages++;
    }

    const texts = [];
    for (const n of notes) {
      if (n.note_type === 'common') {
        const p = n.params || {};
        const t = (typeof p.text === 'string' && p.text.trim()) ? p.text.trim() : (typeof n.text === 'string' && n.text.trim() ? n.text.trim() : '');
        if (t) texts.push(t);
      }
      if (texts.length >= keep) break;
    }
    
    return { set: new Set(texts.map(normalizeText)), raw: texts };
  } catch (error) {
    console.error(`❌ Ошибка получения заметок для сделки ${dealId}:`, error.message);
    return { set: new Set(), raw: [] };
  }
}

// ---- Проверка дубликатов ----------------------------------------------
function sendNoteIfNeed(dealId, text, existing) {
  const norm = normalizeText(text);
  if (existing.set.has(norm)) return 'DUPLICATE';
  
  const res = addNoteToAmo(dealId, text);
  if (res === 'OK') existing.set.add(norm);
  return res;
}

// ---- Отправка заметки в amoCRM (улучшенная) ---------------------------
function addNoteToAmo(dealId, noteText) {
  try {
    const config = getNotesConfig();
    const text = String(noteText || '');
    const body1 = [{ entity_id: Number(dealId), note_type: 'common', params: { text: text } }];
    const body2 = [{ note_type: 'common', params: { text: text } }];
    const headers = {
      'Authorization': 'Bearer ' + config.AMO_TOKEN,
      'Content-Type': 'application/json',
      'Accept': 'application/hal+json'
    };

    return withBackoff(function() {
      // Путь 1: коллекционный
      let resp = UrlFetchApp.fetch(`https://${config.AMO_SUBDOMAIN}.amocrm.ru/api/v4/leads/notes`, {
        method: 'post',
        headers,
        muteHttpExceptions: true,
        followRedirects: true,
        payload: JSON.stringify(body1)
      });
      
      let code = resp.getResponseCode();
      if (code === 400) {
        const txt = resp.getContentText() || '';
        if (/code"\s*:\s*226|Error 226/i.test(txt)) return 'DUPLICATE';
        
        // Fallback путь 2: по сделке
        resp = UrlFetchApp.fetch(`https://${config.AMO_SUBDOMAIN}.amocrm.ru/api/v4/leads/${dealId}/notes`, {
          method: 'post',
          headers,
          muteHttpExceptions: true,
          followRedirects: true,
          payload: JSON.stringify(body2)
        });
        
        code = resp.getResponseCode();
        if (code === 400) {
          const txt2 = resp.getContentText() || '';
          if (/code"\s*:\s*226|Error 226/i.test(txt2)) return 'DUPLICATE';
          
          if (/length|too.*long/i.test(txt2)) {
            body2[0].params.text = text.slice(0, 7999);
            resp = UrlFetchApp.fetch(`https://${config.AMO_SUBDOMAIN}.amocrm.ru/api/v4/leads/${dealId}/notes`, {
              method: 'post',
              headers,
              muteHttpExceptions: true,
              followRedirects: true,
              payload: JSON.stringify(body2)
            });
            code = resp.getResponseCode();
          }
          
          if (code < 200 || code >= 300) {
            throw new Error('amoCRM error ' + code + ': ' + (resp.getContentText() || '').slice(0, 600));
          }
          return 'OK';
        }
        
        if (code < 200 || code >= 300) {
          throw new Error('amoCRM error ' + code + ': ' + (resp.getContentText() || '').slice(0, 600));
        }
        return 'OK';
      }
      
      if (code < 200 || code >= 300) {
        throw new Error('amoCRM error ' + code + ': ' + (resp.getContentText() || '').slice(0, 600));
      }
      return 'OK';
    }, { tries: AMO_RETRY_TRIES, baseMs: AMO_RETRY_BASE_MS });
  } catch (error) {
    console.error(`❌ Ошибка отправки заметки для сделки ${dealId}:`, error.message);
    return 'ERROR';
  }
}

// ---- Retry с экспоненциальной задержкой -------------------------------
function withBackoff(fn, opt) {
  const tries = (opt && opt.tries) || 2;
  const base = (opt && opt.baseMs) || 400;
  let a = 0;
  
  for (; a <= tries; a++) {
    try {
      return fn();
    } catch (e) {
      if (a >= tries) throw e;
      Utilities.sleep(base * Math.pow(2, a) + Math.floor(Math.random() * 120));
    }
  }
}

// ---- Логирование в LOG_NOTES ------------------------------------------
function ensureLogSheet() {
  try {
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName('LOG_NOTES');
    if (!sh) {
      sh = ss.insertSheet('LOG_NOTES');
      sh.getRange(1, 1, 1, 5).setValues([['time', 'dealId', 'row', 'col', 'result']]);
      console.log('✅ Создан лист LOG_NOTES');
    }
    return sh;
  } catch (error) {
    console.error('❌ Ошибка создания листа LOG_NOTES:', error.message);
    return null;
  }
}

function logRow(dealId, row, col, res, extra) {
  if (!ENABLE_LOG_SHEET) return;
  
  try {
    const sh = ensureLogSheet();
    if (!sh) return;
    
    sh.appendRow([
      new Date(),
      String(dealId),
      String(row),
      String(col),
      String(res || '') + (extra ? (' | ' + extra) : '')
    ]);
  } catch (error) {
    console.warn('⚠️ Ошибка логирования строки:', error.message);
  }
}

function logSummary(stats, t0) {
  const ms = Date.now() - t0;
  const msg = `[NOTES] done in ${ms}ms: rows_seen=${stats.rows_seen}, rows_done=${stats.rows_done}, notes_sent=${stats.notes_sent}, dup=${stats.notes_dup}, errors=${stats.rows_error}`;
  console.log(msg);
  
  if (ENABLE_LOG_SHEET) {
    try {
      const sh = ensureLogSheet();
      if (sh) {
        sh.appendRow([new Date(), 'SUMMARY', '', '', msg]);
      }
    } catch (error) {
      console.warn('⚠️ Ошибка логирования итогов:', error.message);
    }
  }
}

// ---- Основная функция отправки заметок (МОДЕРНИЗИРОВАННАЯ) -----------
function sendNotesToAmoCRM() {
  const t0 = Date.now();
  const stats = { rows_seen: 0, rows_done: 0, notes_sent: 0, notes_dup: 0, rows_error: 0 };

  try {
    console.log('🔄 Начало отправки заметок в amoCRM');
    
    const config = getNotesConfig();
    const sh = SpreadsheetApp.getActive().getSheetByName(config.SHEET_NAME);
    if (!sh) throw new Error(`Лист "${config.SHEET_NAME}" не найден`);
    
    const last = sh.getLastRow();
    if (last < 2) {
      console.log('📊 Лист пуст');
      logSummary(stats, t0);
      return;
    }

    const cols = Math.max(27, Math.min(60, sh.getLastColumn()));
    const data = sh.getRange(2, 1, last - 1, cols).getValues();

    console.log(`📊 Всего строк: ${last - 1}, колонок: ${cols}`);

    const queue = [];
    for (let i = 0; i < data.length && queue.length < MAX_ROWS_PER_RUN; i++) {
      const R = i + 2;
      const r = data[i];
      const dealId = r[5];        // F (6-я колонка)
      const transcript = r[21];   // V (22-я колонка)
      const evalText = r[22];     // W (23-я колонка)
      const aa = String(r[26] || '').trim(); // AA (27-я колонка)

      if (!dealId) continue;
      if (!(transcript || evalText)) continue;
      if (aa === 'OK') continue;

      queue.push({
        R,
        dealId,
        transcript: String(transcript || '').trim(),
        assessment: String(evalText || '').trim()
      });
    }

    stats.rows_seen = queue.length;
    console.log(`📝 Кандидатов для отправки: ${queue.length}`);

    if (!queue.length) {
      console.log('ℹ️ Нет строк для обработки');
      logSummary(stats, t0);
      return;
    }

    // Группируем по сделкам
    const grouped = new Map();
    for (const it of queue) {
      if (!grouped.has(it.dealId)) grouped.set(it.dealId, []);
      grouped.get(it.dealId).push(it);
    }

    console.log(`📊 Групп сделок: ${grouped.size}`);

    // Обрабатываем каждую группу
    for (const [dealId, items] of grouped.entries()) {
      if (Date.now() - t0 > NOTES_TIME_BUDGET_MS) {
        console.log('⏰ Превышен бюджет времени');
        break;
      }

      let existing = { set: new Set(), raw: [] };
      try {
        existing = getLeadCommonNotesTextsLimited(dealId, NOTES_PAGES_MAX, DUP_CHECK_RECENT_LIMIT);
      } catch (e) {
        console.warn(`⚠️ Ошибка предзагрузки заметок для сделки ${dealId}:`, e.message);
      }

      for (const it of items) {
        if (Date.now() - t0 > NOTES_TIME_BUDGET_MS) break;

        const { R, transcript, assessment } = it;
        try {
          const resV = transcript ? sendNoteIfNeed(dealId, transcript, existing) : 'SKIP';
          const resW = assessment ? sendNoteIfNeed(dealId, assessment, existing) : 'SKIP';

          if (resV === 'OK') stats.notes_sent++;
          else if (resV === 'DUPLICATE') stats.notes_dup++;
          if (resW === 'OK') stats.notes_sent++;
          else if (resW === 'DUPLICATE') stats.notes_dup++;

          logRow(dealId, R, 'V', resV);
          logRow(dealId, R, 'W', resW);

          const ok = [resV, resW].every(s => s === 'OK' || s === 'DUPLICATE' || s === 'SKIP');
          if (ok) {
            sh.getRange(R, 26).setValue(new Date()); // Z
            sh.getRange(R, 27).setValue('OK');       // AA
            stats.rows_done++;
            console.log(`✅ Строка ${R} обработана успешно`);
          } else {
            sh.getRange(R, 27).setValue('AMO_NOTE_ERROR');
            stats.rows_error++;
            console.log(`❌ Ошибка в строке ${R}`);
          }
        } catch (e) {
          sh.getRange(R, 27).setValue('AMO_NOTE_ERROR');
          stats.rows_error++;
          logRow(dealId, R, '-', 'ERROR', String(e));
          console.error(`❌ Ошибка в строке ${R}:`, e.message);
        }
        Utilities.sleep(PAUSE_BETWEEN_OPS_MS);
      }
    }
  } catch (e) {
    console.error('❌ Глобальная ошибка sendNotesToAmoCRM:', e.message);
  } finally {
    logSummary(stats, t0);
  }
}

// ---- Диагностика причины «rows_seen=0» --------------------------------
function notes_DryRunWhyZero() {
  try {
    console.log('🔍 Диагностика: почему rows_seen=0');
    
    const config = getNotesConfig();
    const sh = SpreadsheetApp.getActive().getSheetByName(config.SHEET_NAME);
    if (!sh) {
      console.log('❌ Нет листа', config.SHEET_NAME);
      return;
    }
    
    const last = sh.getLastRow();
    if (last < 2) {
      console.log('📊 Лист пуст');
      return;
    }

    const cols = Math.max(27, Math.min(60, sh.getLastColumn()));
    const data = sh.getRange(2, 1, last - 1, cols).getValues();
    let seen = 0, ok = 0;
    
    console.log(`📊 Анализируем ${data.length} строк, колонок: ${cols}`);
    
    for (let i = 0; i < Math.min(data.length, 20); i++) { // Показываем первые 20 строк
      const R = i + 2;
      const r = data[i];
      const dealId = r[5];        // F
      const V = r[21];            // V
      const W = r[22];            // W
      const AA = String(r[26] || '').trim(); // AA
      
      if (dealId) seen++;
      
      const why = [];
      if (!dealId) why.push('нет dealId(6)');
      if (!(V || W)) why.push('нет V/W(22/23)');
      if (AA === 'OK') why.push('AA=OK');
      
      if (!why.length) {
        ok++;
        console.log(`✅ OK-кандидат R=${R}, dealId=${dealId}, V=${!!V}, W=${!!W}`);
      } else {
        console.log(`⏭️ SKIP R=${R} → ${why.join('; ')} (dealId=${dealId}, V=${!!V}, W=${!!W}, AA="${AA}")`);
      }
    }
    
    console.log(`📊 Итог: dealId-строк=${seen}; кандидатов к отправке=${ok}`);
    
    // Дополнительная статистика
    let withV = 0, withW = 0, withBoth = 0, withAA = 0;
    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      const V = r[21];
      const W = r[22];
      const AA = String(r[26] || '').trim();
      
      if (V) withV++;
      if (W) withW++;
      if (V && W) withBoth++;
      if (AA === 'OK') withAA++;
    }
    
    console.log(`📈 Статистика по колонкам: V=${withV}, W=${withW}, V+W=${withBoth}, AA=OK=${withAA}`);
    
  } catch (error) {
    console.error('❌ Ошибка диагностики:', error.message);
  }
}

// ---- Тестовая функция для проверки конфигурации ----------------------
function testNotesConfig() {
  try {
    const config = getNotesConfig();
    console.log('✅ Конфигурация заметок:', {
      AMO_SUBDOMAIN: config.AMO_SUBDOMAIN,
      SHEET_NAME: config.SHEET_NAME,
      TOKEN_LENGTH: config.AMO_TOKEN.length
    });
  } catch (error) {
    console.error('❌ Ошибка конфигурации:', error.message);
  }
}

// ---- Проверка конкретных строк ----------------------------------------
function notes_CheckRows(startRow = 2, count = 10) {
  try {
    console.log(`🔍 Проверка строк ${startRow}-${startRow + count - 1}`);
    
    const config = getNotesConfig();
    const sh = SpreadsheetApp.getActive().getSheetByName(config.SHEET_NAME);
    if (!sh) {
      console.log('❌ Нет листа', config.SHEET_NAME);
      return;
    }
    
    for (let i = 0; i < count; i++) {
      const row = startRow + i;
      if (row > sh.getLastRow()) break;
      
      const dealId = sh.getRange(row, 6).getValue();      // F
      const V = sh.getRange(row, 22).getValue();          // V
      const W = sh.getRange(row, 23).getValue();          // W
      const AA = String(sh.getRange(row, 27).getValue() || '').trim(); // AA
      
      console.log(`Строка ${row}:`, {
        dealId: dealId,
        V: V ? `"${String(V).slice(0, 50)}..."` : 'пусто',
        W: W ? `"${String(W).slice(0, 50)}..."` : 'пусто',
        AA: AA || 'пусто'
      });
    }
  } catch (error) {
    console.error('❌ Ошибка проверки строк:', error.message);
  }
}

// ---- Быстрая диагностика ----------------------------------------------
function notes_QuickCheck() {
  try {
    console.log('⚡ Быстрая диагностика заметок');
    
    const config = getNotesConfig();
    console.log('📋 Конфигурация:', {
      sheet: config.SHEET_NAME,
      subdomain: config.AMO_SUBDOMAIN,
      hasToken: !!config.AMO_TOKEN
    });
    
    const sh = SpreadsheetApp.getActive().getSheetByName(config.SHEET_NAME);
    if (!sh) {
      console.log('❌ Лист не найден');
      return;
    }
    
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    console.log(`📊 Размер листа: ${lastRow} строк, ${lastCol} колонок`);
    
    if (lastRow < 2) {
      console.log('❌ Лист пуст');
      return;
    }
    
    // Проверяем заголовки
    const headers = sh.getRange(1, 1, 1, Math.min(lastCol, 30)).getValues()[0];
    console.log('📋 Заголовки (первые 30):', headers.map((h, i) => `${i+1}:${h}`).join(', '));
    
    // Проверяем несколько строк
    notes_CheckRows(2, 5);
    
  } catch (error) {
    console.error('❌ Ошибка быстрой диагностики:', error.message);
  }
}