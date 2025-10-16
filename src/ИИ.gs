// ============================== ИИ.gs (МОДЕРНИЗИРОВАННЫЙ) ==============================
// Совместимый слой для интеграции с другими модулями
// Ключевые улучшения:
// 1. Синхронизация с AMO.gs
// 2. Улучшенная обработка ошибок
// 3. Дополнительные утилиты для отладки

// ---- Совместимый доступ к конфигу из листа "БД" ------------------------
function getConfigFromSheet() {
  try {
    // Прямая реализация без рекурсии
    const dbSheet = SpreadsheetApp.getActive().getSheetByName("БД");
    if (!dbSheet) throw new Error('Лист "БД" не найден');
    
    let subdomain = dbSheet.getRange("B3").getValue();
    if (subdomain) {
      subdomain = String(subdomain).trim().replace(/^https?:\/\//i, '').replace(/\.amocrm\.(ru|eu)$/i, '');
      if (!/^[a-z0-9-]+$/i.test(subdomain)) {
        throw new Error(`Некорректный поддомен: ${subdomain}`);
      }
    }
    
    return {
      AMO_TOKEN: dbSheet.getRange("B2").getValue(),
      AMO_SUBDOMAIN: subdomain,
      SHEET_NAME: dbSheet.getRange("B4").getValue()
    };
  } catch (e) {
    console.error('[ИИ.gs] getConfigFromSheet error:', String(e && e.message || e));
    return null;
  }
}

// ---- Получение конфигурации для ИИ модуля -----------------------------
function getAIConfig() {
  try {
    const dbSheet = SpreadsheetApp.getActive().getSheetByName("БД");
    if (!dbSheet) throw new Error("Лист 'БД' не найден");
    
    return {
      AMO_TOKEN: dbSheet.getRange("B2").getValue(),
      AMO_SUBDOMAIN: dbSheet.getRange("B3").getValue(),
      SHEET_NAME: dbSheet.getRange("B4").getValue(),
      BOT_TOKEN: dbSheet.getRange("B5").getValue(),
      BOT_ID: dbSheet.getRange("B6").getValue(),
      SPREADSHEET_ID: SpreadsheetApp.getActive().getId()
    };
  } catch (error) {
    console.error('[ИИ.gs] getAIConfig error:', error.message);
    return null;
  }
}

// ---- Прокси-хелперы для работы с данными ------------------------------
function II_getLead(leadId) {
  try {
    // Используем функцию из AMO.gs
    return getAmoLead(leadId);
  } catch (e) {
    console.error('[ИИ.gs] II_getLead:', String(e));
    return null;
  }
}

function II_getNotesForLead(leadId) {
  try {
    // Используем функцию из AMO.gs
    return getAmoNotes('leads', leadId) || [];
  } catch (e) {
    console.error('[ИИ.gs] II_getNotesForLead:', String(e));
    return [];
  }
}

function II_getTasksForLead(leadId) {
  try {
    // Используем функцию из AMO.gs
    return getAmoTasks(leadId);
  } catch (e) {
    console.error('[ИИ.gs] II_getTasksForLead:', String(e));
    return 'Ошибка';
  }
}

// ---- Утилиты для работы с листами -------------------------------------
function II_getTargetSheet() {
  try {
    const config = getAIConfig();
    if (!config) return null;
    
    return SpreadsheetApp.openById(config.SPREADSHEET_ID).getSheetByName(config.SHEET_NAME);
  } catch (e) {
    console.error('[ИИ.gs] II_getTargetSheet:', String(e));
    return null;
  }
}

function II_getSheetData(sheetName) {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) return null;
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow < 2) return [];
    
    return sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  } catch (e) {
    console.error('[ИИ.gs] II_getSheetData:', String(e));
    return [];
  }
}

// ---- Утилиты для работы с аудио ---------------------------------------
function II_findAudioUrlInRow(sheet, row) {
  try {
    if (!sheet || !row) return '';
    
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
    console.error('[ИИ.gs] II_findAudioUrlInRow:', String(e));
    return '';
  }
}

function II_getRowDuration(sheet, row) {
  try {
    if (!sheet || !row) return 0;
    return Number(sheet.getRange(row, 12).getValue() || 0); // Колонка L
  } catch (e) {
    console.error('[ИИ.gs] II_getRowDuration:', String(e));
    return 0;
  }
}

// ---- Утилиты для работы с текстом -------------------------------------
function II_hasStopBot(text) {
  try {
    return /\bstop\s*bot\s*:\s*true\b/i.test(String(text || ''));
  } catch (e) {
    console.error('[ИИ.gs] II_hasStopBot:', String(e));
    return false;
  }
}

function II_isEmptyText(text) {
  try {
    const s = String(text || '').trim();
    return !s || 
           /^нужен\s*текст$/i.test(s) || 
           /^без\s*текста$/i.test(s) || 
           /^статус:\s*нужна\s*оценка/i.test(s) ||
           /^нужна\s*оценка$/i.test(s) || 
           /^без\s*оценки$/i.test(s);
  } catch (e) {
    console.error('[ИИ.gs] II_isEmptyText:', String(e));
    return true;
  }
}

function II_clipText(text, maxLength) {
  try {
    const t = String(text || '');
    return (maxLength && t.length > maxLength) ? t.slice(0, maxLength) : t;
  } catch (e) {
    console.error('[ИИ.gs] II_clipText:', String(e));
    return String(text || '');
  }
}

// ---- Утилиты для работы с ProTalk -------------------------------------
function II_getProtalkConfig() {
  try {
    const config = getAIConfig();
    if (!config) return null;
    
    return {
      BOT_TOKEN: config.BOT_TOKEN,
      BOT_ID: config.BOT_ID,
      API_URL: `https://eu1.api.pro-talk.ru/api/v1.0/ask/${config.BOT_TOKEN}`
    };
  } catch (e) {
    console.error('[ИИ.gs] II_getProtalkConfig:', String(e));
    return null;
  }
}

// ---- Диагностические функции ------------------------------------------
function II_selfcheck() {
  try {
    console.log('[ИИ.gs] === САМОПРОВЕРКА ===');
    
    const config = getAIConfig();
    if (!config) {
      console.log('❌ Конфигурация недоступна');
      return false;
    }
    
    console.log('✅ Конфигурация:', {
      sheet: config.SHEET_NAME,
      subdomain: config.AMO_SUBDOMAIN,
      hasToken: !!config.AMO_TOKEN,
      hasBotToken: !!config.BOT_TOKEN,
      hasBotId: !!config.BOT_ID
    });
    
    const sheet = II_getTargetSheet();
    if (!sheet) {
      console.log('❌ Целевой лист недоступен');
      return false;
    }
    
    console.log('✅ Целевой лист:', {
      name: sheet.getName(),
      lastRow: sheet.getLastRow(),
      lastCol: sheet.getLastColumn()
    });
    
    const protalkConfig = II_getProtalkConfig();
    if (!protalkConfig) {
      console.log('❌ Конфигурация ProTalk недоступна');
      return false;
    }
    
    console.log('✅ ProTalk конфигурация:', {
      hasToken: !!protalkConfig.BOT_TOKEN,
      hasBotId: !!protalkConfig.BOT_ID,
      apiUrl: protalkConfig.API_URL
    });
    
    console.log('[ИИ.gs] === САМОПРОВЕРКА ЗАВЕРШЕНА ===');
    return true;
  } catch (e) {
    console.error('[ИИ.gs] selfcheck failed:', String(e && e.message || e));
    return false;
  }
}

function II_testDataAccess() {
  try {
    console.log('[ИИ.gs] === ТЕСТ ДОСТУПА К ДАННЫМ ===');
    
    const sheet = II_getTargetSheet();
    if (!sheet) {
      console.log('❌ Лист недоступен');
      return;
    }
    
    const data = II_getSheetData(sheet.getName());
    console.log(`📊 Данные: ${data.length} строк`);
    
    if (data.length > 0) {
      const firstRow = data[0];
      console.log('📄 Первая строка:', {
        dealId: firstRow[5],      // F
        link: firstRow[1],        // B
        duration: firstRow[11],   // L
        transcript: firstRow[21], // V
        evaluation: firstRow[22]  // W
      });
    }
    
    console.log('[ИИ.gs] === ТЕСТ ЗАВЕРШЕН ===');
  } catch (e) {
    console.error('[ИИ.gs] testDataAccess failed:', String(e));
  }
}

function II_testAudioDetection() {
  try {
    console.log('[ИИ.gs] === ТЕСТ ОБНАРУЖЕНИЯ АУДИО ===');
    
    const sheet = II_getTargetSheet();
    if (!sheet) return;
    
    const data = II_getSheetData(sheet.getName());
    let audioFound = 0;
    let validUrls = 0;
    
    for (let i = 0; i < Math.min(data.length, 10); i++) {
      const row = i + 2;
      const url = II_findAudioUrlInRow(sheet, row);
      const duration = II_getRowDuration(sheet, row);
      
      if (url) {
        audioFound++;
        if (/^https?:\/\//i.test(url)) {
          validUrls++;
          console.log(`✅ Строка ${row}: ${url} (${duration}с)`);
        } else {
          console.log(`⚠️ Строка ${row}: невалидный URL - ${url}`);
        }
      }
    }
    
    console.log(`📊 Найдено аудио: ${audioFound}, валидных URL: ${validUrls}`);
    console.log('[ИИ.gs] === ТЕСТ ЗАВЕРШЕН ===');
  } catch (e) {
    console.error('[ИИ.gs] testAudioDetection failed:', String(e));
  }
}

// ---- Утилиты для работы с событиями -----------------------------------
function II_getRecentEvents(hours = 1) {
  try {
    const config = getAIConfig();
    if (!config) return [];
    
    const timeFrom = Math.floor((new Date().getTime() - hours * 3600 * 1000) / 1000);
    const url = `https://${config.AMO_SUBDOMAIN}.amocrm.ru/api/v4/events?filter[created_at][from]=${timeFrom}&limit=100`;
    
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { 'Authorization': 'Bearer ' + config.AMO_TOKEN },
      muteHttpExceptions: true,
      followRedirects: true
    });
    
    if (response.getResponseCode() !== 200) {
      console.error('❌ Ошибка получения событий:', response.getResponseCode());
      return [];
    }
    
    const json = JSON.parse(response.getContentText() || '{}');
    return json._embedded?.events || [];
  } catch (e) {
    console.error('[ИИ.gs] II_getRecentEvents:', String(e));
    return [];
  }
}

function II_analyzeEvents(events) {
  try {
    console.log('[ИИ.gs] === АНАЛИЗ СОБЫТИЙ ===');
    console.log(`📊 Всего событий: ${events.length}`);
    
    const eventTypes = {};
    const noteEvents = [];
    
    events.forEach(e => {
      eventTypes[e.type] = (eventTypes[e.type] || 0) + 1;
      if (e.type === 'note_added') {
        noteEvents.push(e);
      }
    });
    
    console.log('📈 Типы событий:', eventTypes);
    console.log(`📝 События с заметками: ${noteEvents.length}`);
    
    if (noteEvents.length > 0) {
      console.log('📄 Примеры событий с заметками:');
      noteEvents.slice(0, 3).forEach((event, i) => {
        console.log(`  ${i+1}. ID: ${event.id}, entity: ${event.entity_type}, note_type: ${event.value_after?.[0]?.note?.note_type}`);
      });
    }
    
    console.log('[ИИ.gs] === АНАЛИЗ ЗАВЕРШЕН ===');
  } catch (e) {
    console.error('[ИИ.gs] II_analyzeEvents:', String(e));
  }
}

// ---- Главная функция диагностики --------------------------------------
function II_fullDiagnostic() {
  try {
    console.log('[ИИ.gs] ===========================================');
    console.log('[ИИ.gs] ПОЛНАЯ ДИАГНОСТИКА СИСТЕМЫ');
    console.log('[ИИ.gs] ===========================================');
    
    // 1. Самопроверка
    const selfcheckOk = II_selfcheck();
    if (!selfcheckOk) {
      console.log('❌ Самопроверка не пройдена, останавливаемся');
      return;
    }
    
    // 2. Тест доступа к данным
    II_testDataAccess();
    
    // 3. Тест обнаружения аудио
    II_testAudioDetection();
    
    // 4. Анализ событий
    const events = II_getRecentEvents(2);
    II_analyzeEvents(events);
    
    console.log('[ИИ.gs] ===========================================');
    console.log('[ИИ.gs] ДИАГНОСТИКА ЗАВЕРШЕНА');
    console.log('[ИИ.gs] ===========================================');
  } catch (e) {
    console.error('[ИИ.gs] fullDiagnostic failed:', String(e));
  }
}