// ============================== Мониторинг.gs (МОДЕРНИЗИРОВАННЫЙ) ==============================
// Диагностический вывод событий amoCRM с улучшенным анализом
// Ключевые улучшения:
// 1. Синхронизация с AMO.gs
// 2. Расширенный анализ событий
// 3. Детальная статистика

// ---- Получение конфигурации (синхронизировано с AMO.gs) --------------
function getMonitoringConfig() {
  try {
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
  } catch (error) {
    console.error('❌ Ошибка конфигурации мониторинга:', error.message);
    throw error;
  }
}

// ---- Универсальный запрос к amoCRM ------------------------------------
function amoRequest(url, method = 'GET', body = null) {
  try {
    const options = {
      method: method,
      headers: { 
        'Authorization': `Bearer ${getMonitoringConfig().AMO_TOKEN}`,
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

// ---- Основная функция логирования событий (МОДЕРНИЗИРОВАННАЯ) --------
function logEvents() {
  try {
    console.log('🔍 Начало логирования событий amoCRM');
    const config = getMonitoringConfig();
    const timeFrom = Math.floor((new Date().getTime() - 150 * 60 * 1000) / 1000); // последние 2.5 часа
    let url = `https://${config.AMO_SUBDOMAIN}.amocrm.ru/api/v4/events?filter[created_at][from]=${timeFrom}&limit=250`;
    let allEvents = [];
    let pages = 0;
    const maxPages = 10; // ограничиваем для отладки
    
    console.log(`📊 Получаем события с ${new Date(timeFrom * 1000).toLocaleString()}`);
    
    // Получаем события постранично
    while (url && pages < maxPages) {
      try {
        const response = amoRequest(url);
        if (!response?._embedded?.events) break;
        
        allEvents = allEvents.concat(response._embedded.events);
        url = response._links?.next?.href || '';
        pages++;
        
        console.log(`📄 Страница ${pages}: ${response._embedded.events.length} событий`);
      } catch (error) {
        console.error(`❌ Ошибка получения страницы ${pages + 1}:`, error.message);
        break;
      }
    }
    
    console.log(`📊 Всего получено событий: ${allEvents.length} за ${pages} страниц`);
    
    // Анализируем события
    analyzeEvents(allEvents);
    
    // Показываем примеры событий с заметками
    showNoteEventsExamples(allEvents);
    
  } catch (error) {
    console.error('❌ Ошибка логирования событий:', error.message);
  }
}

// ---- Анализ событий ---------------------------------------------------
function analyzeEvents(events) {
  try {
    console.log('📈 === АНАЛИЗ СОБЫТИЙ ===');
    
    // Статистика по типам событий
    const eventTypes = {};
    const entityTypes = {};
    const noteEvents = [];
    const callEvents = [];
    
    events.forEach(event => {
      // Типы событий
      eventTypes[event.type] = (eventTypes[event.type] || 0) + 1;
      
      // Типы сущностей
      if (event.entity_type) {
        entityTypes[event.entity_type] = (entityTypes[event.entity_type] || 0) + 1;
      }
      
      // События с заметками
      if (event.type === 'note_added') {
        noteEvents.push(event);
        
        // Проверяем, является ли это звонком
        const note = event.value_after?.[0]?.note;
        if (note) {
          const noteType = String(note.note_type || '').toLowerCase();
          const params = note.params || {};
          
          // Стандартные звонки
          if (/^(call_in|call_out)$/i.test(noteType)) {
            callEvents.push({ ...event, callType: 'standard' });
          }
          // Гибридные звонки
          else if (noteType === 'common') {
            const hasLink = params.link && /^https?:\/\//i.test(String(params.link));
            const hasDuration = params.duration && Number(params.duration) > 0;
            const hasCallStatus = params.call_status !== undefined;
            
            if (hasLink || hasDuration || hasCallStatus) {
              callEvents.push({ ...event, callType: 'hybrid' });
            }
          }
        }
      }
    });
    
    console.log('📊 Типы событий:', eventTypes);
    console.log('📊 Типы сущностей:', entityTypes);
    console.log(`📝 События с заметками: ${noteEvents.length}`);
    console.log(`📞 События со звонками: ${callEvents.length}`);
    
    // Показываем распределение по времени
    showTimeDistribution(events);
    
    // Показываем примеры звонков
    if (callEvents.length > 0) {
      console.log('📞 === ПРИМЕРЫ ЗВОНКОВ ===');
      callEvents.slice(0, 5).forEach((event, i) => {
        const note = event.value_after?.[0]?.note;
        console.log(`${i+1}. ${event.callType} звонок:`, {
          id: event.id,
          entity_type: event.entity_type,
          entity_id: event.entity_id,
          note_id: note?.id,
          note_type: note?.note_type,
          has_link: !!note?.params?.link,
          duration: note?.params?.duration
        });
      });
    }
    
  } catch (error) {
    console.error('❌ Ошибка анализа событий:', error.message);
  }
}

// ---- Показ примеров событий с заметками -------------------------------
function showNoteEventsExamples(events) {
  try {
    const noteEvents = events.filter(e => e.type === 'note_added');
    
    if (noteEvents.length === 0) {
      console.log('📝 Событий с заметками не найдено');
      return;
    }
    
    console.log('📝 === ПРИМЕРЫ СОБЫТИЙ С ЗАМЕТКАМИ ===');
    
    // Показываем первые 10 событий
    noteEvents.slice(0, 10).forEach((event, i) => {
      const note = event.value_after?.[0]?.note;
      console.log(`${i+1}. Событие ID ${event.id}:`, {
        entity_type: event.entity_type,
        entity_id: event.entity_id,
        note_id: note?.id,
        note_type: note?.note_type,
        created_at: new Date(event.created_at * 1000).toLocaleString(),
        has_params: !!note?.params,
        params_keys: note?.params ? Object.keys(note.params) : []
      });
    });
    
    // Показываем статистику по типам заметок
    const noteTypes = {};
    noteEvents.forEach(event => {
      const noteType = event.value_after?.[0]?.note?.note_type;
      if (noteType) {
        noteTypes[noteType] = (noteTypes[noteType] || 0) + 1;
      }
    });
    
    console.log('📊 Типы заметок:', noteTypes);
    
  } catch (error) {
    console.error('❌ Ошибка показа примеров:', error.message);
  }
}

// ---- Анализ временного распределения ----------------------------------
function showTimeDistribution(events) {
  try {
    console.log('⏰ === ВРЕМЕННОЕ РАСПРЕДЕЛЕНИЕ ===');
    
    const now = new Date();
    const timeRanges = {
      'Последние 15 мин': 0,
      '15-30 мин назад': 0,
      '30-60 мин назад': 0,
      '1-2 часа назад': 0,
      '2+ часа назад': 0
    };
    
    events.forEach(event => {
      const eventTime = new Date(event.created_at * 1000);
      const diffMinutes = (now - eventTime) / (1000 * 60);
      
      if (diffMinutes <= 15) timeRanges['Последние 15 мин']++;
      else if (diffMinutes <= 30) timeRanges['15-30 мин назад']++;
      else if (diffMinutes <= 60) timeRanges['30-60 мин назад']++;
      else if (diffMinutes <= 120) timeRanges['1-2 часа назад']++;
      else timeRanges['2+ часа назад']++;
    });
    
    Object.entries(timeRanges).forEach(([range, count]) => {
      console.log(`${range}: ${count} событий`);
    });
    
  } catch (error) {
    console.error('❌ Ошибка анализа времени:', error.message);
  }
}

// ---- Мониторинг в реальном времени ------------------------------------
function monitorRealtime() {
  try {
    console.log('🔄 Мониторинг в реальном времени (последние 5 минут)');
    const config = getMonitoringConfig();
    const timeFrom = Math.floor((new Date().getTime() - 5 * 60 * 1000) / 1000);
    const url = `https://${config.AMO_SUBDOMAIN}.amocrm.ru/api/v4/events?filter[created_at][from]=${timeFrom}&limit=100`;
    
    const response = amoRequest(url);
    if (!response?._embedded?.events) {
      console.log('📊 Событий за последние 5 минут не найдено');
      return;
    }
    
    const events = response._embedded.events;
    console.log(`📊 Событий за последние 5 минут: ${events.length}`);
    
    // Анализируем только новые события
    analyzeEvents(events);
    
  } catch (error) {
    console.error('❌ Ошибка мониторинга в реальном времени:', error.message);
  }
}

// ---- Проверка здоровья системы ----------------------------------------
function checkSystemHealth() {
  try {
    console.log('🏥 === ПРОВЕРКА ЗДОРОВЬЯ СИСТЕМЫ ===');
    
    const config = getMonitoringConfig();
    
    // 1. Проверка конфигурации
    console.log('✅ Конфигурация:', {
      hasToken: !!config.AMO_TOKEN,
      subdomain: config.AMO_SUBDOMAIN,
      sheetName: config.SHEET_NAME
    });
    
    // 2. Проверка доступности API
    try {
      const testUrl = `https://${config.AMO_SUBDOMAIN}.amocrm.ru/api/v4/account`;
      const response = amoRequest(testUrl);
      console.log('✅ API доступен');
    } catch (error) {
      console.log('❌ API недоступен:', error.message);
    }
    
    // 3. Проверка целевого листа
    try {
      const sheet = SpreadsheetApp.getActive().getSheetByName(config.SHEET_NAME);
      if (sheet) {
        console.log('✅ Целевой лист доступен:', {
          name: sheet.getName(),
          lastRow: sheet.getLastRow(),
          lastCol: sheet.getLastColumn()
        });
      } else {
        console.log('❌ Целевой лист не найден');
      }
    } catch (error) {
      console.log('❌ Ошибка доступа к листу:', error.message);
    }
    
    // 4. Проверка последних событий
    try {
      const timeFrom = Math.floor((new Date().getTime() - 60 * 60 * 1000) / 1000);
      const url = `https://${config.AMO_SUBDOMAIN}.amocrm.ru/api/v4/events?filter[created_at][from]=${timeFrom}&limit=10`;
      const response = amoRequest(url);
      
      if (response?._embedded?.events) {
        console.log(`✅ События доступны: ${response._embedded.events.length} за последний час`);
      } else {
        console.log('⚠️ События не найдены за последний час');
      }
    } catch (error) {
      console.log('❌ Ошибка получения событий:', error.message);
    }
    
    console.log('🏥 === ПРОВЕРКА ЗАВЕРШЕНА ===');
    
  } catch (error) {
    console.error('❌ Ошибка проверки здоровья:', error.message);
  }
}

// ---- Экспорт событий в лист -------------------------------------------
function exportEventsToSheet(hours = 2) {
  try {
    console.log(`📤 Экспорт событий за последние ${hours} часов`);
    const config = getMonitoringConfig();
    const timeFrom = Math.floor((new Date().getTime() - hours * 3600 * 1000) / 1000);
    let url = `https://${config.AMO_SUBDOMAIN}.amocrm.ru/api/v4/events?filter[created_at][from]=${timeFrom}&limit=250`;
    let allEvents = [];
    let pages = 0;
    const maxPages = 20;
    
    // Получаем события
    while (url && pages < maxPages) {
      try {
        const response = amoRequest(url);
        if (!response?._embedded?.events) break;
        
        allEvents = allEvents.concat(response._embedded.events);
        url = response._links?.next?.href || '';
        pages++;
      } catch (error) {
        console.error(`❌ Ошибка получения страницы ${pages + 1}:`, error.message);
        break;
      }
    }
    
    console.log(`📊 Получено ${allEvents.length} событий за ${pages} страниц`);
    
    // Создаем лист для экспорта
    const sheetName = `Events_Export_${new Date().toISOString().slice(0, 10)}`;
    let sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    
    if (sheet) {
      sheet.clear();
    } else {
      sheet = SpreadsheetApp.getActive().insertSheet(sheetName);
    }
    
    // Заголовки
    const headers = [
      'ID', 'Type', 'Entity Type', 'Entity ID', 'Created At', 'User ID',
      'Note ID', 'Note Type', 'Has Link', 'Duration', 'Call Status'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Данные
    const rows = allEvents.map(event => {
      const note = event.value_after?.[0]?.note;
      return [
        event.id,
        event.type,
        event.entity_type || '',
        event.entity_id || '',
        new Date(event.created_at * 1000),
        event.user_id || '',
        note?.id || '',
        note?.note_type || '',
        note?.params?.link ? 'Да' : 'Нет',
        note?.params?.duration || '',
        note?.params?.call_status || ''
      ];
    });
    
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }
    
    console.log(`✅ Экспорт завершен: ${rows.length} строк в листе "${sheetName}"`);
    
  } catch (error) {
    console.error('❌ Ошибка экспорта:', error.message);
  }
}