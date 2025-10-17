// ============================== Заметки.gs ===========================
// Функции для отправки транскрипций и оценок в amoCRM
// Обрабатывает строки с заполненными колонками V (транскрипция) и W (оценка)

function sendNotesToAmoCRM() {
  try {
    console.log('--- Начало выполнения sendNotesToAmoCRM ---');
    
    // Загрузка параметров из листа "БД"
    const dbSheet = SpreadsheetApp.getActive().getSheetByName("БД");
    if (!dbSheet) throw new Error('Лист "БД" не найден');
    
    const amoToken = dbSheet.getRange("B2").getValue();
    let amoSubdomain = dbSheet.getRange("B3").getValue();
    const dataSheetName = dbSheet.getRange("B4").getValue();

    // Проверка обязательных параметров
    if (!amoToken || !amoSubdomain) throw new Error('Не указан токен или поддомен');
    amoSubdomain = amoSubdomain.replace(/\.amocrm\.(ru|eu)$/i, '').trim();
    
    if (!/^[a-z0-9-]+$/i.test(amoSubdomain)) {
      throw new Error(`Некорректный поддомен: ${amoSubdomain}`);
    }

    console.log(`Поддомен: ${amoSubdomain}`);
    console.log(`Лист данных: ${dataSheetName}`);

    // Загрузка листа с данными
    const dataSheet = SpreadsheetApp.getActive().getSheetByName(dataSheetName);
    if (!dataSheet) throw new Error(`Лист "${dataSheetName}" не найден`);

    // Чтение данных за один запрос
    const lastRow = dataSheet.getLastRow();
    const dataRange = dataSheet.getRange(2, 1, lastRow - 1, 27); // Столбцы A-AA
    const data = dataRange.getValues();
    
    // Фильтрация данных: оставляем только строки где AA заполнено и Z пустое
    const filteredData = data
      .map((row, index) => ({ row, index: index + 2 })) // Сохраняем оригинальный номер строки
      .filter(item => {
        const aaValue = item.row[26]; // Столбец AA (индекс 26 в массиве)
        const zValue = item.row[25];  // Столбец Z (индекс 25 в массиве)
        return (aaValue && !zValue);  // AA заполнено И Z пустое
      });

    console.log(`Всего строк: ${lastRow - 1}`);
    console.log(`Подходит для обработки: ${filteredData.length}`);

    // Если нет подходящих строк - завершаем выполнение
    if (filteredData.length === 0) {
      console.log('Нет строк для обработки');
      return;
    }

    // Обработка отфильтрованных данных
    filteredData.forEach(item => {
      const rowNumber = item.index;
      const rowData = item.row;
      
      const dealId = rowData[5];        // Столбец F
      const transcription = rowData[21]; // Столбец V
      const assessment = rowData[22];   // Столбец W

      console.log(`\nОбработка строки ${rowNumber}: dealId=${dealId}`);

      try {
        // Отправка транскрипции
        if (typeof transcription === 'string' && transcription.trim()) {
          addNoteToAmo(amoSubdomain, amoToken, dealId, transcription.trim());
          console.log('Транскрипция отправлена');
        }

        // Отправка оценки
        if (typeof assessment === 'string' && assessment.trim()) {
          addNoteToAmo(amoSubdomain, amoToken, dealId, assessment.trim());
          console.log('Оценка отправлена');
        }

        // Устанавливаем отметку времени в Z
        dataSheet.getRange(rowNumber, 26).setValue(new Date());
        console.log('Отметка времени установлена в колонку Z');

        // Задержка между запросами
        Utilities.sleep(500); // 0.5 секунд

      } catch (e) {
        console.error(`Ошибка в строке ${rowNumber}: ${e.message}`);
        // Устанавливаем отметку времени даже при ошибке
        dataSheet.getRange(rowNumber, 26).setValue(new Date());
        console.log('Отметка времени установлена в колонку Z после ошибки');
      }
    });

    console.log('--- Выполнение sendNotesToAmoCRM завершено ---');

  } catch (e) {
    console.error(`*** ГЛОБАЛЬНАЯ ОШИБКА: ${e.message}`);
  }
}

// Функция отправки заметки в amoCRM
function addNoteToAmo(subdomain, apiKey, dealId, noteText) {
  try {
    console.log(`--- Начало запроса к amoCRM ---`);
    const url = `https://${subdomain}.amocrm.ru/api/v4/leads/${dealId}/notes`;
    const payload = [{
      "note_type": "common",
      "text": noteText,
      "created_at": Math.floor(Date.now() / 1000)
    }];

    console.log(`URL: ${url}`);
    console.log(`Payload: ${JSON.stringify(payload)}`);

    const options = {
      'method': 'post',
      'headers': { 
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      },
      'payload': JSON.stringify(payload),
      'muteHttpExceptions': true
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const result = JSON.parse(response.getContentText());

    console.log(`Код ответа: ${responseCode}`);
    console.log(`Ответ сервера: ${JSON.stringify(result)}`);

    if (responseCode >= 400) {
      throw new Error(`amoCRM error: ${result.detail || response.getContentText()}`);
    }

    return result;

  } catch (e) {
    console.error(`Ошибка API: ${e.message}`);
    throw e;
  }
}

// Функция для тестирования отправки заметок
function testSendNotes() {
  try {
    console.log('--- Тест отправки заметок ---');
    
    const dbSheet = SpreadsheetApp.getActive().getSheetByName("БД");
    if (!dbSheet) throw new Error('Лист "БД" не найден');
    
    const amoToken = dbSheet.getRange("B2").getValue();
    let amoSubdomain = dbSheet.getRange("B3").getValue();
    
    console.log(`AmoCRM Token: ${amoToken ? 'Найден' : 'НЕ НАЙДЕН'}`);
    console.log(`AmoCRM Subdomain: ${amoSubdomain ? 'Найден' : 'НЕ НАЙДЕН'}`);
    
    if (!amoToken || !amoSubdomain) {
      throw new Error('Не настроены параметры amoCRM в листе БД');
    }
    
    amoSubdomain = amoSubdomain.replace(/\.amocrm\.(ru|eu)$/i, '').trim();
    console.log(`Очищенный поддомен: ${amoSubdomain}`);
    
    if (!/^[a-z0-9-]+$/i.test(amoSubdomain)) {
      throw new Error(`Некорректный поддомен: ${amoSubdomain}`);
    }
    
    console.log('✅ Настройки amoCRM корректны');
    console.log('--- Тест завершен ---');
    
  } catch (e) {
    console.error(`❌ Ошибка теста: ${e.message}`);
  }
}

// Функция для проверки готовности к отправке
function checkReadyToSend() {
  try {
    console.log('--- Проверка готовности к отправке ---');
    
    const dbSheet = SpreadsheetApp.getActive().getSheetByName("БД");
    const dataSheetName = dbSheet.getRange("B4").getValue();
    const dataSheet = SpreadsheetApp.getActive().getSheetByName(dataSheetName);
    
    if (!dataSheet) throw new Error(`Лист '${dataSheetName}' не найден`);
    
    const lastRow = dataSheet.getLastRow();
    const dataRange = dataSheet.getRange(2, 1, lastRow - 1, 27);
    const data = dataRange.getValues();
    
    let readyToSend = 0;
    let alreadySent = 0;
    let incomplete = 0;
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const transcription = row[21]; // V
      const assessment = row[22];    // W
      const aaValue = row[26];       // AA
      const zValue = row[25];        // Z
      
      if (transcription && assessment && aaValue && !zValue) {
        readyToSend++;
      } else if (zValue) {
        alreadySent++;
      } else if (transcription || assessment) {
        incomplete++;
      }
    }
    
    console.log(`Готово к отправке: ${readyToSend}`);
    console.log(`Уже отправлено: ${alreadySent}`);
    console.log(`Неполные данные: ${incomplete}`);
    console.log('--- Проверка завершена ---');
    
  } catch (e) {
    console.error(`❌ Ошибка проверки: ${e.message}`);
  }
}

// Функция для массовой отправки всех готовых заметок
function sendAllReadyNotes() {
  try {
    console.log('--- Массовая отправка всех готовых заметок ---');
    
    const dbSheet = SpreadsheetApp.getActive().getSheetByName("БД");
    const dataSheetName = dbSheet.getRange("B4").getValue();
    const dataSheet = SpreadsheetApp.getActive().getSheetByName(dataSheetName);
    
    if (!dataSheet) throw new Error(`Лист '${dataSheetName}' не найден`);
    
    const lastRow = dataSheet.getLastRow();
    const dataRange = dataSheet.getRange(2, 1, lastRow - 1, 27);
    const data = dataRange.getValues();
    
    let sentCount = 0;
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const transcription = row[21]; // V
      const assessment = row[22];    // W
      const aaValue = row[26];       // AA
      const zValue = row[25];        // Z
      
      if (transcription && assessment && aaValue && !zValue) {
        try {
          sendNotesToAmoCRM();
          sentCount++;
          console.log(`Отправлена группа ${sentCount}`);
          Utilities.sleep(1000); // Пауза между группами
        } catch (e) {
          console.error(`Ошибка отправки группы ${sentCount}: ${e.message}`);
        }
      }
    }
    
    console.log(`Всего отправлено групп: ${sentCount}`);
    console.log('--- Массовая отправка завершена ---');
    
  } catch (e) {
    console.error(`❌ Ошибка массовой отправки: ${e.message}`);
  }
}