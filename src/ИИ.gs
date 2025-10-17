// ============================== ИИ.gs ===========================
// Функция для транскрипции звонков через ProTalk API
// Обрабатывает строки с заполненной колонкой U (вопрос) и пустой колонкой Y (дата начала)

function chatToProTalkBot() {
  let rowToProcess = null;
  let targetSheet = null;
  
  try {
    console.log('--- Начало выполнения chatToProTalkBot ---');
    
    const dbSheet = SpreadsheetApp.getActive().getSheetByName("БД");
    if (!dbSheet) throw new Error("Лист 'БД' не найден");
   
    const botToken = dbSheet.getRange("B5").getValue();
    const botId = dbSheet.getRange("B6").getValue();
    const targetSheetName = dbSheet.getRange("B4").getValue();
   
    console.log(`Настройки: лист ${targetSheetName}, botId=${botId}`);
   
    targetSheet = SpreadsheetApp.getActive().getSheetByName(targetSheetName);
    if (!targetSheet) throw new Error(`Лист '${targetSheetName}' не найден`);
   
    const dataRange = targetSheet.getDataRange();
    const values = dataRange.getValues();
    console.log(`Всего строк в листе: ${values.length}`);
   
    // Поиск строки для обработки (колонка U = 20, Y = 24)
    for (let i = 1; i < values.length; i++) {
      const question = values[i][20]; // U
      const startDate = values[i][24]; // Y
     
      if (question && !startDate) {
        rowToProcess = i + 1;
        break;
      }
      
      // Защита от зависания - ограничиваем поиск
      if (i > 1000) {
        console.log('Достигнут лимит поиска строк (1000)');
        break;
      }
    }
   
    if (!rowToProcess) {
      console.log("Нет строк для обработки транскрипции");
      return;
    }
   
    console.log(`Обрабатывается строка ${rowToProcess} для транскрипции`);
    
    // Запись даты начала в колонку Y (25-й столбец)
    const startDateCell = targetSheet.getRange(rowToProcess, 25);
    startDateCell.setValue(new Date());
    console.log('Отметка времени установлена в колонку Y');
   
    const questionCell = targetSheet.getRange(rowToProcess, 21);
    let question = questionCell.getValue();
    if (!question) throw new Error("Пустой вопрос в строке " + rowToProcess);
   
    console.log(`Вопрос: ${question.substring(0, 100)}...`);
   
    question = '🏁##' + question;
    const requests = question.split("##").map(q => q.trim()).filter(Boolean);
    console.log(`Количество запросов: ${requests.length}`);
   
    const chatId = "chat_" + Date.now();
    const apiUrl = `https://eu1.api.pro-talk.ru/api/v1.0/ask/${botToken}`;
    let lastResponse = null;
   
    const TIMEOUT_MS = 300000;
   
    for (const q of requests) {
      try {
        console.log(`Отправка запроса: ${q.substring(0, 50)}...`);
        
        const payload = {
          bot_id: botId,
          chat_id: chatId,
          message: q
        };
        
        const response = UrlFetchApp.fetch(apiUrl, {
          method: "post",
          contentType: "application/json",
          muteHttpExceptions: true,
          timeout: 30000 // Уменьшенный таймаут - 30 секунд
        });
       
        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();
        console.log(`Код ответа: ${responseCode}`);
       
        if (responseCode !== 200) {
          throw new Error(`HTTP ${responseCode}: ${responseText}`);
        }
       
        const result = JSON.parse(responseText);
        lastResponse = result.done;
        console.log(`Получен ответ: ${lastResponse ? 'OK' : 'ERROR'}`);
        
        // Уменьшенная задержка
        Utilities.sleep(500);
       
      } catch (e) {
        console.error(`Ошибка в запросе: ${e.message}`);
        // Не прерываем выполнение при ошибке одного запроса
        lastResponse = "Ошибка обработки";
        break;
      }
    }
   
    const cleanResponse = lastResponse ? lastResponse.replace(/\*/g, '') : '';
    targetSheet.getRange(rowToProcess, 22).setValue(cleanResponse);
    console.log(`Транскрипция записана в строку ${rowToProcess}`);
    console.log('--- Выполнение chatToProTalkBot завершено ---');
   
  } catch (e) {
    console.error(`*** ОШИБКА в chatToProTalkBot: ${e.message}`);
    
    // Гарантированная запись ошибки
    if (rowToProcess && targetSheet) {
      try {
        targetSheet.getRange(rowToProcess, 22).setValue("Разговор не распознан");
        console.log('Ошибка записана в колонку V');
      } catch (setError) {
        console.error(`Ошибка записи статуса: ${setError.message}`);
      }
    }
  }
}

// Функция для тестирования транскрипции
function testTranscription() {
  try {
    console.log('--- Тест функции транскрипции ---');
    
    const dbSheet = SpreadsheetApp.getActive().getSheetByName("БД");
    if (!dbSheet) throw new Error("Лист 'БД' не найден");
    
    const botToken = dbSheet.getRange("B5").getValue();
    const botId = dbSheet.getRange("B6").getValue();
    
    console.log(`Bot Token: ${botToken ? 'Найден' : 'НЕ НАЙДЕН'}`);
    console.log(`Bot ID: ${botId ? 'Найден' : 'НЕ НАЙДЕН'}`);
    
    if (!botToken || !botId) {
      throw new Error('Не настроены параметры ProTalk в листе БД');
    }
    
    console.log('✅ Настройки ProTalk корректны');
    console.log('--- Тест завершен ---');
    
  } catch (e) {
    console.error(`❌ Ошибка теста: ${e.message}`);
  }
}

// Функция для проверки статуса обработки
function checkProcessingStatus() {
  try {
    console.log('--- Проверка статуса обработки ---');
    
    const dbSheet = SpreadsheetApp.getActive().getSheetByName("БД");
    const targetSheetName = dbSheet.getRange("B4").getValue();
    const targetSheet = SpreadsheetApp.getActive().getSheetByName(targetSheetName);
    
    if (!targetSheet) throw new Error(`Лист '${targetSheetName}' не найден`);
    
    const dataRange = targetSheet.getDataRange();
    const values = dataRange.getValues();
    
    let pendingTranscription = 0;
    let pendingEvaluation = 0;
    let completed = 0;
    
    for (let i = 1; i < values.length; i++) {
      const question = values[i][20]; // U
      const transcription = values[i][21]; // V
      const evaluation = values[i][22]; // W
      const startDate = values[i][24]; // Y
      
      if (question && !startDate) {
        pendingTranscription++;
      } else if (transcription && !evaluation) {
        pendingEvaluation++;
      } else if (transcription && evaluation) {
        completed++;
      }
    }
    
    console.log(`Ожидают транскрипции: ${pendingTranscription}`);
    console.log(`Ожидают оценки: ${pendingEvaluation}`);
    console.log(`Завершено: ${completed}`);
    console.log('--- Проверка завершена ---');
    
  } catch (e) {
    console.error(`❌ Ошибка проверки: ${e.message}`);
  }
}