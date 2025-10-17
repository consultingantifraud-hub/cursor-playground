// ============================== Оценка.gs ===========================
// Функция для оценки звонков через ProTalk API
// Обрабатывает строки с заполненной колонкой U (вопрос) и пустой колонкой Y (дата начала)

function ToProTalkBot() {
  let rowToProcess = null;
  let targetSheet = null;
  
  try {
    console.log('--- Начало выполнения ToProTalkBot ---');
    
    // Получение настроек из листа БД
    const dbSheet = SpreadsheetApp.getActive().getSheetByName("БД");
    if (!dbSheet) throw new Error("Лист 'БД' не найден");
   
    const botToken = dbSheet.getRange("B5").getValue();
    const botId = dbSheet.getRange("B6").getValue();
    const targetSheetName = dbSheet.getRange("B4").getValue();
   
    console.log(`Настройки: лист ${targetSheetName}, botId=${botId}`);
   
    // Получение целевого листа
    targetSheet = SpreadsheetApp.getActive().getSheetByName(targetSheetName);
    if (!targetSheet) throw new Error(`Лист '${targetSheetName}' не найден`);
   
    // Поиск первой незавершенной строки (начиная со второй)
    const dataRange = targetSheet.getDataRange();
    const values = dataRange.getValues();
    console.log(`Всего строк в листе: ${values.length}`);
   
    // Обновленные индексы колонок:
    // U = 20 (индекс 20), Y = 24 (индекс 24)
    for (let i = 1; i < values.length; i++) { 
      const question = values[i][20]; // Колонка U (индекс 20)
      const startDate = values[i][24]; // Колонка Y (индекс 24)
     
      if (question && !startDate) {
        rowToProcess = i + 1; // Номер строки в таблице
        break;
      }
      
      // Защита от зависания - ограничиваем поиск
      if (i > 1000) {
        console.log('Достигнут лимит поиска строк (1000)');
        break;
      }
    }
   
    if (!rowToProcess) {
      console.log("Нет строк для обработки оценки");
      return;
    }
   
    console.log(`Обрабатывается строка ${rowToProcess} для оценки`);
   
    // Запись даты начала в колонку Y (25-й столбец)
    const startDateCell = targetSheet.getRange(rowToProcess, 25); // Колонка Y
    startDateCell.setValue(new Date());
    console.log('Отметка времени установлена в колонку Y');
   
    // Получение вопроса из колонки U (21-й столбец)
    const questionCell = targetSheet.getRange(rowToProcess, 21); // Колонка U
    let question = questionCell.getValue();
    if (!question) throw new Error("Пустой вопрос в строке " + rowToProcess);
   
    console.log(`Вопрос: ${question.substring(0, 100)}...`);
   
    // Подготовка запросов
    question = '🏁##' + question;
    const requests = question.split("##").map(q => q.trim()).filter(Boolean);
    console.log(`Количество запросов: ${requests.length}`);
   
    // Настройки API
    const chatId = "chat_" + Date.now();
    const apiUrl = `https://us1.api.pro-talk.ru/api/v1.0/ask/${botToken}`;
    let lastResponse = null;
   
    // Таймаут 5 минут
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
   
    // Запись ответа в колонку W (23-й столбец)
    const cleanResponse = lastResponse ? lastResponse.replace(/\*/g, '') : '';
    targetSheet.getRange(rowToProcess, 23).setValue(cleanResponse); // Колонка W
    console.log(`Оценка записана в строку ${rowToProcess}`);
    console.log('--- Выполнение ToProTalkBot завершено ---');
   
  } catch (e) {
    console.error(`*** ОШИБКА в ToProTalkBot: ${e.message}`);
   
    // Гарантированная запись ошибки в колонку W
    if (rowToProcess && targetSheet) {
      try {
        targetSheet.getRange(rowToProcess, 23)
                   .setValue("Разговор не распознан"); // Колонка W
        console.log('Ошибка записана в колонку W');
      } catch (setError) {
        console.error("Ошибка записи статуса: " + setError.message);
      }
    }
  }
}

// Функция для тестирования оценки
function testEvaluation() {
  try {
    console.log('--- Тест функции оценки ---');
    
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