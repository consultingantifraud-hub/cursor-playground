// ============================== ииgs_original.gs ==============================
// Оригинальная рабочая версия из вашего кода

function ToProTalkBot() {
  const config = getConfigFromSheet();
  const sheet = SpreadsheetApp.getActive().getSheetByName(config.SHEET_NAME);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 30).getValues();
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const question = row[19]; // T
    const startDate = row[23]; // X
    const response = row[21]; // V
    
    if (!question) continue;
    if (startDate) continue;
    if (response) continue;
    
    // Отмечаем начало обработки
    sheet.getRange(i + 2, 24).setValue(new Date()); // X
    
    // Отправляем запрос в ProTalk
    const protalkResponse = askProTalk(question);
    
    // Записываем ответ
    sheet.getRange(i + 2, 22).setValue(protalkResponse); // V
    
    break; // Обрабатываем только одну строку за раз
  }
}

function askProTalk(question) {
  const config = getConfigFromSheet();
  const url = `https://eu1.api.pro-talk.ru/api/v1.0/ask/${config.BOT_TOKEN}`;
  
  const payload = {
    bot_id: config.BOT_ID,
    chat_id: 'chat_' + Date.now(),
    message: question
  };
  
  const options = {
    method: 'POST',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    timeout: 300000 // 5 минут
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();
    const json = JSON.parse(responseText);
    
    return json.done || json.text || 'Ошибка получения ответа';
  } catch (error) {
    console.log('Ошибка ProTalk:', error);
    return 'Ошибка: ' + error.toString();
  }
}