// ============================== заметки_original.gs ==============================
// Оригинальная рабочая версия из вашего кода

function sendNotesToAmoCRM() {
  const config = getConfigFromSheet();
  const sheet = SpreadsheetApp.getActive().getSheetByName(config.SHEET_NAME);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 27).getValues();
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const dealId = row[5]; // F
    const transcript = row[21]; // V
    const evaluation = row[22]; // W
    const aa = row[26]; // AA
    
    if (aa === 'OK') continue;
    if (!dealId) continue;
    if (!transcript && !evaluation) continue;
    
    if (transcript) {
      addNoteToAmo(dealId, transcript);
    }
    
    if (evaluation) {
      addNoteToAmo(dealId, evaluation);
    }
    
    sheet.getRange(i + 2, 26).setValue(new Date()); // Z
    sheet.getRange(i + 2, 27).setValue('OK'); // AA
    
    Utilities.sleep(500);
  }
}

function addNoteToAmo(dealId, noteText) {
  const config = getConfigFromSheet();
  const url = `https://${config.AMO_SUBDOMAIN}.amocrm.ru/api/v4/leads/${dealId}/notes`;
  
  const payload = [{
    note_type: 'common',
    params: {
      text: noteText
    }
  }];
  
  const options = {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${config.AMO_TOKEN}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload)
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();
    
    if (statusCode === 200) {
      console.log(`Заметка добавлена для сделки ${dealId}`);
    } else {
      console.log(`Ошибка ${statusCode}: ${response.getContentText()}`);
    }
  } catch (error) {
    console.log('Ошибка отправки заметки:', error);
  }
}