function onOpen(e) {
  SpreadsheetApp
  .getActiveSheet()
  .getRange('H3')
  .getNextDataCell(SpreadsheetApp.Direction.DOWN)
  .activate();
  }

function sendFiles(id, link, cifer, text) {
  var email = getEmail(id);
  var message = 'По вашей заявке доступны следующие файлы '+ link;
  message += '\n' + text;
  GmailApp.sendEmail(email, 'Данные по Вашей заявке '+ cifer +' готовы', message);
}

function sendFilesButton() {
  var sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getIndex() != 1) return;
  var currentCell = SpreadsheetApp.getCurrentCell();
  var row = currentCell.getRowIndex();
  var lastRow = sheet.getLastRow();
  
  if (row < 3 || row > lastRow) return;
  
  var id = sheet.getRange(row, 1).getValue();
  var cifer = sheet.getRange(row, 8).getValue();
  var status = sheet.getRange(row, 19).getValue();
  var ui = SpreadsheetApp.getUi();
  var text = '';
  
  var response = ui.prompt('Введите ссылку на файл для отправки', 'Нажмите ДА, чтобы добавить комментарий к письму',  ui.ButtonSet.YES_NO_CANCEL);
  var link = response.getResponseText();
  var button = response.getSelectedButton();
  if (link == '' && (button == ui.Button.YES || button == ui.Button.NO)) {
    ui.alert('Ссылка не может быть пустой');
    return;
  }
  
  if (button == ui.Button.YES) {
    text = ui.prompt('Введите текст комментария').getResponseText();
  }
  try {
      sendFiles(id, link, cifer, text);
      ui.alert('Уведомление и ссылка отправлены заказчику');
     } catch (e) {
      ui.alert('Отправка не удалась. Что-то пошло не так :(');
      console.log("error");
    }
  var status = sheet.getRange(row, 19).getValue();
  sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground('#dcfadf');
    if (status.indexOf('?') != -1) {
    ui.alert('Статус содержит знаки вопроса. Не забудьте исправить нужные');
    } 
  
}

function getEmail(id) {
  try {
  var ss = SpreadsheetApp.openById(id);
  } catch(e) {
    return 'error';
  }
  var range = ss.getSheets()[0].getRange('B2:C30');
  var values = range.getValues();
  for (i = 0; i < values.length; i++) {
    if (values[i][0].indexOf('mail') != -1) {
    return values[i][1];
    }
  }
  return Session.getEffectiveUser().getEmail();

}


//test
function testSendFiles() {
  
}
