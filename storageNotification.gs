function getUnretrievedSpecimens() {
  var result = new Map(); // Хранение оставленных образцов в словаре (email, [id1, id2,...])
  var sheet = SpreadsheetApp.getActive().getSheets()[0];
  var lastRowIndex = sheet.getLastRow();
  console.log('last row = ' + lastRowIndex);
  for (var i = 3; i <= lastRowIndex; i++) { // итерирование по строкам в таблице
    var id = sheet.getRange(i, 1).getDisplayValue();
   
    var storageStatus = sheet.getRange(i, 18).getDisplayValue();
    
    if (storageStatus.indexOf('хранится') != -1) {
      
      var email = getEmail(id);
      
      var specimen = {
        id: id,
        status: storageStatus,
        cifer: sheet.getRange(i, 8).getDisplayValue(),
        date: sheet.getRange(i, 13).getDisplayValue()
      };
      
      
      if (!result.has(email)) {
        result.set(email, [specimen]);
      } else {
        result.get(email).push(specimen);
      }
//      console.log(Utilities.formatString('%s %s %s %s', email, specimen.status, specimen.cifer, specimen.date));
    }
  }
  return result;
}

function remind(key, value) {
  if (key == 'error') {
    return;
  }
  var message = 'Необходимо забрать следующие образцы: \n';
  for (var i = 0; i < value.length; i++) {
    var spc = value[i];
    message += spc.cifer + ' ';
    message += spc.status + ' дата выполнения ';
    message += spc.date + '\n';
  }
  message += 'Пожалуйста, отметьте в журнале ЦКП с помощью примечания, что вы забрали образец, чтобы не получать эти письма в дальнейшем\n';
  message += 'https://docs.google.com/spreadsheets/d/1spLANRDFYjejlCZAcNsftUKLSdpKu853g8wfCKP79PY/edit?usp=sharing';
  try {
  GmailApp.sendEmail(key, 'Напоминание о забытых образцах ', message);
  } catch (e) {
    console.log('Ошибка отправки напоминания');
  }
}

function reminderToRetrieve(specimens) {
  console.log('Pre iterating');
  for (var entry of specimens) {
    console.log('Iterating map ' + entry[0]);
   remind(entry[0], entry[1]); 
  }
}

function test1() {
  reminderToRetrieve(getUnretrievedSpecimens());
}