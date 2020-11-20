function getSpreadSheet () {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss;
}

function getMailTo (ss) {
  const sheet = ss.getSheetByName('送信リスト');
  const lastRow = findLastRow(1, sheet);
  const sheetValue = sheet.getRange(1, 1, lastRow).getValues();
  var res = [];
  for (var idx = 0;idx < sheetValue.length; idx++) {
    res.push(sheetValue[idx][0])
  }
  return res;
}

/*
function getMailFrom (ss) {
  const sheet = ss.getSheetByName('送信リスト');
  const mailFrom = sheet.getRange(1, 2).getValues()[0][0];
  const fromName = sheet.getRange(2, 2).getValues()[0][0];

  return [mailFrom, fromName];
}
*/

function getMailTile (ss) {
  const sheet = ss.getSheetByName('本文');
  const mailTitle = sheet.getRange(1, 2).getValues()[0][0];

  return mailTitle;
}

function getStartDoc (ss, rowNum) {
  const sheet = ss.getSheetByName('本文');
  const sheetValue = sheet.getRange(2, 2).getValues()[0][0];
  var startDoc = DocumentApp.openById(sheetValue).getBody().getText();

  // 本文の時刻などの置換
  const sheet2 = ss.getSheetByName('日程');
  const startTime =  Utilities.formatDate(sheet2.getRange(rowNum, 2).getValues()[0][0], 'Asia/Tokyo', 'HH:mm');
  const endTime = Utilities.formatDate(sheet2.getRange(rowNum, 3).getValues()[0][0], 'Asia/Tokyo', 'HH:mm');
  const restStartTime = Utilities.formatDate(sheet2.getRange(rowNum, 4).getValues()[0][0], 'Asia/Tokyo', 'HH:mm');
  const restEndTime = Utilities.formatDate(sheet2.getRange(rowNum, 5).getValues()[0][0], 'Asia/Tokyo', 'HH:mm');
  const workHour = parseInt(sheet2.getRange(rowNum, 6).getValues()[0][0]).toString();
  const restHour = parseInt(sheet2.getRange(rowNum, 7).getValues()[0][0]).toString();

  var startDoc = startDoc.replace(/{開始時刻}/, startTime);
  var startDoc = startDoc.replace(/{終了時刻}/, endTime);
  var startDoc = startDoc.replace(/{休憩開始時刻}/, restStartTime);
  var startDoc = startDoc.replace(/{休憩終了時刻}/, restEndTime);
  var startDoc = startDoc.replace(/{実働時間}/, workHour);
  var startDoc = startDoc.replace(/{休憩時間}/, restHour);

  return startDoc;
}

function getEndDoc(ss) {
  const sheet = ss.getSheetByName('本文');
  const sheetValue = sheet.getRange(3, 2).getValues()[0][0];
  var endDoc = DocumentApp.openById(sheetValue).getBody().getText();
  Logger.log(endDoc);
  return endDoc;
}

function getSchedule (ss) {
  const sheet = ss.getSheetByName('日程');
  const lastRow = findLastRow(1, sheet);
  const sheetValue = sheet.getRange(2, 2, lastRow - 1, 13).getValues();

  return sheetValue;
}

function getUnsendSchedule (schedule) {
  var today = new Date();
  var startRows = [];
  var endRows = [];
  for (var idx = 0; idx < schedule.length; idx++) {
    if(today.getTime() > schedule[idx][9] && !(schedule[idx][6])) {
      startRows.push(idx);
    }
    if(today.getTime() > schedule[idx][10] && !(schedule[idx][7])) {
      endRows.push(idx);
    }
  }
  return [startRows, endRows];
}

function findLastRow(col, sheet) {

  //指定の列を二次元配列に格納する※シート全体の最終行までとする
  var ColValues = sheet.getRange(1, col, sheet.getLastRow(), 1).getValues();
  // Logger.log(ColValues);// [[列2], [2], [2], [2], [2], [], []]

  //二次元配列のなかで、データが存在する要素のlengthを取得する
  var lastRow = ColValues.filter(String).length;
  return lastRow;// 5

}

function sendEndMail(ss, endRows, title, mailTo) {
  var endDoc = getEndDoc(ss);
  for (var idx = 0; idx < endRows.length; idx++) {
    GmailApp.sendEmail(
　　　mailTo,
　　　title,
　　　endDoc,
    );
    // フラグ管理
    const sheet = ss.getSheetByName('日程');
    const res = sheet.getRange(endRows[idx] + 2, 9).setValue(true);
  }
}


function sendStartMail(ss, startRows, title, mailTo) {
  for (var idx = 0; idx < startRows.length; idx++) {
    GmailApp.sendEmail(
　　　mailTo,
　　　title,
　　　getStartDoc(ss, startRows[idx] + 2)
    );
    // フラグ管理
    const sheet = ss.getSheetByName('日程');
    const res = sheet.getRange(startRows[idx] + 2, 8).setValue(true);
  }
}

function main() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var schedule = getSchedule(ss);
  var [startRows, endRows] = getUnsendSchedule(schedule);
  var mailTitle = getMailTile(ss);
  var mailTo = getMailTo(ss);

  // メール送信
  sendStartMail(ss, startRows, mailTitle, mailTo);
  sendEndMail(ss, endRows, mailTitle, mailTo);
}
