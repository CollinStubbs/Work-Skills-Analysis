function start() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getActiveSheet();
  
  var range = sheet.getDataRange().getValues();
  console.log(range);
}
