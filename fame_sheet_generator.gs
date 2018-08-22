function fillFromLastWeek(){
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  checkForExistingSheet(activeSpreadsheet);
  
  var sheetsArray = activeSpreadsheet.getSheets();
  var latestSheet = activeSpreadsheet.setActiveSheet(sheetsArray[sheetsArray.length-1]);
  var names = latestSheet.getRange('A1:A51');
  var lastWeekTotalFame = latestSheet.getRange('C1:C51');
  var lastWeekFame = latestSheet.getRange('B1:B51');
  var miscDetails = latestSheet.getRange('G1:L51');
  var legend = latestSheet.getRange('G58:J58');
  
  newSheetID = startNewForm(activeSpreadsheet);
  
  var sheetsArray = activeSpreadsheet.getSheets();
  var newSheet = activeSpreadsheet.getSheets()[sheetsArray.length-1]
  
  setSheetFormatting(newSheet);
  
  buildStatusValidation(newSheet);
  
  newSheet.getRange('A1:A51').setValues(names.getValues())
  .setBackgrounds(names.getBackgrounds());
  newSheet.getRange('B1:B51').setFormulas(lastWeekFame.getFormulas());
  newSheet.getRange('D1:D51').setValues(lastWeekFame.getValues());
  newSheet.getRange('E1:E51').setValues(lastWeekTotalFame.getValues());
  newSheet.getRange('E1').setValue('Last Week Total');
  newSheet.getRange('D1').setValue('Last Week Fame');
  newSheet.getRange('C1').setValue('Total Fame');
  newSheet.getRange('B1').setValue('Week Fame');
  newSheet.getRange('A1:L58').setHorizontalAlignment('center');
  newSheet.getRange('A1:E1').setFontWeight('bold').setFontStyle('italic')
  .setBorder(false, false, true, false, false, false);
  newSheet.getRange('G1:L51').setBackgrounds(miscDetails.getBackgrounds())
  .setValues(miscDetails.getValues());
  newSheet.getRange('G58:J58').setBackgrounds(legend.getBackgrounds())
  .setValues(legend.getValues());
}

function checkForExistingSheet(spreadsheet){
  var date = Utilities.formatDate(new Date(), 'GMT-8', "M/d/yy")
  var newSheet = spreadsheet.getSheetByName(date);

    if (newSheet) {
        spreadsheet.deleteSheet(newSheet);
    }
}

function startNewForm(spreadsheet) {
  var date = Utilities.formatDate(new Date(), 'GMT-8', "M/d/yy")

  newSheet = spreadsheet.insertSheet();
  newSheet.setName(date);
  
  return newSheet.getSheetId();
}

function setSheetFormatting(sheet){
  var thisWeekRange = sheet.getRange('B2:B51');
  var lastWeekRange = sheet.getRange('D2:D51');
  
  var aboveFameThresholdRule = SpreadsheetApp.newConditionalFormatRule()
.whenNumberLessThan(1500)
.setBackground('#f4c7c3')
.setRanges([thisWeekRange, lastWeekRange])
.build();
  
  var belowFameThresholdRule = SpreadsheetApp.newConditionalFormatRule()
.whenNumberGreaterThanOrEqualTo(1500)
.setBackground('#b7e1cd')
.setRanges([thisWeekRange, lastWeekRange])
.build();
  
  var rules = sheet.getConditionalFormatRules();
  rules.push(belowFameThresholdRule);
  rules.push(aboveFameThresholdRule);
  sheet.setConditionalFormatRules(rules);
}

function buildStatusValidation(sheet){
  var accountStatus = sheet.getRange('G2:G51');
  var dmStatus = sheet.getRange('I2:J51');
  var resultOfDM = sheet.getRange('J2:J51');
  
  accountStatusRule = SpreadsheetApp.newDataValidation().requireValueInList(
    ['New Recruit', 'Existing Penalty', 'Main Transfer', 'Has Pass']).build();
  dmStatusRule = SpreadsheetApp.newDataValidation().requireValueInList(
    ['Contacted', 'Pending', 'Failed']).build();
  resultOfDMRule = SpreadsheetApp.newDataValidation().requireValueInList(
    ['Penalty', 'Pass', 'Late Pass', 'Warned', 'Left', 'Kicked']).build();
  
  accountStatus.setDataValidation(accountStatusRule);
  dmStatus.setDataValidation(dmStatusRule);
  resultOfDM.setDataValidation(resultOfDMRule);
}
