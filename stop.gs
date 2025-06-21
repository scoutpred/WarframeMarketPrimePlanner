function stopUpdateScript() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Item Tracker");
  sheet.getRange("M6").setValue("STOP");
  SpreadsheetApp.flush();
}
