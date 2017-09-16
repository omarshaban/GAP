function createTrig() {
  SpreadsheetApp.getUi()
      .createMenu('Custom Menu')
      .addItem('Show alert', 'showAlert')
      .addToUi();
  
  var ss = SpreadsheetApp.getActive();
  
  // Creates an edit trigger for spreadsheet
  ScriptApp.newTrigger('editSheet')
     .forSpreadsheet(ss)
     .onEdit()
     .create();
};

function returnOldValue (sheet, e) {
  var oldValue = e.oldValue;
  
  var temp = new Date(1899, 11, 30, 0, 0, oldValue * 86400);
  var temp = Utilities.formatDate(new Date(temp.getTime() + 1000*60*60*6), 'GMT+2', "MMMM dd, yyyy");
  if (temp != 'January 01, 1970') {oldValue = temp};
  
  if (e.oldValue == undefined) {oldValue = ''}
  e.range.setValue(oldValue);
};

function showAlert(code) {
  
  var ui = SpreadsheetApp.getUi();
  
  if (code == 1) {
    var result = ui.alert('لا يمكنك تعديل هذه الخانة' ,ui.ButtonSet.OK);
  }
  
};
