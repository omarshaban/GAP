function activitySimpleEdit(sheet, e, index, editArch, editAxioms, props) {
  var user = Session.getActiveUser();
  var date = Utilities.formatDate(new Date(), 'GMT+2', 'dd-MM-yyyy HH:mm:ss');
  
  var dbSheet = SpreadsheetApp.openById(dbID).getSheetByName('database');
  var dbDummy = SpreadsheetApp.getActive().getSheetByName('database');
  
  var cache = CacheService.getScriptCache();
  if (cache.get('code1') == null) {cacheMe(1, cache)}
  var dbDRV = JSON.parse(cache.get('dbDRV'));
  
  var ID = editAxioms.ID;
  var dbR = editAxioms.dbR;
  
  if (e.value == undefined) {
    e.value = '*deleted*';
  } else if (index<4) {
    if (e.range.getNumberFormat() != "mmmm dd, yyyy") {
      e.range.setValue(e.oldValue).setNumberFormat("mmmm dd, yyyy");
      colorMe(sheet, [[e.range.getRow(), e.range.getColumn()]], '#e38888', props);
      return;
    } else {
      var oldDate = new Date(dbDRV[dbR][editArch.standardDB[index]]);
      e.oldValue = Utilities.formatDate(new Date(oldDate.getTime() + 1000*60*60*6), 'GMT+2', 'MMMM dd, yyyy');
      var reFormatDate = new Date(e.range.getValue());
      e.value = Utilities.formatDate(new Date(reFormatDate.getTime() + 1000*60*60*6), 'GMT+2', "MMMM dd, yyyy");
    };
  };
  
  var backUpR = [user, editArch.standardName[index], e.oldValue, e.value, date];
  dbDRV[dbR][editArch.standardDB[index]] = e.value;
  cache.put('dbDRV', JSON.stringify(dbDRV), 21600);
  editBackUp(ID, backUpR, props);
  dbSheet.getRange((dbR+1), (editArch.standardDB[index] +1)).setValue(e.value);
  dbDummy.getRange((dbR+1), (editArch.standardDB[index] +1)).setValue(e.value);
  colorMe(sheet, [[e.range.getRow(), e.range.getColumn()]], '#bed690', props);
};

// prompt day confirmation for status edit
function activityStatusCheck(sheet, e, editAxioms, props) {
  
  var sheetDR = sheet.getDataRange();
  
  if (e.value == 'مفتوحة') {
    protectSheet(sheet, ['E7:F7']);
    var checkDate = new Date(sheet.getRange('E7').getValue());
    checkDate = Utilities.formatDate(new Date(checkDate.getTime() + 1000*60*60*6), 'GMT+2', "MMMM dd, yyyy");
    var rule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList([checkDate, 'إعتمد التاريخ', 'إلغاء'], true).build();
    sheet.getRange('E7').setDataValidation(rule).setBackground('#ffe599');
    editAxioms.phase = 'statusCheck';
    props.setProperty('editAxioms', JSON.stringify(editAxioms));
    e.range.setBackground('#ffe599');
    return;
  } else if (e.value == 'منتهية') {
    // must check that Performance is exactly 100%
    protectSheet(sheet, ['B7:C7']);
    var checkDate = new Date(sheet.getRange('B7').getValue());
    checkDate = Utilities.formatDate(new Date(checkDate.getTime() + 1000*60*60*6), 'GMT+2', "MMMM dd, yyyy");
    var rule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList([checkDate, 'إعتمد التاريخ', 'إلغاء'], true).build();
    sheet.getRange('B7').setDataValidation(rule).setBackground('#ffe599');
    editAxioms.phase = 'statusCheck';
    props.setProperty('editAxioms', JSON.stringify(editAxioms));
    e.range.setBackground('#ffe599');
    return;
  };
};

function updateStatus(sheet, e, index, editArch, editAxioms, activities, props) {
  var cache = CacheService.getScriptCache();
  if (cache.get('code1') == null) {cacheMe(1, cache)}
  var dbDRV = JSON.parse(cache.get('dbDRV'));
  
  var ID = editAxioms.ID;
  var dbR = editAxioms.dbR;
  
  var eRange = e.range
  if (eRange.getA1Notation() == 'E7') {
    var dbC = 4;
  } else if (eRange.getA1Notation() == 'B7') {
    var dbC = 5;
  };
  
  var rule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireDate().build();
  
  var checkDate = new Date(dbDRV[dbR][dbC])
  checkDate = Utilities.formatDate(new Date(checkDate.getTime() + 1000*60*60*6), 'GMT+2', "MMMM dd, yyyy");
  
  if (e.value == 'إلغاء') {
    eRange.setValue(checkDate).setNumberFormat("mmmm dd, yyyy").setBackground(null);
    sheet.getRange('H11').setValue(dbDRV[dbR][3]).setBackground(null);
    protectSheet(sheet, editArch.unprotected);
    eRange.setDataValidation(rule);
    editAxioms.phase = 'viewEdit';
    props.setProperty('editAxioms', JSON.stringify(editAxioms));
    colorMe(sheet, [[5,8]], '#bed690', props);
    return;
  };
  
  var user = Session.getActiveUser();
  var date = Utilities.formatDate(new Date(), 'GMT+2', 'dd-MM-yyyy HH:mm:ss');
  
  var dbSheet = SpreadsheetApp.openById(dbID).getSheetByName('database');
  var dbDummy = SpreadsheetApp.getActive().getSheetByName('database');
  
  if (e.value == 'إعتمد التاريخ') {
    eRange.setValue(checkDate).setNumberFormat("mmmm dd, yyyy");
    var newStatus = sheet.getRange('H11').getValue();
    var backUpR = [user, 'Status', dbDRV[dbR][3], newStatus, date];
    dbDRV[dbR][3] = newStatus;
    cache.put('dbDRV', JSON.stringify(dbDRV), 21600);
    editBackUp(ID, backUpR, props);
    dbSheet.getRange((dbR+1), 4).setValue(newStatus);
    dbDummy.getRange((dbR+1), 4).setValue(newStatus);
    editAxioms.phase = 'viewEdit';
    eRange.setValue(checkDate).setNumberFormat("mmmm dd, yyyy");
    if (newStatus == 'مفتوحة') {
      var rule2 = SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(['مفتوحة', 'منتهية'], true).build();
      sheet.getRange('H11').setDataValidation(rule2);
      Logger.log(editArch.unprotectedShort);
      var index = editArch.unprotectedShort.indexOf(eRange.getA1Notation());
      editArch.unprotectedShort.splice(index, 1);
      editArch.unprotected.splice(index, 1);
      Logger.log(editArch.unprotectedShort);
      eRange.setDataValidation(rule);
    } else {
      editArch.unprotected = ['H5'];
    };
    props.setProperty('editAxioms', JSON.stringify(editAxioms));
    props.setProperty('editArch', JSON.stringify(editArch));
    protectSheet(sheet, editArch.unprotected);
    colorMe(sheet, [[eRange.getRow(), eRange.getColumn()], [11, 8]], '#bed690', props);
    return;
  };
  
  
};