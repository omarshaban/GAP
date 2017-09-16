function submitActivity(sheet, createArch, props) {
  
  var sheetDR = sheet.getDataRange();
  var sheetDRV = sheetDR.getValues();
  
  var phase = createArch.phase;
  var date = Utilities.formatDate(new Date(), 'GMT+2', 'dd-MM-yyyy HH:mm:ss');
  var user = Session.getActiveUser();
  var newR = new Array(16);
  var backUpGroup = [];
  var error = [];
  
  var activities = JSON.parse(props.getProperty('activities'));
  var actIDs = activities.actIDs;
  var prettyActIDs = activities.prettyActIDs;
  var parentActIDs = activities.parentActIDs;
  
  
  if (phase == 'createParent') {
    var l = parentActIDs.length;
    var maxInc = -1;
    for (i=0; i<l; i++) {
      if (parseInt(parentActIDs[i].split('-')[0]) > maxInc) {
        maxInc = parentActIDs[i].split('-')[0];
      };
    };
    var newID = ++maxInc + '-' + 0;
  } else if (phase == 'createChild' ) {
    var parentAct = sheetDRV[6][7];
    var index = parentActIDs.indexOf(parentAct);
    if (index == -1) {
      error.push([7, 8]);
    } else {
      var maxInc = -1;
      var l = prettyActIDs[index].length;
      for (i=0; i<l; i++) {
        if (parseInt(prettyActIDs[index][i].split('-')[1]) > maxInc) {
          maxInc = prettyActIDs[index][i].split('-')[1];
        };
      };
      var newID = parentAct.split('-')[0] + '-' + (++maxInc);
      var insertIndex = actIDs.indexOf(parentAct.split('-')[0] + '-' + (maxInc-1));
    };
  } else {
    // Mayday Mayday
    hello();
  };
  
  newR[0] = newID;
  newR[3] = 'مغلقة';
  newR[10] = 0;
  newR[11] = 0;
  newR[12] = 0;
  newR[13] = 'N/A';
  
  if (sheetDRV[createArch.createR[1] - 1][createArch.createDRc[1]] != '') {
    backUpGroup.push([user, createArch.createFieldNames[1], '-', sheetDRV[createArch.createR[1] - 1][createArch.createDRc[1]], date])
    newR[createArch.createDBc[1] - 1] = sheetDRV[createArch.createR[1] - 1][createArch.createDRc[1]]
  } else {
    error.push([createArch.createR[1], (createArch.createDRc[1] + 1)])
  };
  
  for (i=2; i<6; i++) {
    var dateRange = sheet.getRange(createArch.createR[i],(createArch.createDRc[i] + 1));
    if (dateRange.getNumberFormat() != "mmmm dd, yyyy" || sheetDRV[createArch.createR[i] - 1][createArch.createDRc[i]] == '') {
      error.push([createArch.createR[i], (createArch.createDRc[i] + 1)]);
      dateRange.clearContent().setNumberFormat("mmmm dd, yyyy")
    } else {
      var reFormatDate = new Date(sheetDRV[createArch.createR[i] - 1][createArch.createDRc[i]])
      reFormatDate = Utilities.formatDate(new Date(reFormatDate.getTime() + 1000*60*60*6), 'GMT+2', "MMMM dd, yyyy");
      backUpGroup.push([user, createArch.createFieldNames[i], '-', reFormatDate, date]);
      newR[createArch.createDBc[i] - 1] = reFormatDate;
    };
  };
  
  l = createArch.createR.length;
  for (i=6; i<l; i++) {
    if (sheetDRV[createArch.createR[i] - 1][createArch.createDRc[i]] != '') {
      backUpGroup.push([user, createArch.createFieldNames[i], '-', sheetDRV[createArch.createR[i] - 1][createArch.createDRc[i]], date])
      newR[createArch.createDBc[i] - 1] = sheetDRV[createArch.createR[i] - 1][createArch.createDRc[i]]
    } else {
      error.push([createArch.createR[i], (createArch.createDRc[i] + 1)])
    };
  };
  
  l = createArch.createTables[0][1];
  if (l > 1) {
    var KPIgroup = [];
    var KPIweight = [0, []];
    var KPIrow = createArch.createTables[0][0] -1;
    var KPInames = ['رقم المقياس', 'المقياس', 'الوحدة', 'الكمية', 'الإنجاز', 'الوزن'];
    for (i=0; i<l-1; i++) {
      var KPIsource = [
        sheetDRV[KPIrow + i][(createArch.createTables[0][2] + 2)],
        sheetDRV[KPIrow + i][(createArch.createTables[0][2] + 1)], 
        sheetDRV[KPIrow + i][createArch.createTables[0][2]], 
        sheetDRV[KPIrow + i][createArch.createTables[0][2]], 
        0, 
        sheetDRV[KPIrow + i][(createArch.createTables[0][2] - 1)]
      ];
      KPIgroup[i] = [];
      for (j=0; j<6; j++) {
        KPIgroup[i].push(KPIsource[j])
        backUpGroup.push([user, KPInames[j], '-', KPIsource[j], date])
      };
      KPIweight[0] += sheetDRV[KPIrow + i][(createArch.createTables[0][2] - 1)];
      KPIweight[1].push([(KPIrow + i + 1), (createArch.createTables[0][2] )])
    };
    newR[createArch.createTables[0][4] - 1] = JSON.stringify(KPIgroup);
    if (KPIweight[0] != 100) {
      error = error.concat(KPIweight[1]);
    };
  } else {
    newR[createArch.createTables[0][4] - 1] = JSON.stringify([]);
  };
  
  l = createArch.createTables[1][1];
  if (l > 1) {
    var DEPgroup = [];
    var DEProw = createArch.createTables[1][0] -1;
    var DEPnames = ['رقم الفعالية', 'نوع الربط', 'lead/lag'];
    for (i=0; i<l-1; i++) {
      var DEPsource = [
        sheetDRV[DEProw + i][(createArch.createTables[0][2] + 5)], 
        sheetDRV[DEProw + i][createArch.createTables[0][2] - 1],
        sheetDRV[DEProw + i][(createArch.createTables[0][2] - 2)]
      ];
      DEPgroup[i] = [];
      for (j=0; j<3; j++) {
        DEPgroup[i].push(DEPsource[j])
        backUpGroup.push([user, DEPnames[j], '-', DEPsource[j], date])
      };
    };
    newR[createArch.createTables[1][4] - 1] = JSON.stringify(DEPgroup);
  } else {
    newR[createArch.createTables[1][4] - 1] = JSON.stringify([]);
  };
  
  
  if (error == '') {
    var ss = SpreadsheetApp.getActive();
    var dbDummy = ss.getSheetByName('database');
    var backUpSheet = SpreadsheetApp.openById(dbID).getSheetByName('database');
    var cache = CacheService.getScriptCache();
    if (cache.get('code1') == null) {cacheMe(1, cache)}
    var dbDRV = JSON.parse(cache.get('dbDRV'));
    
    if (phase == 'createChild' ) {
      
      var parentRow = actIDs.indexOf(parentAct) + 2;
      if (prettyActIDs[index].length == 1) {
        var parentPerformance = JSON.stringify([[newID, 100]]);
      } else {
        var parentPerformance = JSON.parse(dbDummy.getRange(parentRow, 9).getValue());
        parentPerformance = JSON.stringify(parentPerformance.concat([[newID, 0]]));
      };
      
      dbDummy.getRange(parentRow, 9).setValue(parentPerformance);
      backUpSheet.getRange(parentRow, 9).setValue(parentPerformance);
      index = parentActIDs.indexOf(parentAct);
      activities.actIDs.splice((insertIndex + 1), 0, newID);
      activities.prettyActIDs[index].push(newID);
      dbDRV.splice((insertIndex + 2), 0, newR);
      cache.put('dbDRV', JSON.stringify(dbDRV), 21600);
      var rule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(activities.actIDs, false).build();
      ss.getSheetByName('عرض').getRange('H5').setDataValidation(rule);
      updateBU(backUpGroup, newID, props);
      dbDummy.insertRowAfter(insertIndex + 2);
      dbDummy.getRange((insertIndex + 3), 1, 1, 16).setValues([newR]);
      backUpSheet.insertRowAfter(insertIndex + 2);
      backUpSheet.getRange((insertIndex + 3), 1, 1, 16).setValues([newR]);
      props.setProperty('activities', JSON.stringify(activities));
      return 'success';
      
    } else if (phase == 'createParent' ) {
      activities.parentActIDs.push(newID);
      activities.prettyActIDs.push([newID]);
      activities.actIDs.push(newID);
      dbDRV.push(newR);
      cache.put('dbDRV', JSON.stringify(dbDRV), 21600);
      var rule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(activities.actIDs, false).build();
      ss.getSheetByName('عرض').getRange('H5').setDataValidation(rule);
      updateBU(backUpGroup, newID, props);
      backUpSheet.appendRow(newR);
      dbDummy.appendRow(newR);
      props.setProperty('activities', JSON.stringify(activities));
      return 'success';
      
    };
  } else {
    colorMe(sheet, error, '#e38888', props);
    return
  };
};

function setCreateForm(sheet, createArch, props) {
  
  var activities = JSON.parse(props.getProperty('activities'));
  var actIDs = activities.actIDs;
  
  var rule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(['تم', 'إلغاء'], true).build();
  sheet.getRange('B3').setDataValidation(rule);
  
  var l = createArch.createTables[1][1];
  var row = createArch.createTables[1][0];
  rule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(actIDs, false).build();
  var rule2 = SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(['SS', 'SF', 'FS', 'FF'], true).build();
  for (i=0; i<l; i++) {
    sheet.getRange((row + i), 3).setDataValidation(rule2);
    sheet.getRange((row + i), 9).setDataValidation(rule);
    sheet.getRange((row + i), 4, 1, 3).setValues([['=IFERROR(VLOOKUP(I' + (row + i) + ',database!$A$2:$H$'+ (actIDs.length + 1) + ',6,false),"")', '=IFERROR(VLOOKUP(I' + (row + i) + ',database!$A$2:$H$'+ (actIDs.length + 1) + ',5,false),"")', '=IFERROR(VLOOKUP(I' + (row + i) + ',database!$A$2:$H$'+ (actIDs.length + 1) + ',2,false),"")']])
  };
  
  
  if (createArch.phase == 'createChild') {
    sheet.getRange('B5:I5').copyTo(sheet.getRange('B7:I7'));
    sheet.getRange('G7:I7').setValues([['الفعالية الرئيسية','', 'رقم الفعالية ']]);
    
    var parentActIDs = activities.parentActIDs;
    var rule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(parentActIDs, false).build();
    sheet.getRange('H7').setDataValidation(rule);
    sheet.getRange('B7').setValue('=IFERROR(VLOOKUP(H7,database!$A$2:$C$'+ (actIDs.length + 1) + ',2,false),"")')
  };
    
};

function returnCreateForm(sheet, createArch, props) {
  
  var sheetDR = sheet.getDataRange();
  var sheetDRV = sheetDR.getValues();
  
  var l = createArch.createR.length;
  for (i=0; i<l; i++) {
    sheetDRV[(createArch.createR[i] - 1)][createArch.createDRc[i]] = '';
  };
  
  for (k=0; k<createArch.createTables.length; k++) {
    var l = createArch.createTables[k][1];
    var row = createArch.createTables[k][0] - 1;
    for (i=0; i<l; i++) {
      for (j=(createArch.createTables[k][2] - 1); j < (createArch.createTables[k][2] - 1 + createArch.createTables[k][3]); j++) {
        sheetDRV[row + i][j] = '';
      };
    };
  };
  sheetDRV[2][1] = '';
  
  sheetDR.setDataValidation(null);
  sheetDR.setValues(sheetDRV);
  
  if (createArch.createTables[0][1] > 1) {
    sheet.deleteRows(16, (createArch.createTables[0][1] - 1));
    createArch.createTables[0][1] = 1;
  };
  
  if (createArch.createTables[1][1] > 1) {
    sheet.deleteRows(20, (createArch.createTables[1][1] - 1))
    createArch.createTables[1][1] = 1
    createArch.createTables[2][1] = 1
  };
  
  sheet.getRange(createArch.createTables[1][0], 4,1,5).clearContent();
  
  createArch.phase = 'open';
  props.setProperty('createArch', JSON.stringify(createArch));
  
  var rule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(['رئيسية', 'فرعية'], true).build();
  sheet.getRange('H5').setDataValidation(rule);
  
  protectSheet(sheet, ['H5']);
  
};

function editCreateTableKPI(sheet, eRow, eRange, eValue, createArch, props) {
  if (eValue == undefined) {
    if(eRow < (createArch.createTables[0][0] + createArch.createTables[0][1] -1) && eRange.getColumn() == 6 && createArch.createTables[0][1] > 1) {
      createArch.createTables[0][1] -= 1;
      createArch.createTables[1][0] -= 1;
      createArch.createTables[2][0] -= 1;
      props.setProperty('createArch', JSON.stringify(createArch));
      sheet.deleteRow(eRow);
      return;
    } else {
      return;
    };
  } else if (eRow == (createArch.createTables[0][0] + createArch.createTables[0][1] -1)) {
    if (sheet.getRange(eRow, createArch.createTables[0][2], 1, createArch.createTables[0][3]).getValues()[0].filter(String).length == (createArch.createTables[0][3]-2)) {
      createArch.createTables[0][1] += 1;
      createArch.createTables[1][0] += 1;
      createArch.createTables[2][0] += 1;
      props.setProperty('createArch', JSON.stringify(createArch));
      sheet.insertRowAfter(createArch.createTables[0][0] + createArch.createTables[0][1] - 2);
      sheet.getRange(createArch.createTables[0][0], createArch.createTables[0][2], 1, createArch.createTables[0][3])
      .copyFormatToRange(sheet, createArch.createTables[0][2], (createArch.createTables[0][2] + createArch.createTables[0][3]), (createArch.createTables[0][0] + createArch.createTables[0][1] - 1), (createArch.createTables[0][0] + createArch.createTables[0][1] - 1));
      return;
    };
  };
};

function editCreateTableDEP(sheet, eRow, eRange, eValue, createArch, props) {
  if (eValue == undefined) {
    if(eRow < (createArch.createTables[1][0] + createArch.createTables[1][1] -1) && eRange.getColumn() == 9 && createArch.createTables[1][1] > 1) {
      createArch.createTables[1][1] -= 1;
      createArch.createTables[2][1] -= 1;
      props.setProperty('createArch', JSON.stringify(createArch));
      sheet.deleteRow(eRow);
      return;
    } else {
      return;
    };
  } else if (eRow == (createArch.createTables[1][0] + createArch.createTables[1][1] -1)) {
    if (sheet.getRange(eRow, 2, 1, 8).getValues()[0].filter(String).length == 6) {
      var activities = JSON.parse(props.getProperty('activities'));
      var actIDs = activities.actIDs;
      createArch.createTables[1][1] += 1;
      createArch.createTables[2][1] += 1;
      props.setProperty('createArch', JSON.stringify(createArch));
      sheet.insertRowAfter(createArch.createTables[1][0] + createArch.createTables[1][1] - 2);
      var newR = (createArch.createTables[1][0] + createArch.createTables[1][1] - 1);
      sheet.getRange(createArch.createTables[1][0], 2, 1, 8).copyFormatToRange(sheet, 2, 9, newR, newR)
      sheet.getRange(newR, 4, 1, 3).setValues([['=IFERROR(VLOOKUP(I' + newR + ',database!$A$2:$H$'+ (actIDs.length + 1) + ',6,false),"")', '=IFERROR(VLOOKUP(I' + newR + ',database!$A$2:$H$'+ (actIDs.length + 1) + ',5,false),"")', '=IFERROR(VLOOKUP(I' + newR + ',database!$A$2:$H$'+ (actIDs.length + 1) + ',2,false),"")']])
      return;        
    };
  };
};

function colorMe(sheet, ranges, color, props) {
  var l = ranges.length;
  for (i=0; i<l; i++) {
    sheet.getRange(ranges[i][0], ranges[i][1]).setBackgroundColor(color);
  };
  props.setProperty('colorR', JSON.stringify(ranges))
};

function colorMeNot(sheet, ranges) {
  var l = ranges.length;
  for (i=0; i<l; i++) {
    sheet.getRange(ranges[i][0], ranges[i][1]).setBackgroundColor(null);
  };
};