function fillInActivity(sheet, editAxioms, IDrow, activities, props) {
  
  var tableSheet = SpreadsheetApp.getActive().getSheetByName('tables');
  
  var columns = ['E:F', 'B:C', 'B:C', 'E:F', 'B:F', 'H', 'H', 'B:F', 'H', 'H', 'F', 'D', 'B'];
  var columnsShort = ['E', 'B', 'B', 'E', 'B', 'H', 'H', 'B', 'H', 'H', 'F', 'D', 'B'];
  var rows = [7, 7, 9, 9, 5, 9,  7, 11, 13, 11, 13, 13, 13];
  var DRc =  [4, 1, 1, 4, 1, 7,  7,  1,  7,  7,  5,  3,  1];
  var DBc =  [4, 5, 7, 6, 1, 2, 14, 15, 13,  3, 10, 11, 12];
  var fieldNames = ['Start', 'End', 'Baseline End', 'Baseline Start', 'Activity Name', 'Primavera ID', 'Duration', 'Notes', 'Activity Executor', 'Status', 'Labor Cost', 'Material Cost', 'Equipment Cost'];
  
  var editArch = {
    'unprotected': [],
    'unprotectedShort': [],
    'standard': [],
    'standardDB': [],
    'standardName': [],
    'DEP': [],
    'KPI': [],
    'subContractor': [],
    'rules': []
  };
  
  var cache = CacheService.getScriptCache();
  if (cache.get('code1') == null) {cacheMe(1, cache)}
  var dbDRV = JSON.parse(cache.get('dbDRV'));
  
  var subContractor = false;
  if (dbDRV[IDrow][13] != 'N/A') {
    subContractor = true;
  };
    
  var newDEP = JSON.parse(dbDRV[IDrow][9]);
  var newKPI = JSON.parse(dbDRV[IDrow][8]);
  
  var newAxioms = {
    'phase': 'viewEdit',
    'ID': dbDRV[IDrow][0],
    'dbR': IDrow,
    'status': dbDRV[IDrow][3],
    'subContractor': subContractor,
    'tableDEP': [17, newDEP.length],
    'tableKPI': [20, newKPI.length]
  };
  
  if (newAxioms.status == 'منتهية') {newAxioms.phase = 'view'};
  
  if (editAxioms.ID.split('-')[1] != 0 && newAxioms.ID.split('-')[1] == 0) {
    sheet.getRange('B5:F5').merge();
  };
  
  if (newAxioms.ID.split('-')[1] != 0) {
    columns[4] = 'B:D'
    if (editAxioms.ID.split('-')[1] == 0 || editAxioms.ID.split('-')[1] == 'x') {
      tableSheet.getRange('B5:F5').copyTo(sheet.getRange('B5:F5'));
    };
  };
  
  if (editAxioms.subContractor == false) {
    if (newAxioms.subContractor == true) {
      sheet.insertRows(12, 2);
      tableSheet.getRange('B7:I7').copyTo(sheet.getRange('B13:I13'));
      var l = editArch.rows.length;
      for (i=0; i<l; i++) {
        if (rows[i] > 12) {
          rows[i] += 2
        };
      };
      newAxioms.tableDEP[0] += 2;
      newAxioms.tableKPI[0] += 2;
    };
  } else if (newAxioms.subContractor == false) {
    sheet.deleteRows(13, 2);
    var l = rows.length;
    for (i=0; i<l; i++) {
      if (rows[i] > 12) {
        rows[i] -= 2
      };
    };
  } else {
    newAxioms.tableDEP[0] += 2;
    newAxioms.tableKPI[0] += 2;
  };
  
  var parent = false;
  if (newAxioms.ID.split('-')[1] == 0) {
    var index = activities.parentActIDs.indexOf(newAxioms.ID);
    if (activities.prettyActIDs[index].length > 1) {
      parent = true;
    };
  };
  
  var oldParent = false;
  if (editAxioms.ID.split('-')[1] == 0) {
    var index = activities.parentActIDs.indexOf(editAxioms.ID);
    if (activities.prettyActIDs[index].length > 1) {
      oldParent = true;
    };
  };
  
  if (newAxioms.status != 'منتهية') {
    newAxioms.tableDEP[1]++;
    if (parent == false) {newAxioms.tableKPI[1]++};
  };
  
  if (newAxioms.tableDEP[1] > editAxioms.tableDEP[1]) {
    sheet.insertRows((newAxioms.tableDEP[0] + editAxioms.tableDEP[1]), (newAxioms.tableDEP[1] - editAxioms.tableDEP[1]));
    newAxioms.tableKPI[0] += newAxioms.tableDEP[1];
  } else if (newAxioms.tableDEP[1] < editAxioms.tableDEP[1]) {
    sheet.deleteRows(newAxioms.tableDEP[0], (editAxioms.tableDEP[1] - newAxioms.tableDEP[1]))
    newAxioms.tableKPI[0] += newAxioms.tableDEP[1];
  } else if (newAxioms.status != 'منتهية') {
    newAxioms.tableKPI[0]++
  };
  
  if (newAxioms.tableKPI[1] > editAxioms.tableKPI[1]) {
    sheet.insertRows((newAxioms.tableKPI[0] + editAxioms.tableKPI[1]), (newAxioms.tableKPI[1] - editAxioms.tableKPI[1]));
  } else if (newAxioms.tableKPI[1] < editAxioms.tableKPI[1]) {
    sheet.deleteRows(newAxioms.tableKPI[0], (editAxioms.tableKPI[1] - newAxioms.tableKPI[1]))
  };
  
    
  if (oldParent == true && parent == false) {
    tableSheet.getRange('B9:I10').copyTo(sheet.getRange('B'+(newAxioms.tableKPI[0]-2)+':I'+(newAxioms.tableKPI[0]-1)));
  } else if (oldParent == false && parent == true) {
    tableSheet.getRange('B13:I14').copyTo(sheet.getRange('B'+(newAxioms.tableKPI[0]-2)+':I'+(newAxioms.tableKPI[0]-1)));
  };
  
  var sheetDR = sheet.getDataRange();
  var sheetDRV = sheetDR.getValues();
  
  sheetDR.setDataValidation(null);
  var rules = sheetDR.getDataValidations();
  
  var rule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(activities.actIDs, false).build();
  var rule2 = SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(['SS', 'SF', 'FS', 'FF'], true).build();
  var rule3 = SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireNumberGreaterThanOrEqualTo(-1000).build();
  var rule4 = SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireDate().build();
  
  editArch.unprotected.push('H5');
  editArch.unprotectedShort.push('H5');
  rules[4][7] = rule;

  var editA1Not = [];
  var editA1NotShort = [];
  var l = columns.length;
  for (i=0; i<l; i++) {
    editA1NotShort.push(columnsShort[i]+rows[i]);
    var brk = columns[i].split('');
    var index2 = brk.indexOf(':');
    if (index2 != -1) {
      brk.splice(index2, 0, rows[i]);
      brk.push(rows[i]);
      editA1Not.push(brk.join(''));
    } else {
      editA1Not.push(columnsShort[i]+rows[i]);
    };
  };
  
  if (newAxioms.ID.split('-')[1] != 0) {
    var parentIndex = activities.actIDs.indexOf((newAxioms.ID.split('-')[0] + '-' + 0));
    sheetDRV[4][4] = dbDRV[parentIndex + 1][1]
  };
  
  for (i=0; i<l; i++) {
    if (i<4) {
      var temp = new Date(dbDRV[IDrow][DBc[i]])
      sheetDRV[rows[i]-1][DRc[i]] = Utilities.formatDate(new Date(temp.getTime() + 1000*60*60*6), 'GMT+2', 'MMMM dd, yyyy');
      rules[rows[i]-1][DRc[i]] = rule4;
    } else {
      sheetDRV[rows[i]-1][DRc[i]] = dbDRV[IDrow][DBc[i]];
    };
    if (newAxioms.status != 'منتهية' && i<10) {
      if ((newAxioms.status == 'مفتوحة' && i==0) || (parent == true && i==9)) {continue};
      editArch.unprotected.push(editA1Not[i]);
      editArch.unprotectedShort.push(editA1NotShort[i]);
      if (i<8) {
        editArch.standard.push(editA1NotShort[i]);
        editArch.standardDB.push(DBc[i]);
        editArch.standardName.push(fieldNames[i]);
      };
      if (i==6) {
        rules[rows[i]-1][DRc[i]] = rule3;
      };
      if (i==8) {
        rules[rows[i]-1][DRc[i]] = SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(['الشركة', 'مقاول فرعي'], true).build();
      };
      if (i==9) {
        if (newAxioms.status == 'مفتوحة') {
          rules[rows[i]-1][DRc[i]] = SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(['مفتوحة', 'منتهية'], true).build();
        } else {
          rules[rows[i]-1][DRc[i]] = SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(['مغلقة', 'مفتوحة'], true).build();
        };
      };
    };
  };
  
  if (subContractor == false) {
    sheetDRV[rows[8]-1][DRc[8]] = 'الشركة';
  } else {
    var newSubContract = JSON.parse(dbDRV[IDrow][DBc]);
    sheetDRV[rows[8]-1][DRc[8]] = 'مقاول فرعي';
    sheetDRV[14][5] = newSubContract[0];
    sheetDRV[14][3] = newSubContract[1];
    // Figure out percentage completion from KPI then Multiply it by newSubContract[1] to get the value of sheetDRV[14][1]
    editArch.unprotected.push('F15:H15', 'D15');
    editArch.unprotectedShort.push('F15', 'D15');
    editArch.subContractor.push('F15:H15', 'D15');
    rules[14][3] = rule3;
  };
  
  if (newAxioms.tableDEP[1] > 0) {
    var DEProw = newAxioms.tableDEP[0] - 1;
    var DEPendR = newAxioms.tableDEP[0] + newAxioms.tableDEP[1] - 1;
    tableSheet.getRange('B19:I19').copyTo(sheet.getRange('B'+newAxioms.tableDEP[0]+':I'+DEPendR));
    for (i=0; i<newAxioms.tableDEP[1]; i++) {
      sheetDRV[DEProw + i][5] = '=IFERROR(VLOOKUP(I' + (DEProw + i + 1) + ',database!$A$2:$H$'+ (activities.actIDs.length + 1) + ',2,false),"")';
      sheetDRV[DEProw + i][4] = '=IFERROR(VLOOKUP(I' + (DEProw + i + 1) + ',database!$A$2:$H$'+ (activities.actIDs.length + 1) + ',5,false),"")';
      sheetDRV[DEProw + i][3] = '=IFERROR(VLOOKUP(I' + (DEProw + i + 1) + ',database!$A$2:$H$'+ (activities.actIDs.length + 1) + ',6,false),"")';
      if(newAxioms.status != 'منتهية') {
        editArch.DEP.push(('B'+ (DEProw + i + 1)), ('C'+ (DEProw + i + 1)), ('I'+ (DEProw + i + 1)));
        editArch.unprotectedShort.push(('B'+ (DEProw + i + 1)), ('C'+ (DEProw + i + 1)), ('I'+ (DEProw + i + 1)));
        rules[DEProw + i][1] = rule3;
        rules[DEProw + i][2] = rule2;
        rules[DEProw + i][8] = rule;
        if(i == (newAxioms.tableDEP[1]-1)) {
          sheetDRV[DEProw + i][1] = '';
          sheetDRV[DEProw + i][2] = '';
          sheetDRV[DEProw + i][8] = '';
          break;
        };
      };
      sheetDRV[DEProw + i][1] = newDEP[i][2];
      sheetDRV[DEProw + i][2] = newDEP[i][1];
      sheetDRV[DEProw + i][8] = newDEP[i][0];
    };
    if(newAxioms.status != 'منتهية') {
      editArch.unprotected.push('B' + newAxioms.tableDEP[0] + ':C' + DEPendR);
      editArch.unprotected.push('I' + newAxioms.tableDEP[0] + ':I' + DEPendR);
    };
  };
  
  if (newAxioms.tableKPI[1] > 0) {
    var KPIrow = newAxioms.tableKPI[0] - 1;
    var KPIendR = newAxioms.tableKPI[0] + newAxioms.tableKPI[1] - 1;
    if (parent == true) {
      tableSheet.getRange('B15:I15').copyTo(sheet.getRange('B'+newAxioms.tableKPI[0]+':I'+KPIendR));
      for (i=0; i<newAxioms.tableKPI[1]; i++) {
        sheetDRV[KPIrow + i][1] = newKPI[i][1];
        sheetDRV[KPIrow + i][8] = newKPI[i][0];
        sheetDRV[KPIrow + i][2] = '=IFERROR(VLOOKUP(I' + (KPIrow + i + 1) + ',database!$A$2:$H$'+ (activities.actIDs.length + 1) + ',4,false),"")';
        sheetDRV[KPIrow + i][5] = '=IFERROR(VLOOKUP(I' + (KPIrow + i + 1) + ',database!$A$2:$H$'+ (activities.actIDs.length + 1) + ',2,false),"")';
        sheetDRV[KPIrow + i][4] = '=IFERROR(VLOOKUP(I' + (KPIrow + i + 1) + ',database!$A$2:$H$'+ (activities.actIDs.length + 1) + ',5,false),"")';
        sheetDRV[KPIrow + i][3] = '=IFERROR(VLOOKUP(I' + (KPIrow + i + 1) + ',database!$A$2:$H$'+ (activities.actIDs.length + 1) + ',6,false),"")';
        if(newAxioms.status != 'منتهية') {
          editArch.KPI.push('B' + (KPIrow + i + 1));
          editArch.unprotectedShort.push('B' + (KPIrow + i + 1));
          rules[KPIrow + i][1] = rule3;
        };
      };
      if(newAxioms.status != 'منتهية') {
        editArch.unprotected.push('B' + newAxioms.tableKPI[0] + ':B' + KPIendR);
        
      };
    } else {
      tableSheet.getRange('B11:I11').copyTo(sheet.getRange('B'+newAxioms.tableKPI[0]+':I'+KPIendR));
      for (i=0; i<newAxioms.tableKPI[1]; i++) {
        // Remove edit capabaility of إنجاز such that it can only be edited from "Activity Performance Management" - or atleast not allow change when activity is closed.
        if(newAxioms.status != 'منتهية') {
          editArch.KPI.push(('B' + (KPIrow + i + 1)), ('C' + (KPIrow + i + 1)), ('D' + (KPIrow + i + 1)), ('E' + (KPIrow + i + 1)), ('F' + (KPIrow + i + 1)), ('G' + (KPIrow + i + 1)));
          editArch.unprotectedShort.push(('B' + (KPIrow + i + 1)), ('C' + (KPIrow + i + 1)), ('D' + (KPIrow + i + 1)), ('E' + (KPIrow + i + 1)), ('F' + (KPIrow + i + 1)), ('G' + (KPIrow + i + 1)));
          rules[KPIrow + i][1] = rule3;
          rules[KPIrow + i][2] = rule3;
          rules[KPIrow + i][3] = rule3;
          if(i == (newAxioms.tableKPI[1]-1)) {
            sheetDRV[KPIrow + i][1] = '';
            sheetDRV[KPIrow + i][2] = '';
            sheetDRV[KPIrow + i][3] = '';
            sheetDRV[KPIrow + i][4] = '';
            sheetDRV[KPIrow + i][5] = '';
            sheetDRV[KPIrow + i][8] = '';
            break;
          };
        };
        sheetDRV[KPIrow + i][1] = newKPI[i][5];
        sheetDRV[KPIrow + i][2] = newKPI[i][4];
        sheetDRV[KPIrow + i][3] = newKPI[i][3];
        sheetDRV[KPIrow + i][4] = newKPI[i][2];
        sheetDRV[KPIrow + i][5] = newKPI[i][1];
        sheetDRV[KPIrow + i][6] = newKPI[i][0];
      };
      if(newAxioms.status != 'منتهية') {
        editArch.unprotected.push('B' + newAxioms.tableKPI[0] + ':I' + KPIendR);
      };
    };
  };
  
  props.setProperty('editAxioms', JSON.stringify(newAxioms));
  props.setProperty('editArch', JSON.stringify(editArch));
  
  protectSheet(sheet, editArch.unprotected);
  sheetDR.setValues(sheetDRV).setDataValidations(rules);
  
  return;
};

function retrunEditForm(sheet, editAxioms, activities, props) {
  
  var rows = [7, 7, 9, 9, 5, 9,  7, 11, 13, 11, 13, 13, 13];
  var DRc =  [4, 1, 1, 4, 1, 7,  7,  1,  7,  7,  5,  3,  1];
  
  if (editAxioms.ID.split('-')[1] != 0) {
      sheet.getRange('B5:F5').merge();
  };
  
  if (editAxioms.subContractor == true) {
    sheet.deleteRows(13, 2);
  };
  
  if (editAxioms.tableDEP[1] > 0) {
    sheet.deleteRows(17, editAxioms.tableDEP[1]);
  };
  
  if (editAxioms.tableKPI[1] > 0) {
    sheet.deleteRows(20, editAxioms.tableKPI[1]);
  };
  
  var sheetDR = sheet.getDataRange();
  var sheetDRV = sheetDR.getValues();
  
  var l = rows.length;
  for (i=0; i<l; i++) {
    sheetDRV[rows[i]-1][DRc[i]] = '';
  };
  
  sheetDR.setDataValidation(null);
  
  var oldParent = false;
  if (editAxioms.ID.split('-')[1] == 0) {
    var index = activities.parentActIDs.indexOf(editAxioms.ID);
    if (activities.prettyActIDs[index].length > 1) {
      oldParent = true;
    };
  };
  
  var editAxioms = {
    'phase': 'open',
    'ID': 'nada-x',
    'dbR': 'nada',
    'status': 'منتهية',
    'subContractor': false,
    'tableDEP': [17, 0],
    'tableKPI': [20, 0]
  };
  
  if (oldParent == true) {
    SpreadsheetApp.getActive().getSheetByName('tables').getRange('B10:I10').copyTo(sheet.getRange('B'+(editAxioms.tableKPI[0]-1)+':I'+(editAxioms.tableKPI[0]-1)));
    sheetDRV[editAxioms.tableKPI[0] - 3][1] = 'مقاييس الإنجاز ووزنها';
  };
  
  var editArch = {
    'unprotected': ['H5'],
    'unprotectedShort': ['H5'],
    'standard': [],
    'standardDB': [],
    'DEP': [],
    'KPI': [],
    'subContractor': [],
    'rules': []
  };
  
  var rule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(activities.actIDs, false).build();
  
  sheet.getRange('H5').setDataValidation(rule);
  protectSheet(sheet, editArch.unprotected);
  
  props.setProperty('editAxioms', JSON.stringify(editAxioms));
  props.setProperty('editArch', JSON.stringify(editArch));
  
  sheetDR.setValues(sheetDRV);

};