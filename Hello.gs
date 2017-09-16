/*
DEP clashes - Report color coding
Earliest Start according to DEP (Color code bottle neck)
clashes between start end and duration
edit color me to return old color not null
Ready architecture of all actIDs and set it to a cache that is independent onOpen (runs every hour?)


*/


var dbID = '1zgIUsHLpP0rOhBaooddL9-EMguiA24_abD9M2y6hP2w';

function hello () {
  
  var ss = SpreadsheetApp.getActive();
  var editForm = ss.getSheetByName('editForm');
  var createForm = ss.getSheetByName('createForm');
  var dbSheet = SpreadsheetApp.openById(dbID).getSheetByName('database');
  var props = PropertiesService.getScriptProperties();
  var cache = CacheService.getScriptCache();
  
  props.deleteAllProperties();
  
  createForm.showSheet();
  editForm.showSheet();
  
  var sheets = ss.getSheets();
  var l = sheets.length;
  for (i=0; i<l; i++) {
    if (sheets[i].getName() != editForm.getName() && sheets[i].getName() != createForm.getName() && sheets[i].getName() != 'tables') {
      ss.deleteSheet(sheets[i]);
    };
  };
  
  var dbDummy = dbSheet.copyTo(ss).setName('database');
  var editSheet = editForm.copyTo(ss).setName('عرض');
  var createSheet = createForm.copyTo(ss).setName('إضافة');
  
  protectSheet(dbDummy, null);
  protectSheet(editSheet, ['H5']);
  protectSheet(createSheet, ['H5']);
  
  var createC = ['H', 'B:F', 'B:C', 'E:F', 'B:C', 'E:F', 'H', 'B:F', 'H'];
  var createCshort = ['H', 'B', 'B', 'E', 'B', 'E', 'H', 'B', 'H'];
  var createR =   [5, 5, 7, 7, 9, 9, 9, 11, 7];
  var createDRc = [7, 1, 1, 4, 1, 4, 7,  1, 7];
  // createTables is [[first Row in table, Number of Rows, first Column in table, number of Columns, Column number in database(A=1)]]
  var createTables = [[15, 1, 3, 6, 9],[19, 1, 2, 2, 10],[19, 1, 9, 1, 10]];
  var createDBc = ['x', 2, 6, 5, 8, 7, 3, 16, 15];
  var createFieldNames = ['x', 'Activity Name', 'End', 'Start', 'Baseline End', 'Baseline Start', 'Primavera ID', 'Notes', 'Duration']
  
  var createArch = {
    'phase': 'open',
    'createC': createC,
    'createCshort': createCshort,
    'createR': createR,
    'createDRc': createDRc,
    'createTables': createTables,
    'createDBc': createDBc,
    'createFieldNames': createFieldNames
  };
  
  var dbDRV = dbDummy.getDataRange().getValues();
  var l = dbDRV.length;
  var actIDs = [];
  var parentActIDs = [];
  var prettyActIDs = [];
  for (i=1; i<l; i++) {
    actIDs.push(dbDRV[i][0]);
  };
  
  var sortActIDs = actIDs.slice(0).sort(function (a, b) {
    return a.split('-')[1] - b.split('-')[1];
  });
    sortActIDs.sort(function (a, b) {
    return a.split('-')[0] - b.split('-')[0];
  });
  
  var j = 0;
  prettyActIDs[j] = [sortActIDs[0]];
  parentActIDs = [sortActIDs[0]];
  for (i=1; i<l-1; i++) {
    if (prettyActIDs[j][0].split('-')[0] == sortActIDs[i].split('-')[0]) {
      prettyActIDs[j].push(sortActIDs[i]);
    } else {
      prettyActIDs[++j] = [sortActIDs[i]];
      parentActIDs.push(sortActIDs[i]);
    };
  };
  
  var rule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(actIDs, false).build();
  editSheet.getRange('H5').setDataValidation(rule);
  
  var activities = {
    'actIDs': actIDs,
    'parentActIDs': parentActIDs,
    'prettyActIDs': prettyActIDs
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
  
  BUmap(props);
  
  props.setProperty('activities', JSON.stringify(activities));
  props.setProperty('createArch', JSON.stringify(createArch));
  props.setProperty('editAxioms', JSON.stringify(editAxioms));
  props.setProperty('editArch', JSON.stringify(editArch));
  cache.put('code1', 'available', 21600);
  cache.put('dbDRV', JSON.stringify(dbDRV), 21600);
  
  editForm.hideSheet();
  createForm.hideSheet();
};