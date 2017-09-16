function fillInActivity2(sheet, editAxioms, IDrow, activities, props) {
  
  var tableSheet = SpreadsheetApp.getActive().getSheetByName('tables');
  
  var fields = [
    {
      'fieldName': 'ID',
      'columns': 'H',
      'columnsShort': 'H',
      'row': 5,
      'DRc': 7,
      'DBc': 0
    },
    {
      'fieldName': 'Activitiy Name',
      'columns': 'B:F',
      'columnsShort': 'B',
      'row': 5,
      'DRc': 7,
      'DBc': 0
    },
    {
      'fieldName': 'Start',
      'columns': 'E:F',
      'columnsShort': 'E',
      'row': 7,
      'DRc': 4,
      'DBc': 4
    }
    
  
  ];
  
}