function BUmap(props) {
  
  var BackupFolder = DriveApp.getFolderById('0B83_VULstkR9NUNzVWM5YWh4T2s');
  var files = BackupFolder.getFiles();
  var buMap = [];
  var i = 0;
  
  while (files.hasNext()){
    var file = files.next();
    buMap[i]  = [file.getName(), file.getId()]
    i++;
  }
  
  props.setProperty('buMap', JSON.stringify(buMap));
};

function updateBU(backUpGroup, newID, props){
  
  var parentID = newID.split('-')[0];
  var num1 = Math.floor(parentID/10) * 10;
  var fileName = num1 + '-' + (num1+9);
  var sheetName = parentID;
  var BUss = null;
  
  var buMap = JSON.parse(props.getProperty('buMap'));
  var l = buMap.length;
  for (i=0; i<l; i++) {
    if (buMap[i][0] == fileName) {
      var BUss = SpreadsheetApp.openById(buMap[i][1]);
      break;
    };
  };
  
  if (BUss == null){
    var BackupFolder = DriveApp.getFolderById('0B83_VULstkR9NUNzVWM5YWh4T2s');
    var newFile = DriveApp.getFileById('1IWcmDAD3H-0pTYlcUEP1WjR-j9019HB8n4N1RiJ3S2Q').makeCopy(BackupFolder).setName(fileName);
    var BUss = SpreadsheetApp.openById(newFile.getId())
    
    buMap.push([fileName, newFile.getId()]);
    props.setProperty('buMap', JSON.stringify(buMap));
    
    var sheets = BUss.getSheets();
    for (i=0; i<10; i++) {
      sheets[i].setName(num1+i);
    };
  };
  
  var buSheet = BUss.getSheetByName(sheetName);
  var index = parseInt(newID.split('-')[1]);
  
  var lastReqC = (index+1) * 5;
  var lastC = buSheet.getLastColumn();
  if (lastReqC > lastC) {
    buSheet.insertColumnsAfter(lastC, (lastReqC - lastC));
  };
  if (index != 0) {buSheet.getRange(1, 1, 2, 5).copyTo(buSheet.getRange(1, (1 + index*5), 2, 5))};
  buSheet.getRange(1, (3 + index*5)).setValue(newID);
  buSheet.getRange(3, (1 + index*5), backUpGroup.length, 5).setValues(backUpGroup);
};

function editBackUp(ID, backUpR, props) {

  var parentID = ID.split('-')[0];
  var num1 = Math.floor(parentID/10) * 10;
  var fileName = num1 + '-' + (num1+9);
  var sheetName = parentID;
  var BUss = null;
  
  var buMap = JSON.parse(props.getProperty('buMap'));
  var l = buMap.length;
  for (i=0; i<l; i++) {
    if (buMap[i][0] == fileName) {
      var BUss = SpreadsheetApp.openById(buMap[i][1]);
      break;
    };
  };
  
  var buSheet = BUss.getSheetByName(sheetName);
  var index = parseInt(ID.split('-')[1]);
  
  var maxR = buSheet.getLastRow();
  var lastR = buSheet.getRange(1, index*5 +1, maxR, 1).getValues().filter(String).length;
  buSheet.getRange(lastR+1, index*5 +1, 1, 5).setValues([backUpR]);
};