function protectSheet(sheet, unprotectedCells){
  
  
  var me = Session.getEffectiveUser();
  var protection = sheet.protect();
  
  if(unprotectedCells == null){
    protection.addEditor(me);
    protection.removeEditors(protection.getEditors());
    if (protection.canDomainEdit()) {
      protection.setDomainEdit(false);
    }
    return;
  }
  
  var l = unprotectedCells.length;
  
  if( l == 1){
    protection.setUnprotectedRanges([sheet.getRange(unprotectedCells[0])])
    protection.addEditor(me);
    protection.removeEditors(protection.getEditors());
    if(protection.canDomainEdit()) {
      protection.setDomainEdit(false);
    }
    return;
  }
  
  if(l>1){
    var unprotectedRanges = []
    for (i=0; i<l; i++){
      var temp = unprotectedCells[i].split(',');
      if (isNaN(temp[0]) == false) {
        unprotectedRanges = unprotectedRanges.concat(sheet.getRange(temp[0], temp[1], temp[2], temp[3]));
        continue;
      } else {
        unprotectedRanges = unprotectedRanges.concat(sheet.getRange(unprotectedCells[i]));
      };
    }
    protection.setUnprotectedRanges(unprotectedRanges);
    protection.addEditor(me);
    protection.removeEditors(protection.getEditors());
    if(protection.canDomainEdit()){
      protection.setDomainEdit(false);
    }
    return;
  }
}