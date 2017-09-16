// for semi-complex dynamic protections
function createArchitecture(sheet, phase, index, howMany, newFields, protect, createArch, props) {
  
  createArch.phase = phase;
  
  if (phase == 'closed') {protectSheet(sheet, null)}
  
  if (howMany != 0) {
    var l = createArch.createR.length;
    for (i=0; i<l; i++) {
      if (createArch.createR[i] > index) {
        createArch.createR[i] += howMany;
      };
    };
    
    var l = createArch.createTables.length;
    for (i=0; i<l; i++) {
      if (createArch.createTables[i][0] > index) {
        createArch.createTables[i][0] += howMany
      };
    };
    
    if (howMany > 0) {
      sheet.insertRowsAfter(index, howMany);
    } else if (howMany < 0) {
      sheet.deleteRows((index+howMany+1),(howMany*(-1)))
    };
  };
  
  props.setProperty('createArch', JSON.stringify(createArch));
  
  if (protect == true) {
    var createA1Not = [];
    var createA1NotShort = [];
    var createDRr = [];
    var l = createArch.createC.length;
    for (i=0; i<l; i++) {
      createA1NotShort.push(createArch.createCshort[i]+createArch.createR[i])
      createDRr.push(createArch.createR[i] - 1);
      var brk = createArch.createC[i].split('');
      var index2 = brk.indexOf(':');
      if (index2 != -1) {
        brk.splice(index2, 0, createArch.createR[i]);
        brk.push(createArch.createR[i])
        createA1Not.push(brk.join(''));
      } else {
        createA1Not.push(createArch.createC[i]+createArch.createR[i]);
      };
    };
    
    var unprotectedRanges = createA1Not;
    if (newFields != null) {unprotectedRanges = unprotectedRanges.concat(newFields)};
    var l = createArch.createTables.length;
    for (i=0; i<l; i++) {
      unprotectedRanges = unprotectedRanges.concat(createArch.createTables[i][0]  + ',' + createArch.createTables[i][2] + ',' + createArch.createTables[i][1] + ',' + createArch.createTables[i][3]);
    };
    protectSheet(sheet, unprotectedRanges);
  };
  
  return createArch
  
};