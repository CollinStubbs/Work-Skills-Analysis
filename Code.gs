var selfReg = [['E', 0],['G',0],['S',0],['I',0],['NI',0]]; //E,G,S,I,NI
var collab = [['E', 0],['G',0],['S',0],['I',0],['NI',0]];
var indep = [['E', 0],['G',0],['S',0],['I',0],['NI',0]];
var initiative = [['E', 0],['G',0],['S',0],['I',0],['NI',0]];
var organ = [['E', 0],['G',0],['S',0],['I',0],['NI',0]];
var resp = [['E', 0],['G',0],['S',0],['I',0],['NI',0]];

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Work Skills Analysis')
      .addItem('Analyze Data', 'analyze')
      .addToUi();
  //console.log("test1");
}

function analyze() {
  var ss = SpreadsheetApp.getActive();
  
  for(var j = 7; j<12;j++){
  var sheet = ss.getSheetByName("Grade "+parseInt(j));
    console.log("Grade "+parseInt(j));
  
  var range = sheet.getDataRange().getValues();
  
  for(var i = 0; i<range.length; i++){
    switch(range[i][9]){
      case 'Self-Regulation':
        addSkill(selfReg, range[i][10]);
        break;
      case 'Independent Work':
        addSkill(indep, range[i][10]);
        break;
      case 'Collaboration':
        addSkill(collab, range[i][10]);
        break;
      case 'Initiative':
        addSkill(initiative, range[i][10]);
        break;
      case 'Organization':
        addSkill(organ, range[i][10]);
        break;
      case 'Responsibility':
        addSkill(resp, range[i][10]);
        break;
    }
  }
  
  var sss = SpreadsheetApp.create("Grade "+parseInt(j)+" - Work Skills");
    createEmbeddedChart("Self-Regulation", sss);
    createEmbeddedChart("Organization", sss);
    createEmbeddedChart("Collaboration", sss);
    createEmbeddedChart("Independent Work", sss);
    createEmbeddedChart("Initiative", sss);
  createEmbeddedChart("Responsibility", sss);
    
    sss.deleteSheet(sss.getSheetByName('Sheet1'));
  var fileId = sss.getId();
  var file = DriveApp.getFileById(fileId);
  DriveApp.getFoldersByName('Work Skills Analysis').next().addFile(file);
  }
}

function createEmbeddedChart(skillName, spread){
  var sheet = spread.insertSheet(skillName);
  
   switch(skillName){
      case 'Self-Regulation':
        var range = sheet.getRange(1,1,selfReg.length,selfReg[0].length).setValues(selfReg);
        break;
      case 'Independent Work':
          var range = sheet.getRange(1,1,indep.length,indep[0].length).setValues(indep);
       break;
      case 'Collaboration':
          var range = sheet.getRange(1,1,collab.length,collab[0].length).setValues(collab);
        break;
      case 'Initiative':
          var range = sheet.getRange(1,1,initiative.length,initiative[0].length).setValues(initiative);
        break;
      case 'Organization':
          var range = sheet.getRange(1,1,organ.length,organ[0].length).setValues(organ);
        break;
      case 'Responsibility':
          var range = sheet.getRange(1,1,resp.length,resp[0].length);
       range.setValues(resp);
       
        break;
    }        

  
  var chart = sheet.newChart()
  .setChartType(Charts.ChartType.PIE)
  .addRange(range)
  .setPosition(4, 4, 0, 0)
  .setOption('title', skillName)
  .setOption('legend', {textStyle: {fontSize: 14, bold: true, }})
  .setOption('pieSliceText', 'value')
  .build()
 

 sheet.insertChart(chart);
}


//increases the count for that skills rating
function addSkill(skill, rating){
  switch(rating){
    case 'E':
      skill[0][1]++;
      break;
    case 'G':
      skill[1][1]++;
      break;
    case 'S':
      skill[2][1]++;
      break;
    case 'I':
      skill[3][1]++;
      break;
    case 'NI':
      skill[4][1]++;
      break;
  }
  
}
