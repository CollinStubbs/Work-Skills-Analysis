var selfReg = [['E', 0],['G',0],['S',0],['I',0],['NI',0]]; //E,G,S,I,NI
var collab = [['E', 0],['G',0],['S',0],['I',0],['NI',0]];
var indep = [['E', 0],['G',0],['S',0],['I',0],['NI',0]];
var initiative = [['E', 0],['G',0],['S',0],['I',0],['NI',0]];
var organ = [['E', 0],['G',0],['S',0],['I',0],['NI',0]];
var resp = [['E', 0],['G',0],['S',0],['I',0],['NI',0]];

var NITracker = [[0,0,0,0,0,0],[0,0,0,0,0,0],[0,0,0,0,0,0],[0,0,0,0,0,0],[0,0,0,0,0,0]];
var folder = null;

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Work Skills Analysis')
  .addItem('Analyze Data', 'analyze')
  .addToUi();
  //console.log("test1");
}

function analyze() {
  var ss = SpreadsheetApp.getActive();
  var currentD = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMM yyyy");
  
  folder = DriveApp.createFolder('Work Skills Analysis - '+currentD);
  
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
    // DriveApp.getFoldersByName('Work Skills Analysis').next().addFile(file);
    folder.addFile(file);
  
    NITracker[j-7][0] = selfReg[4][1];
    NITracker[j-7][1] = organ[4][1];
    NITracker[j-7][2] = collab[4][1];
    NITracker[j-7][3] = indep[4][1];
    NITracker[j-7][4] = initiative[4][1];
    NITracker[j-7][5] = resp[4][1];
    
    
    reset();
  }
  createDataPage();
}

function createDataPage(){
  var mean = 0;
  var means = [0,0,0,0,0,0];
  var grades = [0,0,0,0,0];
  var count = 0;
  
  for(var i = 0; i<NITracker.length; i++){
    for(var j = 0; j<NITracker[0].length; j++){
      switch(j){
        case 0:
          means[j]+=NITracker[i][j];
          break;
        case 1:
          means[j]+=NITracker[i][j];
          break;
        case 2:
          means[j]+=NITracker[i][j];
          break;
        case 3:
          means[j]+=NITracker[i][j];
          break;
        case 4:
          means[j]+=NITracker[i][j];
          break;
        case 5:
          means[j]+=NITracker[i][j];
          break;
        default:
          break;          
      }
      grades[i]= grades[i]+NITracker[i][j];
      mean = mean+NITracker[i][j];
      count++;
    }
  }
  mean = mean/5;
  for(var i = 0; i<6;i++){
   means[i]=means[i]/5; 
  }
  
  var ss = SpreadsheetApp.create('Data Analysis');
  var sheet = ss.getActiveSheet();
  var fileId = ss.getId();
  var file = DriveApp.getFileById(fileId);
  folder.addFile(file);
  
  
      
  sheet.getRange(1, 3, 1, 1).setValues([['Mean NI\'s']]);
  sheet.getRange(1, 4, 1, 1).setValues([['Self-Regulation NI\'s']]);
  sheet.getRange(1, 5, 1, 1).setValues([['Organization NI\'s']]);
  sheet.getRange(1, 6, 1, 1).setValues([['Collaboration NI\'s']]);
  sheet.getRange(1, 7, 1, 1).setValues([['Independent Work NI\'s']]);
  sheet.getRange(1, 8, 1, 1).setValues([['Initiative NI\'s']]);
  sheet.getRange(1, 9, 1, 1).setValues([['Responsibility NI\'s']]);
  
  sheet.getRange(2, 2, 1, 8).setValues([['Grade 9\'s', means[0], NITracker[0][0], NITracker[0][1], NITracker[0][2], NITracker[0][3], NITracker[0][4], NITracker[0][5]]]);
  sheet.getRange(3, 2, 1, 8).setValues([['Grade 9\'s', means[0], NITracker[1][0], NITracker[1][1], NITracker[1][2], NITracker[1][3], NITracker[1][4], NITracker[1][5]]]);
  sheet.getRange(4, 2, 1, 8).setValues([['Grade 9\'s', means[0], NITracker[2][0], NITracker[2][1], NITracker[2][2], NITracker[2][3], NITracker[2][4], NITracker[2][5]]]);
  sheet.getRange(5, 2, 1, 8).setValues([['Grade 9\'s', means[0], NITracker[3][0], NITracker[3][1], NITracker[3][2], NITracker[3][3], NITracker[3][4], NITracker[3][5]]]);
  sheet.getRange(6, 2, 1, 8).setValues([['Grade 9\'s', means[0], NITracker[4][0], NITracker[4][1], NITracker[4][2], NITracker[4][3], NITracker[4][4], NITracker[4][5]]]);
  
  sheet.getRange(8, 1, 1, 2).setValues([['Mean NI Count per Grade', mean]]);
  sheet.getRange(9, 1, 1, 2).setValues([['NI count for 7', grades[0]]]);
  sheet.getRange(10, 1, 1, 2).setValues([['NI count for 8', grades[1]]]);
  sheet.getRange(11, 1, 1, 2).setValues([['NI count for 9', grades[2]]]);
  sheet.getRange(12, 1, 1, 2).setValues([['NI count for 10', grades[3]]]);
  sheet.getRange(13, 1, 1, 2).setValues([['NI count for 11', grades[4]]]);
  
}

function reset(){
  selfReg = [['E', 0],['G',0],['S',0],['I',0],['NI',0]]; //E,G,S,I,NI
  collab = [['E', 0],['G',0],['S',0],['I',0],['NI',0]];
  indep = [['E', 0],['G',0],['S',0],['I',0],['NI',0]];
  initiative = [['E', 0],['G',0],['S',0],['I',0],['NI',0]];
  organ = [['E', 0],['G',0],['S',0],['I',0],['NI',0]];
  resp = [['E', 0],['G',0],['S',0],['I',0],['NI',0]]; 
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
  .setOption('legend', {textStyle: {fontSize: 14, bold: true}})
  .setOption('pieSliceText', 'value')
  .setOption('titleTextStyle', {color: 'black', bold: true})
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
