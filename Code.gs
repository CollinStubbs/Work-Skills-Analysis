var selfReg = [0,0,0,0,0]; //E,G,S,I,NI
var collab = [0,0,0,0,0];
var indep = [0,0,0,0,0];
var initiative = [0,0,0,0,0];
var organ = [0,0,0,0,0];
var resp = [0,0,0,0,0];

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Work Skills Analysis')
      .addItem('Analyze Data', 'analyze')
      .addToUi();
  //console.log("test1");
}

function analyze() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("Grade 7");
  
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
  
  var ss = SpreadsheetApp.create("Grade 7 - Work Skills");
  ss.insertSheet("Self-Regulation").insertImage(getChartIMG(getChart('Self-Regulation')), 1, 1);
  ss.insertSheet("Organization").insertImage(getChartIMG(getChart('Organization')), 1, 1);
  ss.insertSheet("Collaboration").insertImage(getChartIMG(getChart('Collaboration')), 1, 1);
  ss.insertSheet("Independent Work").insertImage(getChartIMG(getChart('Independent Work')), 1, 1);
  ss.insertSheet("Initiative").insertImage(getChartIMG(getChart('Initiative')), 1, 1);
  ss.insertSheet("Responsibility").insertImage(getChartIMG(getChart('Responsibility')), 1, 1);
  
  var fileId = ss.getId();
  var file = DriveApp.getFileById(fileId);
  DriveApp.getFoldersByName('Work Skills Analysis').next().addFile(file);
 
}
 
function doGet() {
  return HtmlService
      .createTemplateFromFile('skills_analysis')
      .evaluate();
}

function getChartIMG(chart) {
 return chart.getBlob().getAs('image/png').setName("areaBlob"); 
}

function getChart(skillName, sheet){
  
   var chartBuilder = Charts.newPieChart()
       .setTitle(skillName)
       .setDimensions(600, 500);
  
     switch(skillName){
      case 'Self-Regulation':
        chartBuilder.setDataTable(dataTable(selfReg));
        break;
      case 'Independent Work':
         chartBuilder.setDataTable(dataTable(indep));
         break;
      case 'Collaboration':
         chartBuilder.setDataTable(dataTable(collab));
        break;
      case 'Initiative':
         chartBuilder.setDataTable(dataTable(initiative));
        break;
      case 'Organization':
         chartBuilder.setDataTable(dataTable(organ));
        break;
      case 'Responsibility':
         chartBuilder.setDataTable(dataTable(resp));
        break;
    }          
  
  return chartBuilder.build();
}
function dataTable(rating){
  var data = Charts.newDataTable()
  .addColumn(Charts.ColumnType.STRING, "Rating")
  .addColumn(Charts.ColumnType.NUMBER, "Count")
  .addRow(['E', rating[0]])
  .addRow(['G', rating[1]])
  .addRow(['S', rating[2]])
  .addRow(['I', rating[3]])
  .addRow(['NI', rating[4]])
  .build();
  return data;
}

//increases the count for that skills rating
function addSkill(skill, rating){
  switch(rating){
    case 'E':
      skill[0]++;
      break;
    case 'G':
      skill[1]++;
      break;
    case 'S':
      skill[2]++;
      break;
    case 'I':
      skill[3]++;
      break;
    case 'NI':
      skill[4]++;
      break;
  }
  
}
