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
  
  var chartBuilder = Charts.newPieChart()
       .setTitle('Self Regulation')
       .setDimensions(600, 500)
       .set3D()
       .setDataTable(dataTable(selfReg));
  
  var chart = chartBuilder.build();
   return UiApp.createApplication().add(chart);
  //console.log("Self Reg: "+selfReg+", Ind: "+indep+", Collab: "+collab+", Init: "+initiative+", Organization: "+organ+", Responsibility: "+resp);
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
