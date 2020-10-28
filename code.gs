function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('DevelapMe Project Parser Tools')
      .addItem('Get Column Help', 'columnHelp')
      .addSeparator()
      .addToUi();
}
function onInstall(e){
  onOpen(e);
};
 
const columns = "A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z".split(",");

function columnHelp(){
  var tmp = HtmlService.createTemplateFromFile('main.html');
  var activeSheet = SpreadsheetApp.getActive().getActiveSheet();
  var activeColumnIndex = activeSheet.getActiveCell().getColumn();
  var resourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Resources (for app scripts)");
  var validationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Validation (for app scripts)");
  
  tmp.activeColumnName = activeSheet.getRange(1, activeColumnIndex).getValue();
  tmp.activeSheetName = activeSheet.getName();
  
  var resourceHeader = resourceSheet.getRange("1:1").getValues();
  var resSheetColIndex = resourceHeader[0].indexOf("sheet");
  var resColNameColIndex = resourceHeader[0].indexOf("column_name");
  var resDescriptionColIndex = resourceHeader[0].indexOf("description");
  var resInputTypeColIndex = resourceHeader[0].indexOf("input_type");
  var resValidationTypeColIndex = resourceHeader[0].indexOf("validation_type");
  var colNameValues = resourceSheet.getRange(`${columns[resColNameColIndex]}:${columns[resColNameColIndex]}`).getValues().map(val => val[0]);
  tmp.description = "";
  tmp.input_type = "not set";
  tmp.valTypes = "";
  var resRowIndexs = [];
  colNameValues.map((r,i)=>{
                    if(r.trim() === tmp.activeColumnName)
                     {
                       if(resourceSheet.getRange(i+1,resSheetColIndex+1).getValue() === tmp.activeSheetName){
                          tmp.description = resourceSheet.getRange(i+1,resDescriptionColIndex+1).getValue();
                         tmp.input_type = resourceSheet.getRange(i+1,resInputTypeColIndex+1).getValue();
                         const validationType = resourceSheet.getRange(i+1,resValidationTypeColIndex+1).getValue();
                         const valTypes = validationSheet.getRange("A:A").getValues().map(t=>t[0]);
                         const valTypeRowIndex = valTypes.indexOf(validationType)+1;
                         const valTypeRow = validationSheet.getRange(`B${valTypeRowIndex}:ZZ${valTypeRowIndex}`).getValues();
                         tmp.valTypes = valTypeRow[0].filter(v=>v.trim().length).join("<br/>");
                       }
                     }
                 
                    });

  var html = tmp.evaluate();
  SpreadsheetApp.getUi().showSidebar(html);
};
 
function getCurrentCell()
{
  var s=SpreadsheetApp.getActive().getActiveSheet().getActiveCell().getValue();
  return s;
}