function onOpen(e){
  var ui = SpreadsheetApp.getUi();
  ui.createAddonMenu()
  .addItem("Control Panel", "controlPanel")
  .addToUi();           
};
 
function onInstall(e){
  onOpen(e);
};
 
const columns = "A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z".split(",");

function controlPanel(){
  var tmp = HtmlService.createTemplateFromFile('main.html');
  var activeSheet = SpreadsheetApp.getActive().getActiveSheet();
  var activeColumnIndex = activeSheet.getActiveCell().getColumn();
  var resourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Resources (for app scripts)");
  
  tmp.activeColumnName = activeSheet.getRange(1, activeColumnIndex).getValue();
  tmp.activeSheetName = activeSheet.getName();
  
  var resourceHeader = resourceSheet.getRange("1:1").getValues();
  var resSheetColIndex = resourceHeader[0].indexOf("sheet");
  var resColNameColIndex = resourceHeader[0].indexOf("column_name");
  var resDescriptionColIndex = resourceHeader[0].indexOf("description");
  var colNameValues = resourceSheet.getRange(`${columns[resColNameColIndex]}:${columns[resColNameColIndex]}`).getValues().map(val => val[0]);
  tmp.description = "";
  var resRowIndexs = [];
  colNameValues.map((r,i)=>{
                    if(r.trim() === tmp.activeColumnName)
                     {
                       if(resourceSheet.getRange(i+1,resSheetColIndex+1).getValue() === tmp.activeSheetName){
                           resRowIndexs.push(resourceSheet.getRange(i+1,resDescriptionColIndex+1).getValue());
                       }
                     }
                 
                    });
  tmp.resRowIndexs = JSON.stringify(resRowIndexs);
  var html = tmp.evaluate();
  SpreadsheetApp.getUi().showSidebar(html);
};
 
function getCurrentCell()
{
  var s=SpreadsheetApp.getActive().getActiveSheet().getActiveCell().getValue();
  return s;
}