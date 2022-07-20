
function showFormInSidebar() {      
  var form = HtmlService.createTemplateFromFile('AA_Index').evaluate().setTitle('Informations de Login');
  SpreadsheetApp.getUi().showSidebar(form);
}

//PROCESS FORM
function processForm(formObject){ 
  if(formObject.uSfdc!==""){setUsername(formObject.uSfdc);}
  if(formObject.mdpSfdc!==""){setPassword(formObject.mdpSfdc);}
  if(formObject.keySfdc!==""){setSecurityToken(formObject.keySfdc);}  
  Browser.msgBox("Data updated. New user : "+formObject.uSfdc)
}

//INCLUDE HTML PARTS, EG. JAVASCRIPT, CSS, OTHER HTML FILES
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
