/**
* A special function thats run when the spreadsheet is open.
* Adds the menu to the spreadsheet.
* Also launches the sidebar for adding new items.
*/
function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu('Suspension Options')
  .addItem('Add Suspension', 'launchAdder')
  .addSeparator()
  .addItem('Add New Year', 'newSheet')
  .addItem('Repair Title Bar', 'repairTitle')
  .addToUi();
  
  launchAdder();
}

/**
* Launches the sidebar to add new suspension items.
*
*/
function launchAdder() {
  var html = HtmlService.createHtmlOutputFromFile('adder')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('New Suspension')
      .setWidth(300);
  SpreadsheetApp.getUi()
      .showSidebar(html);
}

/**
* Launches the message box for setting the sheet name.
*
*/
function launchSheetBox() {
  var htmlDlg = HtmlService.createHtmlOutputFromFile('newSheetBox')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(400)
      .setHeight(125);
  SpreadsheetApp.getUi()
      .showModalDialog(htmlDlg, 'New Order Sheet Name');
}

/**
* Add a new line to the top of the sheet and populate it. 
* 
* @param {String} date, location, product, problem, ra, cus, cusPhone, cusEmail, sro, wo, status
* The information for the piece of suspension to be added.
*
*/
function addNewTop(date, location, product, problem, ra, cus, cusPhone, cusEmail, sro, wo, status) {
  // Get the active sheet and spreadsheet
  var ss = SpreadsheetApp.getActive();
  var s = ss.getActiveSheet();
  
  // Add a new row to the top of the list.
  s.insertRowBefore(2);
  
  // Add the info to the newly created row.
  s.getRange(2, 1, 1, 11).setValues([[date,location,product,problem,ra,cus,cusPhone,cusEmail,sro,wo,status]]);
}

/**
* Adds a new sheet to the spreadsheet for a new year.
*/
function newSheet() {
  // Get the active sheet and spreadsheet
  var ss = SpreadsheetApp.getActive();
  var s = ss.getActiveSheet();
  
  // Add a new sheet the the spread sheet based on a template.
  var templateSheet = ss.getSheetByName('templateSheet');
  ss.insertSheet(0, {template: templateSheet});
  
  // Get the list of sheet names.
  var sheets = ss.getSheets();
  
  // Name the new sheet the next year.
  s = ss.getActiveSheet();
  // If the last year is not a number ask for the new sheet name.
  // Else add one to the last year and make that the name.
  if(isNaN(sheets[1].getName())) {
    launchSheetBox();
  } else {
    s.setName(String(Number(sheets[1].getName()) + 1));
  }
  
  // Add title bar protection.
  var range = s.getRange('1:1');
  range.protect().setDescription('Title Bar');
}

/**
* Sets the sheet name to the desired value.
*
* @param {String} name The new name of the active sheet.
*/
function setSheetName(name) {
  // Get the active sheet and spreadsheet
  var ss = SpreadsheetApp.getActive();
  var s = ss.getActiveSheet();
  s.setName(name);
}

/**
* Repairs the title bar if accidental changes are made.
*
*/
function repairTitle() {
  // Get the active sheet and spreadsheet
  var ss = SpreadsheetApp.getActive();
  var s = ss.getActiveSheet();
  
  // Get the range of the title bar
  var range = s.getRange('A1:K1');
  
  // Sets the colour and font to the proper format.
  range.setBackground('#38761d').setFontColor('white').setFontFamily('Arial').setFontWeight('bold').setFontSize(10);
  
  // Sets the correct words
  range.setValues([['Date','Location','Product','Problem Description','RA #','Customer','Phone','Email','SRO #','WO #','Status']]);
}

/**
// Old function that adds to the botom of the sheet
function addNew(date, location, product, problem, ra, cus, cusPhone, cusEmail, sro, wo, status) {
  var ss = SpreadsheetApp.getActive();
  var s = ss.getActiveSheet();
  
  // Find fist emepty row.
  var range = s.getRange(2, 1, s.getMaxRows());
  var values = range.getValues();
  var empty = 0;
  for(var i=0; i<values.length; i++) {
    if(values[i][0] == 0) {
      empty = i + 2;
      break;
    }
  }
  // Check if the row is actually empty!
  range = s.getRange(empty, 1, 1, s.getMaxColumns());
  values = range.getValues()
  for(var i=0; i<values[0].length; i++) {
    // Add "needs date" to a row missing the date and trys again.
    if (values[0][i] !=0) {
      s.getRange(empty, 1).setValue("needs date!");
      addNew(date, location, product, problem, ra, cus, cusPhone, cusEmail, sro, wo, status);
      return;
    }
  }
  // Add the info to the new row.
  s.getRange(empty, 1, 1, 11).setValues([[date,location,product,problem,ra,cus,cusPhone,cusEmail,sro,wo,status]]);
} */

/**
* Scrolls to the last line of the list.
*
*/
function scrollLast() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mysheet = ss.getActiveSheet();

  var lastrow = mysheet.getLastRow();
  Logger.log(lastrow);
  mysheet.setActiveCell(mysheet.getDataRange().offset(lastrow, 0, 1, 1));
}

function test() {
  repairTitle();
  //launchSheetBox();
}
