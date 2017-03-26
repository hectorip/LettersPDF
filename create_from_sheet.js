// Row number from where to fill in the data (starts as 1 = first row)
var CUSTOMER_ID = 2;

var DOCUMENTS_FOLDER = "0BwKY5Th4tTyfZ1dIeUFBTHhhQVk";
var templates_basic = {
  "presidente": "",
  "fiscal": "1IzKtPMpkOs9IvmqB7JSoaoD_vIm8bgyRrLmqUPauGM0",
  "primer_ministro": "",
  "ministro_justicia": "",
  "ministro_relaciones_exteriores": "",
  "presidente_tribunal_supremo": ""
}
// Column B is the data holder, 
var placeholders = [
  "name",
  "address",
  "postal_code",
  "city",
  "country",
  "content",
  "",
  "",
  "presidente",
  "fiscal",
  "primer_ministro",
  "ministro_justicia",
  "ministro_relaciones_exteriores",
  "presidente_tribunal_supremo"
  ]

// In which spreadsheet we have all the customer data
var PERSON_DATA_DOC = "1q4kAVi4Y74y3Lco23iBZn0qWMSE6grNnLjGW75Uh67s";

function getData(sheet){
  var dataRange = sheet.getRange(1, 2, 15);
  var data = dataRange.getValues();
  var letter_data = {};

  for (i in data) {
    var row = data[i];
    var placeholder = placeholders[i];
    if(placeholder){
      letter_data[placeholder] = map_symbols(row[0]);
    }
  }
  Logger.log(letter_data);
  return letter_data;
}

function map_symbols(val){
  if(val="✓"){
    return true;
  } else if(val="✕") {
    return false;
  }
  return val;
}

function getRowAsArray(sheet, row) {
  var dataRange = sheet.getRange(row, 1, 1, 99);
  var data = dataRange.getValues();
  var columns = [];

  for (i in data) {
    var row = data[i];
    for(var l=0; l<99; l++) {
        var col = row[l];
        // First empty column interrupts
        if(!col) {
            break;
        }

        columns.push(col);
    }
  }

  return columns;
}

function get_or_create_folder(folder_name){
  var documents_folder = DriveApp.getFolderById(DOCUMENTS_FOLDER);
  var folders = targetFolder.getFoldersByName(folder_name);
  if(folders){
    return folders[0];
  }
  return documents_folder.createFolder(folder_name);
}

/**
 * Duplicates a Google Apps doc
 *
 * @return a new document with a given name from the orignal
 */
function createDuplicateDocument(sourceId, name, foler_name) {
    var source = DriveApp.getFileById(sourceId);
    var targetFolder = get_or_create_folder(folder_name);
    var newFile = source.makeCopy(name, targetFolder);

    return DocumentApp.openById(newFile.getId());
}

/**
 * Search a paragraph in the document and replaces it with the generated text 
 */
function replaceParagraph(doc, keyword, newText) {
  var ps = doc.getParagraphs();
  for(var i=0; i<ps.length; i++) {
    var p = ps[i];
    var text = p.getText();

    if(text.indexOf(keyword) >= 0) {
      p.clear()
      p.setText(newText);
    }
  } 
}

/**
 * Script entry point
 */
function createLetter() {

  var data = SpreadsheetApp.openById(PERSON_DATA_DOC);
  //var CUSTOMER_ID = Browser.inputBox("Enter customer number in the spreadsheet", Browser.Buttons.OK_CANCEL);

  // Fetch variable names
  // they are column names in the spreadsheet
  var sheet = data.getSheets()[0];
  // var columns = getRowAsArray(sheet, 1);
  var columns = placeholders;

  Logger.log("Processing columns:" + columns);

  var customerData = getRowAsArray(sheet, CUSTOMER_ID);
  var letterData = getData(sheet);
  Logger.log("Processing data:" + customerData);
  return
  // Assume first column holds the name of the customer
  var customerName = customerData[0];

  var target = createDuplicateDocument(SOURCE_TEMPLATE, customerName + " Letter");

  Logger.log("Created new document:" + target.getId());

  for(var i=0; i<columns.length; i++) {
      var key = "@" + columns[i];
      var text = customerData[i] || "";
      var value = text;
      replaceParagraph(target, key, value);
  }

}