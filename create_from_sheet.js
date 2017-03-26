
function onOpen() {
   var ss = SpreadsheetApp.getActiveSpreadsheet();
   var menuEntries = [];
   menuEntries.push({name: "Crear Cartas", functionName: "createLetter"});

   ss.addMenu("Cartas", menuEntries);
}
 

var DOCUMENTS_FOLDER = "0BwKY5Th4tTyfZ1dIeUFBTHhhQVk";
var TEMPLATES_FOLDER = "0BwKY5Th4tTyfUE04QjgyWXNvU2M";
var FINISHED_FOLDER = "0BwKY5Th4tTyfSlc2NC1NQnp1MXM";
var templates_basic = {
  "presidente": "1fiXtkp3HHTKD4o-hb4S1XvZlrP3AAmSZ9pwvt_KQGK8",
  "fiscal": "1lQZnZvIH6k77-vw1xcZ2lCZl2aJvDjw435-7fhr69J4",
  "primer_ministro": "1cb9gE4Q1JSdR1hkkNNDWujKHm2jhznvDjpN10lxWmtQ",
  "ministro_justicia": "1vEKucFB3GSGRaiNt8OIyZ25kcfqNbB2CUUFRXd-rBe8",
  "ministro_relaciones_exteriores": "1uU6-NqqMC3-nTqoPl-FzRg0U8cxUfjdDed2UbingypA",
  "presidente_tribunal_supremo": "1mbhviEQW1Poba5uTtjwEUHg1ss-B-kxatzeikO-ZMkg"
}

// Column B is the data holder.
var placeholders = [
  "name",
  "address",
  "postal_code",
  "city",
  "country",
  "content",
  "template",
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
var PERSON_DATA_DOC = SpreadsheetApp.getActive().getId();

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
  if(val == "✓"){
    return true;
  } else if(val == "✕") {
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

function get_or_create_folder(folder_name, parent_folder){
  var folders = parent_folder.getFoldersByName(folder_name);
  if(folders.hasNext()){
    return folders.next();
  }
  return parent_folder.createFolder(folder_name);
}

/**
 * Duplicates a Google Apps doc
 *
 * @return a new document with a given name from the orignal
 */
function createDuplicateDocument(templates_folder, name, targetFolder) {
    var source = templates_folder.getFilesByName(name).next();
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

function get_parent(doc){
  var folders = DriveApp.getFileById(doc).getParents();
  var folder = folders.next();
  return folder;
}

function get_templates(style){
  var templates_folder = DriveApp.getFolderById(TEMPLATES_FOLDER);
  var style_folder = templates_folder.getFoldersByName(style);
  if(style_folder.hasNext()){
    var folder = style_folder.next(); 
  }
  return folder;
}

function createLetter() {

  var data = SpreadsheetApp.openById(PERSON_DATA_DOC);
  var sheet = data.getSheets()[0];
  var columns = placeholders;
  var letter_data = getData(sheet);

  Logger.log("Processing columns:" + columns);
  Logger.log("Processing data:" + letter_data);
  // Assume first column holds the name of the customer
  var parent_folder = get_parent(PERSON_DATA_DOC); 
  var targetFolder = get_or_create_folder("Cartas", parent_folder);
  var templates = get_templates(letter_data["template"])
  for(letter in templates_basic){
    Logger.log("iterating: " + letter);
    Logger.log("Result: " + letter_data[letter]);
    if(letter_data[letter]){
      var target = createDuplicateDocument(templates, letter, targetFolder);
      Logger.log("Created new document:" + target.getId());

      for(var i=0; i<=5; i++){
          var ph_name = columns[i];
          var key = "@" + ph_name;
          var text = letter_data[ph_name] || "";
          replaceParagraph(target, key, text);
      }
    }
  }
}