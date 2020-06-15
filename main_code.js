
  function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //Appears on the menu bar of current active spreadsheet 
  var menu = [{
  name: "JPEG to Excel",
  functionName: "convertJPEGtoGoogleDocs"
  }, {
  name: "Clear value",
  functionName: "clear_value"
  }, ];
    //The new menu called "Custo" is added onto the spreadsheet. 
    ss.addMenu("Custo", menu);
    }
    
    function convertJPEGtoGoogleDocs() {

    var Input_folder_ID  = Browser.inputBox('Enter Folder URL','Insert Google drive folder URL',Browser.Buttons.OK_CANCEL);
    var srcfolderId = getIdFromUrl(Input_folder_ID);

    // Folder ID for destination of the files to be converted
    var dstfolderId = DriveApp.getFolderById(srcfolderId).createFolder('Doc').getId();
    var folder_name = DriveApp.getFolderById(srcfolderId).getName()
    var files_JPEG = DriveApp.getFolderById(srcfolderId).getFilesByType(MimeType.JPEG); 

    //Rename the folder name to translated file name.
    SpreadsheetApp.getActiveSpreadsheet().setName('Translation '+ folder_name);

    var ss = SpreadsheetApp.getActiveSpreadsheet();//Current google sheet
    var sheet = ss.getSheetByName('Sheet1');//Specified sheet

    var files_GoogleDoc = DriveApp.getFolderById(dstfolderId).getFilesByType(MimeType.GOOGLE_DOCS);
    var i = 1;

    while (files_JPEG.hasNext()) {
    var file_object = files_JPEG.next();
    Drive.Files.insert({title: file_object.getName(),
    mimeType: MimeType.GOOGLE_DOCS,
    parents: [{id: dstfolderId}]}, 
      file_object.getBlob(), 
        {ocr: false}); 
  }

  while (files_GoogleDoc.hasNext()) {
  var file_object = files_GoogleDoc.next(); 
  var file_ID = file_object.getId();
  var file_name = file_object.getName();
  var docid = DocumentApp.openById(file_ID);
  var file_content = docid.getBody().getText();
  i++;
  var cell_A = 'A' + i
  var cell_B = 'B' + i
  ss.getRange(cell_A).setValue(file_name);
  ss.getRange(cell_B).setValue(file_content);
  }

  //Add set value on C2 and E2 cells
  var setValue_on_C2 = sheet.getRange('C2').setValue('=GOOGLETRANSLATE(B2,"ja","en")');
  var setValue_on_E2 = sheet.getRange('E2').setValue('=upper(D2)');

  //Drag and drop the C2 and E2 cells
  var destinationRange_C2 = sheet.getRange('C2').offset(0, 0, sheet.getLastRow() - 1);
  //autofill function on C column
  sheet.getRange('C2').autoFill(destinationRange_C2, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

  var destinationRange_D2 = sheet.getRange('D2').offset(0, 0, sheet.getLastRow() - 1);
  destinationRange_C2.copyTo(destinationRange_D2, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

  var destinationRange_E2 = sheet.getRange('E2').offset(0, 0, sheet.getLastRow() - 1);
  //autofill function on E column
  sheet.getRange('E2').autoFill(destinationRange_E2, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

  //Select the entire sheet and text wrap enable
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).activate();
  sheet.getActiveRangeList().setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  }

  //This will remove this string "https://drive.google.com/open?id="
  function getIdFromUrl(url) {
  return url.match(/[-\w]{25,}/);
  }

  function clear_value() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();//Current google sheet
  var sheet = ss.getSheetByName('Sheet1');//Specified sheet
  //Clear all cells from A2 to bottom right. 
  var clear_value = sheet.getRange(2, 1, sheet.getLastRow(), 5).clear();
  //If the number or row is larger than 25 than deleted those rows to keep things simple.
  if(sheet.getMaxRows() > 25) {
  var deleterow = sheet.deleteRows(25,sheet.getMaxRows()-25);
  }
  }
