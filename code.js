const doc_id = "<id шаблона файла>";
const folder_temp_id = "<id временной папки>";
const folder_pdf_id = "<id папки в которую будут падать пдф>";
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("dataBase");

function createBulkPDFs() {
  
  var data = sheet.getRange(2, 1, sheet.getLastRow()-1, 6).getValues();
  var result = data.filter( // Фильтруем массив на пустые ячейки и флажок
    function(item){
      if (item[4] === true) {
        return !item.some(function(cell) {
          return cell === "";
        })
      }
    });
  
  data.forEach((item, index) => {
               
               try{
               
               if (item[0] === true){
    
               createPDF(item[3], item[4], item[5], item[3] + " " + item[4], doc_id, folder_temp_id, folder_pdf_id);
  
  var currentDate = Utilities.formatDate(new Date(), "GMT+3", "Дата dd.MM.yyyy\nВремя HH:mm:ss:SSS");
               item.splice(0, 2)
               item.unshift(true, "Файл создан")
               sheet.getRange(index+2, 2)
               .setValue(item[1])
               .setNote(currentDate)
               .setFontColor('green')
                 } // End if (data[0] == true){
               } catch(err) {
               sheet.getRange(index+2, 2)
               .setValue("Ошибка")
               .setNote(err)
               .setFontColor('red')
               } // End catch(err)
           }) // End forEach

           

}

function createPDF(firstName, lastName, balance, pdfName, doc_id, folder_temp_id, folder_pdf_id) {
  
  var docFile = DriveApp.getFileById(doc_id);
  var tempFolder = DriveApp.getFolderById(folder_temp_id);
  var pdfFolder = DriveApp.getFolderById(folder_pdf_id);
  var tempFile = docFile.makeCopy(tempFolder);
  var tempDocFile = DocumentApp.openById(tempFile.getId());
  var body = tempDocFile.getBody();
  body.replaceText("{firstName}", firstName);
  body.replaceText("{lastName}", lastName);
  body.replaceText("{balance}", balance);
  tempDocFile.saveAndClose();
  
  var pdfContentBlob = tempDocFile.getAs(MimeType.PDF);
 
  pdfFolder.createFile(pdfContentBlob).setName(pdfName);
  tempFolder.removeFile(tempFile);
  
}


function filesFolder(){
  var pdfFolder = DriveApp.getFolderById(folder_pdf_id);
  var files = pdfFolder.getFiles();
  
  while (files.hasNext()) {
    var file = files.next();
    Logger.log(file.getName());
}

//  console.log(files)
}


function setTrue() {
  var data = sheet.getRange(2, 1, sheet.getLastRow()).getValues();
  
  data.forEach((item, index) => {
  if (item != "") {
    item = true;
    sheet.getRange(index+2, 1)
    .setValue(item) 
  }
})
}


function setFalse() {
  var data = sheet.getRange(2, 1, sheet.getLastRow()).getValues();
  
  data.forEach((item, index) => {
  if (item != "") {
    item = false;
    sheet.getRange(index+2, 1)
    .setValue(item) 
  }
})
}


function clearStatus() {
  var range = sheet.getRange(2, 2, sheet.getLastRow())
  .clear()
  .clearNote()
}
