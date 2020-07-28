/*  Code source: https://github.com/hadaf
 *  This is the main method that should be invoked. 
 *  Copy and paste the ID of your template Doc in the first line of this method.
 *
 *  Make sure the first row of the data Sheet is column headers.
 *
 *  Reference the column headers in the template by enclosing the header in square brackets.
 *  Example: "This is [header1] that corresponds to a value of [header2]."
 */

function doMerge(rowNum, subNumber, lastName, folderID, templateID, spreadsheetID, sheetName) {
  let selectedTemplateId = templateID;
  
  let templateFile = DriveApp.getFileById(selectedTemplateId);
  let targetFolder = DriveApp.getFolderById(folderID);
  let mergedFile = templateFile.makeCopy(targetFolder); //make a copy of the template file to use for the merged File. 
  // Note: It is necessary to make a copy upfront, and do the rest of the content manipulation inside this single copied file, 
  // otherwise, if the destination file and the template file are separate, a Google bug will prevent copying of images from the 
  // template to the destination. See the description of the bug here: https://code.google.com/p/google-apps-script-issues/issues/detail?id=1612#c14
  mergedFile.setName(subNumber + " Professional Reimbursement - " + lastName);//give a custom name to the new file (otherwise it is called "copy of ...")
  let mergedDoc = DocumentApp.openById(mergedFile.getId());
  let bodyElement = mergedDoc.getBody();//the body of the merged document, which is at this point the same as the template doc.
  let bodyCopy = bodyElement.copy();//make a copy of the body
  
  bodyElement.clear();//clear the body of the mergedDoc so that we can write the new data in it.
  
  let ss = SpreadsheetApp.openById(spreadsheetID);
  let sheet = ss.getSheetByName(sheetName);
  
  let rows = sheet.getDataRange();
  let numRows = rows.getNumRows();
  let values = rows.getValues();
  let fieldNames = values[0];//First row of the sheet must be the the field names

  let data = sheet.getDataRange().getValues();
  //Logger.log(">>"+data);
  
      let row = values[rowNum-1];
      let body = bodyCopy.copy();
      
      // Match field names with data and replace on template
      for (let f = 0; f < fieldNames.length; f++) { 
        if(fieldNames[f].substring(fieldNames[f].length - 4, fieldNames[f].length)=="Cost" && typeof row[f] == "number") {
          body.replaceText("\\[" + fieldNames[f] + "\\]", "$" + row[f].toFixed(2));
        } else if (f<12 && fieldNames[f].substring(fieldNames[f].length - 4, fieldNames[f].length)=="Date") {
          Logger.log(typeof row[f]);
          body.replaceText("\\[" + fieldNames[f] + "\\]", Utilities.formatDate(row[f], "EST", "MMM dd, yyyy"));
        } else {
          body.replaceText("\\[" + fieldNames[f] + "\\]", row[f]);//replace [fieldName] with the respective data value
        }
      }
    
      let date = Utilities.formatDate(new Date(), "EST", "MMM dd, yyyy");
      body.replaceText("\\[Today\\]", date);
      let numChildren = body.getNumChildren();//number of the contents in the template doc
     
      for (let c = 0; c < numChildren; c++) {//Go over all the content of the template doc, and replicate it for each row of the data.
        let child = body.getChild(c);
        child = child.copy();
        if (child.getType() == DocumentApp.ElementType.HORIZONTALRULE) {
          mergedDoc.appendHorizontalRule(child);
        } else if (child.getType() == DocumentApp.ElementType.INLINEIMAGE) {
          mergedDoc.appendImage(child.getBlob());
        } else if (child.getType() == DocumentApp.ElementType.PARAGRAPH) {
          mergedDoc.appendParagraph(child);
        } else if (child.getType() == DocumentApp.ElementType.LISTITEM) {
          mergedDoc.appendListItem(child);
        } else if (child.getType() == DocumentApp.ElementType.TABLE) {
          mergedDoc.appendTable(child);
        } else {
          Logger.log("Unknown element type: " + child);
        }
      }
    
  mergedDoc.saveAndClose();
  return mergedDoc;
}

function formatStringCurrency(tempValue) {
  // Description
  return Utilities.formatString("$%.2f", +tempValue);
}