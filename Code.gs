// GLOBAL VARIABLES
const SPREADSHEET_ID = '1L3bJ52H8QYTxQEK-FitIFlcCdob-FNPVANZCNrYRrkQ'; // Replace with your spreadsheet ID

const CACHE_PROP = CacheService.getPublicCache();
const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
const SETTINGS_SHEET = "_Settings";
const CACHE_SETTINGS = false;
const SETTINGS_CACHE_TTL = 900;
const cache = JSONCacheService();
const SETTINGS = getSettings();
const subSheet = ss.getSheetByName("Submissions");
const approvedSheet = ss.getSheetByName("Approved");
const compSheet = ss.getSheetByName("Completed");
const rejectSheet = ss.getSheetByName("Rejected");
const paidSheet = ss.getSheetByName("Paid");

//**************************************************************************************************************************************************

function doGet(e) {
  let idVal = e.parameter.idNum;
  let step = e.parameter.st;
  let template = null; 
  
  if (idVal) {
    let theSheet = subSheet;
    switch(step) {
      case "1": 
        template = HtmlService.createTemplateFromFile('building.html');
        break;
      case "2":
        template = HtmlService.createTemplateFromFile('district.html');
        break;
      case "3":
        template = HtmlService.createTemplateFromFile('postform.html');
        theSheet = approvedSheet;
        break;
      case "4":
        template = HtmlService.createTemplateFromFile('finalapprove.html');
        theSheet = approvedSheet;
        break;
      case "5":
        template = HtmlService.createTemplateFromFile('approvePayment.html');
        theSheet = compSheet;
        break;
      case "6":
        template = HtmlService.createTemplateFromFile('okToPay.html');
        theSheet = compSheet;
        break;
      default:
        template = HtmlService.createTemplateFromFile('TempDown.html');
    }
    let rowNum = findRow(theSheet, idVal);
    Logger.log("RN: "+rowNum);
    if (rowNum == -1) {
      template = HtmlService.createTemplateFromFile("done.html");
    } else {
      let last = theSheet.getLastColumn();
      let header = theSheet.getRange(1, 1, 1, last).getValues()[0];
      let rowData = theSheet.getRange(rowNum,1, 1, last).getValues()[0];
      let theData = getJsonFromData(header, rowData);
      //Logger.log(theData);
      template.info = theData;
    }
  }
  else {
    template = HtmlService.createTemplateFromFile('index.html');
    //template = HtmlService.createTemplateFromFile('TempDown.html'); // uncomment this line (comment out index.html) for initial submission site down
    template.perMile = SETTINGS.PER_MILE;
  }
  // template = HtmlService.createTemplateFromFile('TempDown.html'); // Uncomment out when making changes to show ENTIRE site down
  let html = template.evaluate();
  return HtmlService.createHtmlOutput(html).setTitle("OFCS Reimbursement Form");
}
//**************************************************************************************************************************************************
function createFolder(folderName) {
  const parentFolderId = SETTINGS.ATTACH_FOLDER_ID;
  try {
    let parentFolder = DriveApp.getFolderById(parentFolderId);
    let folders = parentFolder.getFoldersByName(folderName);
    let folder;
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = parentFolder.createFolder(folderName);
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }
    
    return {
      'folderId' : folder.getId()
    }
  } catch (e) {
    return {
      'error' : e.toString()
    }
  }
}

function uploadFile(base64Data, fileName, folderId) {
  try {
    let foldPath = "";
    if (folderId) {
      foldPath = "https://drive.google.com/drive/folders/"+folderId;
    }
    CACHE_PROP.put('foldPath', foldPath, 3600);
  
    let splitBase = base64Data.split(','), type = splitBase[0].split(';')[0]
    .replace('data:', '');
    let byteCharacters = Utilities.base64Decode(splitBase[1]);
    let ss = Utilities.newBlob(byteCharacters, type);
    ss.setName(fileName);
    let folder = DriveApp.getFolderById(folderId);
    let files = folder.getFilesByName(fileName);
    let file;
    while (files.hasNext()) {
      // delete existing files with the same name.
      file = files.next();
      folder.removeFile(file);
    }
    file = folder.createFile(ss);
    
    return {
      'folderId' : folderId,
      'fileName' : file.getName()
    };
  } catch (e) {
    return {
      'error' : e.toString()
    };
  }
}
//**************************************************************************************************************************************************
function submit (data, statusVal) {
  let finalDoc = null;
  let json = JSON.parse(data);
  json.status = statusVal;
  let sheet = subSheet;
  if (statusVal == "Pending Final" || statusVal == "Approved for Pay" || statusVal == "Return with Comments") {
    sheet = approvedSheet;
  } else if (statusVal == "OK to Pay" || statusVal == "Payment Issued" || statusVal == "Return with Comments 2" || statusVal == "Return with Comments 3") {
    sheet = compSheet;
  }
  let rowNum = -1;
  if (statusVal == "Pending") {
    json.submissionID = +new Date();
    let today_date = new Date();
    json.timestamp = today_date;
    let submitDateString = today_date.toDateString();//today_Date_Str = "'" + ((today_date.getMonth() < 9) ? "0" : "") + String(today_date.getMonth() + 1)+ "-" + ((today_date.getDate() < 10) ? "0" : "") + String(today_date.getDate()) + "-" +today_date.getFullYear();
    json.submitDate = submitDateString;
    json.buildSignature = getBuildingInfo(json.building).adminName;
  } else {
    json.buildSignature = getBuildingInfo(json.building).adminName;
    json.districtSignature = SETTINGS.DISTRICT_ADMIN;
    rowNum = findRow(sheet, json.submissionID);
  }
  if (statusVal == "Pending Final") {
    json.uploadLink = CACHE_PROP.get('foldPath');
  }
  writeJSONtoSheet(json, sheet, rowNum);
  if (statusVal == "District Approved") {
    Logger.log("Ready to move");
    finalDoc = doMerge(rowNum, json.submissionID, json.lastName, SETTINGS.DESTINATION_FOLDER_ID, SETTINGS.FORM_TEMPLATE_ID, SPREADSHEET_ID, "Submissions");
    moveCompleted(rowNum,subSheet,approvedSheet);
  } else if (statusVal == "Building Rejected" || statusVal == "District Rejected") {
    moveCompleted(rowNum,subSheet,rejectSheet);
  } else if (statusVal == "Approved for Pay") {
    moveCompleted(rowNum, approvedSheet, compSheet);
  } else if (statusVal == "Payment Issued") { 
    finalDoc = doMerge(rowNum, json.submissionID, json.lastName, SETTINGS.DESTINATION_FOLDER_FINAL_ID, SETTINGS.FORM_TEMPLATE_FINAL_ID, SPREADSHEET_ID, "Completed");
    moveCompleted(rowNum, compSheet, paidSheet); 
  }
  if (statusVal == "Return with Comments 2") {
    moveCompleted(rowNum, compSheet, approvedSheet);
  }
  sendEmail (json, statusVal, finalDoc);
  
}
//**************************************************************************************************************************************************
function sendEmail (data, status, attachDoc) {
  Logger.log(status);
  let htmlBody = "";
  let recipients = "jvanarnhem@ofcs.net";
  switch(status) {
      case "Pending":
          recipients = getBuildingInfo(data.building).adminEmail;
          htmlBody = "<h2>A Professional Reimbursement Form was submitted. </h2>";
          htmlBody += '<p><strong>Click <a href="' + ScriptApp.getService().getUrl()
              + '?idNum=' + data.submissionID
              + '&st=1">on this link</a> to see full details of report and to approve/deny proposal.</strong></p>';
          htmlBody += emailSummary(data);
          break;
      case "Building Approved":
          recipients = SETTINGS.DISTRICT_EMAIL;
          htmlBody = "<h2>A Professional Reimbursement Form was submitted. </h2>";
          htmlBody += '<p><strong>Click <a href="' + ScriptApp.getService().getUrl()
             + '?idNum=' + data.submissionID
             + '&st=2">on this link</a> to see full details of report and to approve/deny proposal.</strong></p>';
          htmlBody += emailSummary(data);
          htmlBody += "<p>Admin Name: " + data.buildSignature + "<br>";
          htmlBody += "Administrator Comments:" + data.adminComments + "</p>";
          break;
      case "District Approved":
          switch(data.route) {
              case "District":
                recipients = SETTINGS.DISTRICT_NOTICE + "," + data.email + "," + getBuildingInfo(data.building).secEmail;
                break;
              case "Grant/SPED":
                recipients = SETTINGS.GRANT_NOTICE + "," + data.email + "," + getBuildingInfo(data.building).secEmail;
                break;
              default:
                recipients = getBuildingInfo(data.building).adminEmail + "," + data.email + "," + getBuildingInfo(data.building).secEmail;
          }
          htmlBody = "<p><strong>IMPORTANT: Keep this email and use the link below to submit final costs (and receipts) for reimbursement.</strong></p>";
          htmlBody += '<p><strong><a href="' + ScriptApp.getService().getUrl() + '?idNum=' + data.submissionID
               + '&st=3" target=_blank>Link to submit final costs and receipts.</a></strong></p>';
          htmlBody += "<p><span style='color:red'><strong>Applicant:</strong> Your submission has been pre-approved.  Please get the purchase order number from the secretary. The PO number is needed for the post submission reimbursement request.</p>";
          htmlBody += "<p><strong>Secretary:</strong> Please draft a purchase order for this pre-approved professional development event.  Please send the PO number to the attendee. The PO number is needed for the post submission reimbursement request.</p>";
          htmlBody += "<p>You will not receive reimbursement until the final form is submitted and approved.</p>";
          htmlBody += "<p>If you have any questions, please contact James Tatman</span></p>";
          htmlBody += emailSummary(data);
          htmlBody += "<p>Admin Name: " + data.buildSignature + "<br>";
          htmlBody += "Administrator Comments:" + data.adminComments + "<br>";
          htmlBody += "<p>District Admin: " + SETTINGS.DISTRICT_ADMIN + "<br>";
          htmlBody += "District Admin Comments:" + data.districtComments + "</p>";
          break;
      case "Building Rejected":
          recipients = data.email;
          htmlBody = "<p>The following Professional Reimbursement Form was rejected by your building administrator:</p>";
          htmlBody += emailSummary(data);
          htmlBody += "<p>Admin Name: " + data.buildSignature + "<br>";
          htmlBody += "Administrator Comments:" + data.adminComments + "</p>";
          htmlBody += '<p>Unfortunately you will need to complete the application again if you wish to reapply.</p>';
          htmlBody += '<p>Click <a href="' + ScriptApp.getService().getUrl() + '">here</a> to resubmit your application.</p>';
          break;
       case "District Rejected":
          recipients = data.email;
          htmlBody = "<p>The following Professional Reimbursement Form Form was rejected by your district administrator:</p>";
          htmlBody += emailSummary(data);
          htmlBody += "<p>Admin Name: " + data.buildSignature + "<br>";
          htmlBody += "Administrator Comments:" + data.adminComments + "<br>";
          htmlBody += "<p>District Admin: " + SETTINGS.DISTRICT_ADMIN + "<br>";
          htmlBody += "District Admin Comments:" + data.districtComments + "</p>";
          htmlBody += '<p>Unfortunately you will need to complete the application again if you wish to reapply.</p>';
          htmlBody += '<p>Click <a href="' + ScriptApp.getService().getUrl() + '">here</a> to resubmit your application.</p>';
          break;
      case "Pending Final":
          recipients = getBuildingInfo(data.building).adminEmail;
          htmlBody = "<p>The following Professional Reimbursement Form was submitted for final approval for payment:</p>";
          htmlBody += '<p><strong>Click <a href="' + ScriptApp.getService().getUrl()
             + '?idNum=' + data.submissionID
             + '&st=4">on this link</a> to see full details of report and to approve/deny proposal.</strong></p>';
          htmlBody += emailSummary(data);
          htmlBody += "<p>Admin Name: " + data.buildSignature + "<br>";
          htmlBody += "Administrator Comments:" + data.adminComments + "<br>";
          htmlBody += "<p>District Admin: " + SETTINGS.DISTRICT_ADMIN + "<br>";
          htmlBody += "District Admin Comments:" + data.districtComments + "</p>";
          break;
      case "Approved for Pay":
          recipients = SETTINGS.BOE_SEC;
          htmlBody = "<p>The following Professional Reimbursement Form was approved for payment:</p>";
          htmlBody += '<p><strong>Click <a href="' + ScriptApp.getService().getUrl()
             + '?idNum=' + data.submissionID
             + '&st=5">on this link</a> to see full details of report and to approve/deny proposal.</strong></p>';
          htmlBody += emailSummary(data);
          break;
      case "OK to Pay":
          recipients = SETTINGS.ACCOUNTS_PAY;
          htmlBody = "<p>The following Professional Reimbursement Form was approved for payment:</p>";
          htmlBody += '<p><strong>Click <a href="' + ScriptApp.getService().getUrl()
             + '?idNum=' + data.submissionID
             + '&st=6">on this link</a> to review the full details of report and to issue payment.</strong></p>';
          htmlBody += emailSummary(data);
          break;
      case "Payment Issued":
          recipients = data.email;
          htmlBody = "<p>Your reimbursement has been processed for payment.  You will receive a second alert regarding electronic reimbursement payment which will include the payment date and amount.</p>";
          htmlBody += emailSummary(data);
          break;
      case "Return with Comments":
          recipients = data.email;
          htmlBody = "<p><strong>IMPORTANT: Your application was returned and action must be taken before payment is issued.</strong></p>";
          htmlBody += "<p>Return comments: " + data.returnComments1 + "</p>";
          htmlBody += '<p><strong><a href="' + ScriptApp.getService().getUrl() + '?idNum=' + data.submissionID
               + '&st=3" target=_blank>Link to submit final costs and receipts.</a></strong></p>';
          htmlBody += "<p>You will not receive reimbursement until the final form is re-submitted and approved.</p>";
          htmlBody += "<p>If you have any questions, please contact James Tatman</p>";
          htmlBody += emailSummary(data);
          htmlBody += "<p>Admin Name: " + data.buildSignature + "<br>";
          htmlBody += "Administrator Comments:" + data.adminComments + "<br>";
          htmlBody += "<p>District Admin: " + SETTINGS.DISTRICT_ADMIN + "<br>";
          htmlBody += "District Admin Comments:" + data.districtComments + "</p>";
      
          break;
      case "Return with Comments 2":
          recipients = data.email;
          htmlBody = "<p><strong>IMPORTANT: Your application was returned and action must be taken before payment is issued.</strong></p>";
          htmlBody += "<p>Return comments: " + data.returnComments2 + "</p>";
          htmlBody += '<p><strong><a href="' + ScriptApp.getService().getUrl() + '?idNum=' + data.submissionID
               + '&st=3" target=_blank>Link to submit final costs and receipts.</a></strong></p>';
          htmlBody += "<p>You will not receive reimbursement until the final form is re-submitted and approved.</p>";
          htmlBody += "<p>If you have any questions, please contact James Tatman</p>";
          htmlBody += emailSummary(data);
          htmlBody += "<p>Admin Name: " + data.buildSignature + "<br>";
          htmlBody += "Administrator Comments:" + data.adminComments + "<br>";
          htmlBody += "<p>District Admin: " + SETTINGS.DISTRICT_ADMIN + "<br>";
          htmlBody += "District Admin Comments:" + data.districtComments + "</p>";
      
          break;
       case "Return with Comments 3":
          recipients = data.email;
          htmlBody = "<p><strong>IMPORTANT: Your application was returned and action must be taken before payment is issued.</strong></p>";
          htmlBody += "<p>Return comments: " + data.returnComments3 + "</p>";
          htmlBody += '<p><strong><a href="' + ScriptApp.getService().getUrl() + '?idNum=' + data.submissionID
               + '&st=3" target=_blank>Link to submit final costs and receipts.</a></strong></p>';
          htmlBody += "<p>You will not receive reimbursement until the final form is re-submitted and approved.</p>";
          htmlBody += "<p>If you have any questions, please contact James Tatman</p>";
          htmlBody += emailSummary(data);
          htmlBody += "<p>Admin Name: " + data.buildSignature + "<br>";
          htmlBody += "Administrator Comments:" + data.adminComments + "<br>";
          htmlBody += "<p>District Admin: " + SETTINGS.DISTRICT_ADMIN + "<br>";
          htmlBody += "District Admin Comments:" + data.districtComments + "</p>";
      
          break;
      default:
          htmlBody = "<h2>Default reached for send email -- NOT GOOD</h2>";
    }
    
    if(attachDoc) {
      // CHANGE EMAIL ADDRESS HERE to "adminemail"  
      MailApp.sendEmail({
        to: recipients,
        subject: "Professional Reimbursement Submission: "+data.lastName + " #" + data.submissionID,
        htmlBody: htmlBody,
        attachments: [attachDoc.getAs(MimeType.PDF)]
      });
    } else {
        // CHANGE EMAIL ADDRESS HERE to "adminemail"  
      MailApp.sendEmail({
        to: recipients,
        subject: "Professional Reimbursement Submission: "+data.lastName + " #" + data.submissionID,
        htmlBody: htmlBody,
      });
    }


}
//**************************************************************************************************************************************************
function emailSummary(data) {
    let theCost = data.actTotalCost ? data.actTotalCost : data.totalCost;
    htmlBody = "<p>&nbsp;</p>";
    htmlBody += "<h4>Summary: </h4>";
    htmlBody += "<p>Submitter: " + data.firstName + " " + data.lastName + "<br>";
    htmlBody += "Date submitted: " + data.submitDate + "<br>";
    htmlBody += "Meeting information: " + data.meetingInfo  + "<br>";
    htmlBody += "Meeting Dates: " + data.startDate + " to " + data.endDate + "<br>";
    htmlBody += "Driving Info: " + data.drivingInfo + "<br>";
    htmlBody += "Total Cost: " + theCost + "</p>";
    return htmlBody;
}
//**************************************************************************************************************************************************
function getBuildingInfo (theBuilding) {
    let build = {};
    switch(theBuilding) {
      case "HS":
        build.adminName = SETTINGS.HS_ADMIN;
        build.adminEmail = SETTINGS.HS_EMAIL;
        build.secEmail = SETTINGS.HS_SEC;
        break;
      case "MS":
        build.adminName = SETTINGS.MS_ADMIN;
        build.adminEmail = SETTINGS.MS_EMAIL;
        build.secEmail = SETTINGS.MS_SEC;
        break;
      case "OFIS":
        build.adminName = SETTINGS.IS_ADMIN;
        build.adminEmail = SETTINGS.IS_EMAIL;
        build.secEmail = SETTINGS.IS_SEC;
        break;
      case "FL":
        build.adminName = SETTINGS.FL_ADMIN;
        build.adminEmail = SETTINGS.FL_EMAIL;
        build.secEmail = SETTINGS.FL_SEC;
        break;
      case "ECC":
        build.adminName = SETTINGS.ECC_ADMIN;
        build.adminEmail = SETTINGS.ECC_EMAIL;
        build.secEmail = SETTINGS.ECC_SEC;
        break;
      case "BUS":
        build.adminName = SETTINGS.BUS_ADMIN;
        build.adminEmail = SETTINGS.BUS_EMAIL;
        build.secEmail = SETTINGS.BUS_SEC;
        break;
      case "BOE":
        build.adminName = SETTINGS.BOE_ADMIN;
        build.adminEmail = SETTINGS.BOE_EMAIL;
        build.secEmail = SETTINGS.BOE_SEC;
        break;
      default:
        console.log("Something went horribly wrong...");
  }
  return build;
}
//**************************************************************************************************************************************************
function moveCompleted(rowNum, source, target) {
  
  let targetRange = target.getRange(target.getLastRow() + 1, 1);
  let dataToMove = source.getRange(rowNum, 1, 1, source.getLastColumn()).moveTo(targetRange);
  
  source.deleteRow(rowNum);
}        
//**************************************************************************************************************************************************
function getSettings() { 
  console.log(CACHE_SETTINGS);

  //if(CACHE_SETTINGS) {
  //  settings = cache.get("_settings");
  //}
  
  if(settings == undefined) {
    let sheet = ss.getSheetByName(SETTINGS_SHEET);
    let values = sheet.getDataRange().getValues();
  
    var settings = {};
    for (let i = 1; i < values.length; i++) {
      let row = values[i];
      settings[row[0]] = row[1];
    }
    
    cache.put("_settings", settings, SETTINGS_CACHE_TTL);
  }
  console.log(settings);
  return settings;
}
function JSONCacheService() {
  let _cache = CacheService.getPublicCache();
  let _key_prefix = "_json#";
  
  let get = function(k) {
    let payload = _cache.get(_key_prefix+k);
    if(payload !== undefined) {
      JSON.parse(payload);
    }
    return payload
  }
  
  let put = function(k, d, t) {
    _cache.put(_key_prefix+k, JSON.stringify(d), t);
  }
  
  return {
    'get': get,
    'put': put
  }
}

//**************************************************************************************************************************************************
// Written by Amit Agarwal www.ctrlq.org

function writeJSONtoSheet(json, sheet, rowNum) {
  
  let keys = Object.keys(json).sort();
  let last = sheet.getLastColumn();
  let header = sheet.getRange(1, 1, 1, last).getValues()[0];
  let newCols = [];

  for (let k = 0; k < keys.length; k++) {
    if (header.indexOf(keys[k]) === -1) {
      newCols.push(keys[k]);
    }
  }

  if (newCols.length > 0) {
    sheet.insertColumnsAfter(last, newCols.length);
    sheet.getRange(1, last + 1, 1, newCols.length).setValues([newCols]);
    header = header.concat(newCols);
  }

  let row = [];

  for (let h = 0; h < header.length; h++) {
    row.push(header[h] in json ? json[header[h]] : "");
    if (header[h]=="status") {
      Logger.log("h = "+h);
      conditionalFormat(sheet, h);
    }
  }
  
  if (rowNum == -1) {
    sheet.appendRow(row);
  } else {
    //Logger.log(rowNum);
    //Logger.log(row);
    sheet.getRange(rowNum,1, 1, row.length).setValues([row]);
  } 
}

function conditionalFormat(sheet, statCol) {
  let range = sheet.getRange(1,statCol+1,sheet.getLastRow()+1,1);
  Logger.log(range.getValues());
  let rule1 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Pending")
    .setBackground("#33FFF3")
    .setRanges([range])
    .build();
  let rule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Building Approved")
    .setBackground("#ffff66")
    .setRanges([range])
    .build();
  let rule3 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("District Approved")
    .setBackground("#66ff66")
    .setRanges([range])
    .build();
  let rule4 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Building Rejected")
    .setBackground("#ffc266")
    .setRanges([range])
    .build();
  let rule5 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("District Rejected")
    .setBackground("#ff9999")
    .setRanges([range])
    .build();
  let rule6 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Pending Final")
    .setBackground("#00ffff")
    .setRanges([range])
    .build();
  let rule7 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Approved for Pay")
    .setBackground("#F2FC2D")
    .setRanges([range])
    .build();
  let rule8 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("OK to Pay")
    .setBackground("#50CA1F")
    .setRanges([range])
    .build();
  let rule9 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Payment Issued")
    .setBackground("#FC2DF2")
    .setRanges([range])
    .build();
  let rule10 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Return with Comments")
    .setBackground("#C285E9")
    .setRanges([range])
    .build();
  let rule11 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Return with Comments 2")
    .setBackground("#C285E9")
    .setRanges([range])
    .build();
  let rule12 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Return with Comments 3")
    .setBackground("#C285E9")
    .setRanges([range])
    .build();
  let rules = sheet.getConditionalFormatRules();
  rules.push(rule1);
  rules.push(rule2);
  rules.push(rule3);
  rules.push(rule4);
  rules.push(rule5);
  rules.push(rule6);
  rules.push(rule7);
  rules.push(rule8);
  rules.push(rule9);
  rules.push(rule10);
  rules.push(rule11);
  rules.push(rule12);
  sheet.setConditionalFormatRules(rules);
}

function getJsonFromData(headers, rowData)
{

  let obj = {};
  let cols = headers.length;
 
  for (let col = 0; col < cols; col++) 
  {
    // fill object with new values
    obj[headers[col]] = rowData[col];    
  }
  return obj;  
}

function findRow(sheet, subID)
{  
    let columnValues = sheet.getRange(2, 1, sheet.getLastRow()).getValues(); //1st is header row
    let searchResult = columnValues.findIndex(subID); //Row Index - 2
    if (searchResult == -1) {
        return -1;
    }
    return searchResult+2;
}

Array.prototype.findIndex = function(search){
  if(search == "") return false;
  for (let i=0; i<this.length; i++)
    if (this[i] == search) return i;

  return -1;
}

//**************************************************************************************************************************************************
function testStuff() {
  Logger.log(SETTINGS.MS_EMAIL);
}
