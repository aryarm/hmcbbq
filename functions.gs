theFile = SpreadsheetApp.getActiveSpreadsheet();

function getSheetByID(spreadSheet, sheet_id) {
  sheets = spreadSheet.getSheets();
  for (i=0;i<sheets.length;i++) {
    if (sheets[i].getSheetId()==sheet_id) {
      return sheets[i];
    }
  }
}

// abstraction for config data; specifies cell locations for each piece of data
function getConfigData(adminFile, dataType) {
  // create a map that associates a cell reference with each cell that contains it
  var dataMap = new Object();
  dataMap['disabledApp'] = 'B16';
  dataMap['user_fname'] = 'B1';
  dataMap['user_lname'] = 'B2';
  dataMap['user_cellPhone'] = 'B3';
  dataMap['dorm'] = 'B4';
  dataMap['dayOfBBQ'] = 'B5';
  dataMap['enabledEmail'] = 'B6';
  dataMap['email_recipients'] = 'B7';
  dataMap['email_subject'] = 'B8';
  dataMap['daytoSendEmail'] = 'B9';
  dataMap['timetoSendEmail'] = 'B10';
  dataMap['reminderEmail_withHTML'] = 'B11';
  dataMap['bbqRequestFormContent'] = 'B12';
  dataMap['daytoSubmitForm'] = 'B13';
  dataMap['timetoSubmitForm'] = 'B14';
  dataMap['list_dorm'] = 'B2';
  dataMap['list_userName'] = 'B4';
  dataMap['list_dayOfBBQ'] = 'C4';
  if (!(dataType in dataMap)) {
    return false;
  }
  return adminFile.getRange(dataMap[dataType]);
}

stuff_is_gonna_work = PropertiesService.getDocumentProperties() !== null;
if (stuff_is_gonna_work) {
  adminFile = getSheetByID(theFile, PropertiesService.getDocumentProperties().getProperty('admin_id'));
  if (adminFile != null) {
    listFile = getSheetByID(theFile, PropertiesService.getDocumentProperties().getProperty('list_id'));
    if (listFile == null) {
      throw new Error("Error: couldn't find the bbq participant list");
    }
    dayOfBBQ = getConfigData(adminFile,'dayOfBBQ').getValue();
    
    //  formURL = "note to developer: add your own website here for testing"
    formURL = "https://hmc.formstack.com/forms/index.php";
  }
}

function sendReminderEmail() {
  if (!getConfigData(adminFile,'disabledApp').getValue()
      && !getConfigData(adminFile,'enabledEmail').getValue()) {
    updateSpreadSheet();
    MailApp.sendEmail({to: getConfigData(adminFile, 'email_recipients').getValue(),
                       subject: getConfigData(adminFile, 'email_subject').getValue(),
                       htmlBody: getConfigData(adminFile,'reminderEmail_withHTML').getValue(),
                       name: getConfigData(adminFile,'user_fname').getValue()+" "+getConfigData(adminFile,'user_lname').getValue()});
  }
}

// given a capatalized string representing a day of the week, return the date in a nice format (as a string)
function getDateOfWeek(dayOfWeek) {
  var weekday = new Array(7);
  weekday[0] = "Sunday";
  weekday[1] = "Monday";
  weekday[2] = "Tuesday";
  weekday[3] = "Wednesday";
  weekday[4] = "Thursday";
  weekday[5] = "Friday";
  weekday[6] = "Saturday";
  dayOfWeek = weekday.indexOf(dayOfWeek);
  if (dayOfWeek < 0){
    throw new Error("invalid input for the dayOfBBQ variable");
  }
  
  var today = new Date();
  today.setDate(today.getDate() + (dayOfWeek - 1 - today.getDay() + 7) % 7 + 1);
  var dd = today.getDate();
  var mm = today.getMonth()+1; //January is 0!
  var yyyy = today.getFullYear();
  if(dd<10) {
    dd='0'+dd
  }
  if(mm<10) {
    mm='0'+mm
  } 
  today = mm+'/'+dd+'/'+yyyy;
  
  return today;
}

// deletes all empty rows and makes non-empty rows look good
function removeEmptySpaces() {
  numNonEmptyCells = listFile.getRange("C6:C").getValues().filter(String).length;
  for (i=6; i<numNonEmptyCells+6; i++){
    if (listFile.getRange("C"+i).isBlank()){ // pick out blank cells
// move it
      listFile.deleteRow(i); // delete the blank row
      lastRow = listFile.getRange(listFile.getLastRow()+":"+listFile.getLastRow());
      listFile.appendRow([""]); // add a new empty row at the bottom
      lastRow.copyFormatToRange(listFile.getSheetId(), 1, 1, listFile.getLastRow(), listFile.getLastRow()); // copy the correct formatting to the new row
      i--;
    }
  }
}

function updateSpreadSheet() {
  newDate = getDateOfWeek(dayOfBBQ);
  if (listFile.getRange('C2').getValue() != newDate){
    listFile.getRange('C2').setValue(newDate);
  }
  removeEmptySpaces();
}

function exportAsExcel() {
  var url = "https://docs.google.com/feeds/download/spreadsheets/Export?key="+theFile.getId()+"&gid="+listFile.getSheetId()+"&exportFormat=xlsx";

  var options = {
    headers: {
      // the authorization is not needed; it's added automagically
//      Authorization:"Bearer "+ScriptApp.getOAuthToken()
    },
    muteHttpExceptions : true        /// Get failure results
  };

  var response = UrlFetchApp.fetch(url, options);
  var status = response.getResponseCode();
  var result = response.getContentText();
  if (status != 200) {
    // Get additional error message info, depending on format
    if (result.toUpperCase().indexOf("<HTML") !== -1) {
      var message = result;
    }
    else if (result.indexOf('errors') != -1) {
      message = JSON.parse(result).error.message;
    }
    throw new Error('Error (' + status + ") " + message );
  }

  var doc = response.getBlob();
  return doc;
}

function processParticipantList() {
  updateSpreadSheet();
  excelFile = exportAsExcel();
  return excelFile;
}

// returns number of vegetarians
function numVegetarians() {
  return listFile.getRange("D:D").getValues().filter(function(cell){
    cell = JSON.stringify(cell); // convert from object to string so that indexOf can properly work its magic
    return cell.indexOf('Vegetarian')>-1 || cell.indexOf('vegetarian')>-1;
  }).length;
}

function generatePOSTRequest() {
  bbqDate = Utilities.formatDate(new Date(getDateOfWeek(dayOfBBQ)), 'GMT - 8 : 00', 'MMM d yyyy').split(' ');
  var excelBlob = processParticipantList();
  excelBlob.setName('BBQ Participant List.xlsx');
  excelBlob.setContentTypeFromExtension();
  var formData = {
	'_submit' : String(1),
	'field29720171' : getConfigData(adminFile,'dorm').getValue(),
	'field29720189D' : String(bbqDate[1]),
	'field29720189Format' : 'MDY',
	'field29720189M' : bbqDate[0],
	'field29720189Y' : String(parseInt(bbqDate[2])),
	'field29720229-first' : getConfigData(adminFile,'user_fname').getValue(),
	'field29720229-last' : getConfigData(adminFile,'user_lname').getValue(),
	'field29720244' : dayOfBBQ,
	'field29720261' : String(numVegetarians()),
	'field29720276' : getConfigData(adminFile,'bbqRequestFormContent').getValue(),
	'field29720287' : excelBlob,
	'field29896104' : Session.getEffectiveUser().getEmail(),
	'field29896114' : String(getConfigData(adminFile,'user_cellPhone').getValue()),
	'form' : String(1912025),
	'referrer_type' : 'link',
	'style_version' : String(3),
	'viewkey' : '5cZgEY4w8w',
	'viewparam' : String(90931)
 };
 if (!(dayOfBBQ == 'Friday' || dayOfBBQ == 'Saturday' || dayOfBBQ == 'Sunday')){
   formData['field29720244_other'] = dayOfBBQ;
 }
 return formData;
}

function myStringify(jsonObject) {
  stringified = '';
  formData_mapping = {
    'field29720171' : 'Dorm',
    'field29720189D' : 'Day of BBQ',
    'field29720189M' : 'Month of BBQ',
    'field29720229-first' : 'First Name',
    'field29720229-last' : 'Last Name',
    'field29720244' : 'Day of Week of BBQ',
    'field29720261' : 'Number of Vegetarians',
    'field29720276' : 'Message to Dining Hall',
    'field29896104' : 'Email',
    'field29896114' : 'Cell Phone'
  };
  for(var prop in jsonObject) {
    if (prop in formData_mapping){
      stringified += "<i>"+formData_mapping[prop]+"</i> : "+jsonObject[prop]+"<br>";
    }
  }
  return stringified;
}

function submitForm() {
  if (!getConfigData(adminFile,'disabledApp').getValue()) {
    var requestBody = generatePOSTRequest();
    var options = {
      method: "post",
      payload: requestBody,
      muteHttpExceptions: true
    };
    var request = UrlFetchApp.fetch(formURL, options);
    Logger.log(request.getContentText());
    
    
    // send the admin an email notifying them that it was sent
    MailApp.sendEmail({to: Session.getEffectiveUser().getEmail(),
                       subject: "BBQ Form Submitted!",
                       htmlBody: "<b>Here's the info that was submitted:</b><br>"+myStringify(requestBody),
                       name: "BBQ Script",
                       attachments: [requestBody['field29720287']]});
  }
}
