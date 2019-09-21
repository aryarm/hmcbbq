blankSheet = SpreadsheetApp.openByUrl("https://docs.google.com/a/g.hmc.edu/spreadsheets/d/10aYWsGoIRAG2tnaAMkJ75Mz8UcIX2TJMfmIlt0oryYM/edit?usp=sharing");
user_sheet = SpreadsheetApp.getActiveSpreadsheet();

function copyDataOver(previous_list_sheet, previous_admin_sheet, user_list, user_admin) {
  previous_list_data = previous_list_sheet.getRange('B6:D').getValues();
  previous_admin_data = previous_admin_sheet.getRange('B1:B').getValues();
  user_list.getRange('B6:D').setValues(previous_list_data);
  user_admin.getRange('B1:B').setValues(previous_admin_data);
}

function copyBlankSheet() {
  alertUser = false; // determine whether to tell the user they must copy their data over
  // get sheets to copy from
  blankList = blankSheet.getSheets()[0];
  blankAdmin = blankSheet.getSheets()[1];
  // copy the sheets into the user's document
  user_list = blankList.copyTo(user_sheet);
  // check that the name "The List" doesn't already exist before setting it, and rename the other one if it is
  previous_list_sheet = user_sheet.getSheetByName('The List');
  if (previous_list_sheet != null) {
    alertUser = true;
    previous_list_sheet.setName("The List- Old");
  }
  user_list.setName('The List');
  user_admin = blankAdmin.copyTo(user_sheet);
  // check that the name "Admin" doesn't already exist before setting it, and rename the other one if it is
  previous_admin_sheet = user_sheet.getSheetByName('Admin');
  if (previous_admin_sheet != null) {
    previous_admin_sheet.setName("Admin- Old");
  } else {
    alertUser = false;
  }
  user_admin.setName('Admin');
  // store the sheet IDs for later use
  documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperties({
    'list_id' : user_list.getSheetId(),
    'admin_id' : user_admin.getSheetId()
  });
  if (alertUser) {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert("Copy Old Data?","You have a previous copy of The List and the Admin sheet already. You should make sure to copy the data from those sheets over to the ones we just made for you and then delete the old sheets. If you'd like, we can copy the data over for you.\nWould you like us to copy over your old data?", ui.ButtonSet.YES_NO);
    if (response == ui.Button.YES) {
      copyDataOver(previous_list_sheet, previous_admin_sheet, user_list, user_admin);
    }
  }
}

// sets permission on the stuff that only the admin should be able to edit
function setPermissions() {
  documentProperties = PropertiesService.getDocumentProperties();
  user_list_sheet = getSheetByID(user_sheet, documentProperties.getProperty('list_id'));
  // Protect range A1:D5, then remove all other users from the list of editors.
  var protection = user_list_sheet.getRange('A1:D5').protect();
  // Ensure the current user is an editor before removing others. Otherwise, if the user's edit
  // permission comes from a group, the script will throw an exception upon removing the group.
  protection.addEditor(Session.getEffectiveUser());
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
  
  user_admin_sheet = getSheetByID(user_sheet, documentProperties.getProperty('admin_id'));
  // Protect sheet, then remove all other users from the list of editors.
  var protection = user_admin_sheet.protect();
  // Ensure the current user is an editor before removing others. Otherwise, if the user's edit
  // permission comes from a group, the script will throw an exception upon removing the group.
  protection.addEditor(Session.getEffectiveUser());
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
  
  // lastly, let's make the file editable by anybody with the link, so that functions.exportAsExcel() will work
  currFile = DriveApp.getFileById(user_sheet.getId());
  currFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT).setShareableByEditors(false);
}

// returns an integer (0 through 6) day from the string dayOfWeek (ex- input of Sunday returns 0)
function getDay(dayOfWeek){
  if (dayOfWeek == "") { // check for empty values
    return false;
  }
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
    return false;
  }
  return dayOfWeek;
}

// returns an integer number (from 1 to 24) representing an hour from the specified hour string (ex inputs include '12 pm' or '1 am'); returns false on failure
function getHour(hour) {
  if (hour == "") { // check for empty values
    return false;
  }
  hour = hour.split(" ");
  AMorPM = hour[1];
  hour = parseInt(hour[0]);
  if ((hour > 24 || hour < 1) || !(AMorPM == 'am' || AMorPM == 'pm')) { // check for malformed data
    return false;
  }
  // add 12 to any time from 1:00 pm to 12:00 am
  if ((AMorPM == 'pm' && hour != 12) || (hour == 12 && AMorPM == 'am')) {
    hour += 12;
  }
  return hour;
}

// a trigger runs this function each hour; checks which functions need to be run and runs them
function runTriggerFunctions() {
  // get admin sheet
  documentProperties = PropertiesService.getDocumentProperties();
  user_admin_sheet = getSheetByID(user_sheet, documentProperties.getProperty('admin_id'));
  tempLog = 0;
  if (!getConfigData(user_admin_sheet,'disabledApp').getValue()) { // if the app isn't disabled
    // get current day and time
    currDay = (new Date).getDay();
    currHour = (new Date).getHours();
    // correct for 12 am, since getHours() represents it as 0 instead of 24
    if (currHour == 0) {
      currHour = 24;
    }
    // check whether submitForm needs to get run
    if (getDay(getConfigData(user_admin_sheet,'daytoSubmitForm').getValue()) == currDay
        && getHour(getConfigData(user_admin_sheet,'timetoSubmitForm').getValue()) == currHour) { // if the day and time match the current day and time, we need to submit the form
      submitForm();
    }
    if (!getConfigData(user_admin_sheet,'enabledEmail').getValue()) { // if the email is enabled
      // check whether sendReminderEmail needs to get run
      if (getDay(getConfigData(user_admin_sheet,'daytoSendEmail').getValue()) == currDay
          && getHour(getConfigData(user_admin_sheet,'timetoSendEmail').getValue()) == currHour) { // if the day and time match the current day and time, we need to submit the form
        sendReminderEmail();
      }
    }
  }
}

// fills several cells with appropriate formulas and data
function populateCells() {
  documentProperties = PropertiesService.getDocumentProperties();
  user_admin_sheet = getSheetByID(user_sheet, documentProperties.getProperty('admin_id'));
  user_list_sheet = getSheetByID(user_sheet, documentProperties.getProperty('list_id'));
  user_admin_sheet_URL = user_sheet.getUrl()+'#gid='+user_admin_sheet.getSheetId();
  // make and set the default reminder email text
  defaultEmail = "Here's the link to change your subscription:<br><a href='"+user_sheet.getUrl()+"#gid="+user_list_sheet.getSheetId()+"'>BBQ Participant List</a><br>If you want a veggie burger, please write 'vegetarian' or 'Vegetarian' so that the system sends the correct number to the Hoch. Please do it before 11:00 am on Tuesday!<br><br>-The BBQ Team";
  getConfigData(user_admin_sheet, 'reminderEmail_withHTML').setValue(defaultEmail);
  // make and set the formula for the dorm list
  list_dorm = '=IMPORTRANGE("'+user_admin_sheet_URL+'","Admin!'+getConfigData(user_admin_sheet, 'dorm').getA1Notation()+'")';
  getConfigData(user_list_sheet, 'list_dorm').setFormula(list_dorm);
  // make and set the formula for the user's full name
  list_userName = '=IMPORTRANGE("'+user_admin_sheet_URL+'","Admin!'+getConfigData(user_admin_sheet, 'user_fname').getA1Notation()+'")&" "&'+
    'IMPORTRANGE("'+user_admin_sheet_URL+'","Admin!'+getConfigData(user_admin_sheet, 'user_lname').getA1Notation()+'")';
  getConfigData(user_list_sheet, 'list_userName').setFormula(list_userName);
  // make and set the formula for the day of the BBQ
  list_dayOfBBQ = '=IMPORTRANGE("'+user_admin_sheet_URL+'","Admin!'+getConfigData(user_admin_sheet, 'dayOfBBQ').getA1Notation()+'")';
  getConfigData(user_list_sheet, 'list_dayOfBBQ').setFormula(list_dayOfBBQ);
}

function setUp(e) {
  copyBlankSheet();
  setPermissions();
  // make an hourly trigger
  ScriptApp.newTrigger('runTriggerFunctions').timeBased().everyHours(1).create();
  populateCells();
}

function onInstall(e) {
  setUp(e);
  storeEmail();
}
