function emailMeTheEmailList() {
  var scriptProperties = PropertiesService.getScriptProperties();
  email_list = scriptProperties.getProperty('email_list');
  MailApp.sendEmail({to: Session.getEffectiveUser().getEmail(),
                     subject: 'HMC BBQ Script Email List',
                     htmlBody: "Here is a list of those who have installed the script and the dates they installed it:<br>"+email_list});
}

function logTheEmailList() {
  var scriptProperties = PropertiesService.getScriptProperties();
  email_list = scriptProperties.getProperty('email_list');
  Logger.log(email_list);
}

function storeEmail() {
  var scriptProperties = PropertiesService.getScriptProperties();
  email_list = scriptProperties.getProperty('email_list');
  new_email = Session.getActiveUser().getEmail()+' (<'+(new Date()).toLocaleString()+'>)';
  if (email_list === null){
    scriptProperties.setProperty('email_list', new_email);
  } else {
    email_list += '; '+new_email;
    scriptProperties.setProperty('email_list', email_list);
  }
}

// be careful with this! it will delete all emails in the email list!
function clearEmailList() {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty('email_list');
}
