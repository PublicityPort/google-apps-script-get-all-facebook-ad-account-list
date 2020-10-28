//Enter spreadsheet URL where you have your long lived token
const TOKEN_SPREADSHEET_URL= '';

//Create or set tab name as 'token'
const TOKEN_TAB_NAME = 'token';

//Enter spreadsheet URL where you want to store all account informations
const FB_ACC_SS_URL = '';

//Create or set tab name as 'Ad Accounts Information'
const FB_ACC_TAB_NAME = 'Ad Accounts Information';

var newAdAcc = '';
var updateAdAcc = '';

//var facebookUrl = '';

function updateAccountsList() {
  const ss = SpreadsheetApp.openByUrl(TOKEN_SPREADSHEET_URL);
  const sheet = ss.getSheetByName(TOKEN_TAB_NAME);
  
  // user access token linked to a Facebook app
  const TOKEN = sheet.getRange(1, 1).getValue();
  
  PropertiesService.getScriptProperties().setProperty('longLivedToken', TOKEN);
  var facebookUrl = `https://graph.facebook.com/v7.0/me/adaccounts?access_token=${TOKEN}`;
  do{
    facebookUrl = updateAccountsListToSS(facebookUrl, TOKEN);
  }while(facebookUrl != null);
  sendEmails();
}


function updateAccountsListToSS(facebookUrl, TOKEN){
  
  const results = fetchFBJSONResponse(facebookUrl);
  
  searchAndUpdateAccountInfo(results.data, TOKEN);
  
  //Logger.log(results.data);
  //Logger.log(results.paging.next);
  //Logger.log(acc_data.length)
  
  return results.paging.next;
}

function searchAndUpdateAccountInfo(acc_data, TOKEN){
  const ss = SpreadsheetApp.openByUrl(FB_ACC_SS_URL);
  const sheet = ss.getSheetByName(FB_ACC_TAB_NAME);
  
  const adAccountIdCol = sheet.createTextFinder('Ad Account ID').findNext().getColumn();
  const statusCol = sheet.createTextFinder('Status').findNext().getColumn();
  const associatedSSUrlCol = sheet.createTextFinder('SpreadSheet_URL').findNext().getColumn();
  const nameCol = sheet.createTextFinder('Name').findNext().getColumn();  
  
  for(i=0;i<acc_data.length;i++){
    //Logger.log(acc_data[i].id);
    var AD_ACC_ID = acc_data[i].id;
    var accountRange =  sheet.createTextFinder(AD_ACC_ID).findNext();
    const lastRow = sheet.getLastRow();
    if(accountRange != null){
      var accountRow =  accountRange.getRow();
      var accStatus = sheet.getRange(accountRow, statusCol).getValue();
      var getSSUrl = sheet.getRange(accountRow, associatedSSUrlCol).getValue();
      //Logger.log("account Id Row:" + accountRow);
      //Logger.log("accStatus:" + accStatus);
      //Logger.log("getSSUrl:" + getSSUrl);
      
      if(getSSUrl.toString() == '' && accStatus.toString().toLowerCase() !== "ignore"){
        updateAdAcc += AD_ACC_ID + ',';
      }
      
      //Logger.log("updateAdAcc:" + updateAdAcc);
    }else{
      sheet.getRange(lastRow + 1, adAccountIdCol).setValue(AD_ACC_ID);
      var facebookUrl = `https://graph.facebook.com/v7.0/${AD_ACC_ID}/?fields=name&access_token=${TOKEN}`;
      const results = fetchFBJSONResponse(facebookUrl);
      sheet.getRange(lastRow + 1, nameCol).setValue(results.name);
      newAdAcc += AD_ACC_ID + ',';
    }
  }
  
}

function sendEmails(){
    if(newAdAcc != ''){
    var message = `A new ad accounts ${newAdAcc} is added. Setup a spreadsheet URL`;
      MailApp.sendEmail("no-reply@domain.com", "Added a new Ad Account", message);
  }
  
  if(updateAdAcc != ''){
    var message = `These ad accounts ${updateAdAcc} has missing spreadsheet URL`;
      MailApp.sendEmail("no-reply@domain.com", "Update spreadsheet URLs", message);
  }
}

function fetchFBJSONResponse(fbUrl){
  const encodedFacebookUrl = encodeURI(fbUrl);
  
  const options = {
    'method' : 'get'
  };
  
  // Fetches & parses the URL 
  const fetchRequest = UrlFetchApp.fetch(encodedFacebookUrl, options);
  return JSON.parse(fetchRequest.getContentText());
}
