// Doesn't require special permission, just follow setup and authorize
// Original script by John McLaughlin (loghound@gmail.com)
// Modifications - Simon Bromberg (http://sbromberg.com)
// Modifications - Mark Leavitt (PDX Quantified Self organizer) www.markleavitt.com
// Modifications 2020 - Jozef Jarosciak - joe0.com
// Modifications 2022 - Josh Kybett - JKybett.uk
//    -Replaced discontinued UiApp code to use HtmlService instead.
//    -Replace deprecated v1 FitBit API with current standard v2 FitBit API
//    -Now fetches data using daily summaries rather than per-item ranges to avoid hitting API limits when getting single-day data.
//    -Adapted to get data for more features of FitBit.
//    -Friendlier setup UI.
//
// Current version on GitHub: https://github.com/JKybett/GoogleFitBit/blob/main/FitBit.gs
//
// This is a free script: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
// GNU General Public License for more details.
// 
// You should have received a copy of the GNU General Public License
// along with this program. If not, see <https://www.gnu.org/licenses/>.
// 
// Copyright (c) 2022 JKybett

/*
 * Do not change these key names. These are just keys to access these properties once you set them up by running the Setup function from the Fitbit menu
 */
// Key of ScriptProperty for Firtbit consumer key.
var CONSUMER_KEY_PROPERTY_NAME = "fitbitConsumerKey";
// Key of ScriptProperty for Fitbit consumer secret.
var CONSUMER_SECRET_PROPERTY_NAME = "fitbitConsumerSecret";

var SERVICE_IDENTIFIER = 'fitbit'; // usually do not need to change this either

// List of all things this script logs
var LOGGABLES = [
  
                 "activeScore",
                 
                 "activityCalories",
                 "caloriesBMR",
                 "caloriesOut",
                 "elevation",
                 "fairlyActiveMinutes",
                 "floors",
                 "lightlyActiveMinutes",
                 "marginalCalories",
                 "sedentaryMinutes",
                 "steps",
                 "veryActiveMinutes",
                 "bmi",
                 "weight",
                 "awakeCount",
                 "awakeDuration",
                 "awakeningsCount",
                 "duration",
                 "efficiency",
                 "endTime",
                 "minutesAfterWakeup",
                 "minutesAsleep",
                 "minutesAwake",
                 "minutesToFallAsleep",
                 "restlessCount",
                 "restlessDuration",
                 "startTime",
                 "timeInBed",

                 
                 "calories",
                 "carbs",
                 "fat",
                 "fiber",
                 "protein",
                 "sodium",
                 "water"
                 ];
                 
//List of loggables that come from the activities section of API
var LOGGABLE_ACTIVITIES = [
                 "activeScore",
                 "activityCalories",
                 "caloriesBMR",
                 "caloriesOut",
                 "elevation",
                 "fairlyActiveMinutes",
                 "floors",
                 "lightlyActiveMinutes",
                 "marginalCalories",
                 "sedentaryMinutes",
                 "steps",
                 "veryActiveMinutes"
                 ];

//List of loggables that come from the weight section of API
var LOGGABLE_WEIGHT = [
                 "bmi",
                 "weight"
                 ];

//List of loggables that come from the sleep section of API
var LOGGABLE_SLEEP = [
                 "awakeCount",
                 "awakeDuration",
                 "awakeningsCount",
                 "duration",
                 "efficiency",
                 "endTime",
                 "minutesAfterWakeup",
                 "minutesAsleep",
                 "minutesAwake",
                 "minutesToFallAsleep",
                 "restlessCount",
                 "restlessDuration",
                 "startTime",
                 "timeInBed"
                 ];

//List of loggables that come from the food section of API
var LOGGABLE_FOOD = [
                 "calories",
                 "carbs",
                 "fat",
                 "fiber",
                 "protein",
                 "sodium",
                 "water"
                 ];


/*
******************************************************************************************************************
******************************************************************************************************************
******************************************************************************************************************
*/

// /*
//   Used to display information to the user via cell B3 to let them know that scripts have stopped actively running.
// */
// function done(){
//   getSheet().getRange("R3C2").setValue("Ready");
// }

// /********************

// */
// function isConfigured() {
//   return getConsumerKey() != "" && getConsumerSecret() != "";
// }

// /********************

// */
// function getProperty(key) {
//   // Logger.log("get property "+key);
//   return PropertiesService.getScriptProperties().getProperty(key);
// }

/********************

*/
function setProperty(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, value);
}

/********************

*/
function getSheet(){
  try {
    var spreadSheetID = PropertiesService.getScriptProperties().getProperty("spreadSheetID");
    // console.log(spreadSheetID);
    var spreadSheet = SpreadsheetApp.openById(spreadSheetID.toString());
    var sheetID = PropertiesService.getScriptProperties().getProperty("sheetID");
    var sheet = spreadSheet.getSheets().filter(
      function(s) {return s.getSheetId().toString() === sheetID.toString()}
    )[0];
    return sheet;
  } catch (error) {
    return null;
  }
}

/********************

*/
function setSheet(sheet){
  if(sheet == null){
    setProperty("sheetID", "");
    setProperty("spreadSheetID", "");
  } else {
    setProperty("sheetID", sheet.getSheetId().toString());
    setProperty("spreadSheetID", sheet.getParent().getId().toString());
  }
}

/********************

*/
function setConsumerKey(consumerKey) {
  setProperty(CONSUMER_KEY_PROPERTY_NAME, consumerKey);
}

/********************

*/
function getConsumerKey() {
  var consumer = PropertiesService.getScriptProperties().getProperty(CONSUMER_KEY_PROPERTY_NAME);
  if (consumer == null) {
      consumer = "";
  }
  return consumer;
}

/********************

*/
function setConsumerSecret(secret) {
  setProperty(CONSUMER_SECRET_PROPERTY_NAME, secret);
}

/********************

*/
function getConsumerSecret() {
  var secret = PropertiesService.getScriptProperties().getProperty(CONSUMER_SECRET_PROPERTY_NAME);
  if (secret == null) {
      secret = "";
  }
  return secret;
}

function clearService(){
  OAuth2.createService(SERVICE_IDENTIFIER)
  .setPropertyStore(PropertiesService.getUserProperties())
  .reset();
  setConsumerKey("");
  setConsumerSecret("");
  setSheet(null);
}

function getFitbitService() {
  // Create a new service with the given name. The name will be used when
  // persisting the authorized token, so ensure it is unique within the
  // scope of the property store
  
  if (!(getConsumerKey() != "" && getConsumerSecret() != "")) {
      setup();
      return;
  }
                
  return OAuth2.createService(SERVICE_IDENTIFIER)
  
  // Set the endpoint URLs, which are the same for all Google services.
  .setAuthorizationBaseUrl('https://www.fitbit.com/oauth2/authorize')
  .setTokenUrl('https://api.fitbit.com/oauth2/token')
  
  // Set the client ID and secret, from the Google Developers Console.
  .setClientId(getConsumerKey())
  .setClientSecret(getConsumerSecret())
  
  // Set the name of the callback function in the script referenced
  // above that should be invoked to complete the OAuth flow.
  .setCallbackFunction('authCallback')
  
  // Set the property store where authorized tokens should be persisted.
  .setPropertyStore(PropertiesService.getUserProperties())
  .setScope('activity nutrition sleep weight profile settings')
  // but not desirable in a production application.
  //.setParam('approval_prompt', 'force')
  .setTokenHeaders({
      'Authorization': 'Basic ' + Utilities.base64Encode(getConsumerKey() + ':' + getConsumerSecret())
  });
}

function submitData(form) {
  switch(form.task){
    case "setup": saveSetup(form); break;
    case "sync" : syncDate(new Date(form.year, form.month-1, form.day)); break;
    case "syncMany" : syncMany(new Date(form.firstYear, form.firstMonth-1, form.firstDay),new Date(form.secondYear, form.secondMonth-1, form.secondDay)); break;
    case "BackToFitBitAPI" : firstRun();break;
    case "FitBitAPI" : setup();break;
    //case "credits" : credits();break;
  }
}

// function saveSetup saves the setup params from the UI
function saveSetup(e) {
  //problemPrompt(e.spreadSheetID);
  var doc = SpreadsheetApp.openById(e.spreadSheetID);
  if(parseInt(e.newSheet)>0){
    if(e.sheetID.length<1){
      e.sheetID="FitbitData";
    }
    doc=doc.insertSheet(e.sheetID.toString());
    e.sheetID = doc.getSheetId();
  }
  var doc = SpreadsheetApp.openById(e.spreadSheetID);
  doc=doc.getSheets().filter(
      function(s) {return s.getSheetId().toString() === e.sheetID.toString();}
    )[0];
  //problemPrompt("'"+e.sheetID+"'");
  setSheet(doc);
  working();
  doc.getRange("R2C2").setValue(new Date(e.year, e.month-1, e.day));
  console.log(e);
  setConsumerKey(e.consumerKey);
  setConsumerSecret(e.consumerSecret);
  var i=2;
  var cell = doc.getRange("R4C2");
  var titles = [];
  var wanted = [];
  while(!cell.isBlank()){
    titles.push(cell.getValue());
    cell = doc.getRange("R4C"+(++i));
    wanted.push(false);
  }
  var index = -1;
  for (const [key, value] of Object.entries(e.loggables)) {
    index = titles.findIndex(e=>{return e==value});
    if(index<0){
      titles.push(value);
      wanted.push(true);
    } else {
      wanted[index]=true;
    }
  }
  for(i=0;i<wanted.length;i++){
    if(!wanted[i]){
      titles[i]="";
    }
  }
  i=0;
  for (const [key, value] of Object.entries(titles)) {
    doc.getRange("R4C"+(2+i)).setValue(value);
    i++;
  }
  doc.getRange("R1C1").setValue("Sheet last synced: never");
  doc.getRange("R2C1").setValue("Sheet Start Date:");
  doc.getRange("R3C1").setValue("Status:");
  doc.getRange("R4C1").setValue("Date");
  authWindow();
  working("Ready");
}

/*
*******************

*/
function fetchNeeded(doc,loggables){
  var titles = doc.getRange("4:4").getValues();
  return loggables.some(r=> titles[0].includes(r))
}


/*
*******************

*/

function logAllTheThings(doc, row, entries) {
  if (!entries) return false;
  
  var titles = doc.getRange("4:4").getValues()[0];
  var relevantCols = [];
  var minCol = 9999, maxCol = 0;
  
  for (const [k, v] of Object.entries(entries)) {
    var colIndex = titles.indexOf(k);
    if (colIndex >= 0) { 
      relevantCols[colIndex] = v;  
      minCol = Math.min(minCol, colIndex + 1);
      maxCol = Math.max(maxCol, colIndex + 1);
    }
  }
  
  if (minCol > maxCol) return false;
  
  var numCols = maxCol - minCol + 1;
  var rowData = new Array(numCols).fill('');
  
  for (var col = minCol; col <= maxCol; col++) {
    var colIndex = col - 1;
    rowData[col - minCol] = relevantCols[colIndex] !== undefined ? relevantCols[colIndex] : '';
  }
  
  doc.getRange(row, minCol, 1, numCols).setValues([rowData]);
  return true;
}


// function logAllTheThings(doc, row, entries) {
//   if (!entries) return false;
  
//   var titles = doc.getRange("4:4").getValues()[0];
//   var relevantCols = [];  // Znalezione kolumny dla danych
//   var rowData = [];
//   var minCol = 9999;
//   var maxCol = 0;
  
//   // 1. Znajdź wszystkie pasujące kolumny AUTOMATYCZNIE
//   for (const [k, v] of Object.entries(entries)) {
//     var colIndex = titles.indexOf(k);
//     if (colIndex >= 0) {
//       relevantCols[colIndex] = v;
//       minCol = Math.min(minCol, colIndex + 1);
//       maxCol = Math.max(maxCol, colIndex + 1);
//     }
//   }
  
//   if (minCol > maxCol) return false;  // Brak pasujących kolumn
  
//   // 2. Bufor TYLKO dla znalezionego zakresu
//   var numCols = maxCol - minCol + 1;
//   rowData = new Array(numCols).fill('');
  
//   // 3. Wypełnij bufor
//   for (var col = minCol; col <= maxCol; col++) {
//     var colIndex = col - 1;
//     rowData[col - minCol] = relevantCols[colIndex] || '';
//   }
  
//   // 4. Batch zapis TYLKO potrzebnego zakresu!
//   doc.getRange(row, minCol, 1, numCols).setValues([rowData]);
  
//   return true;
// }


// function logAllTheThings(doc, row, entries) {
//   if (!entries) {
//     console.log("No data. Abort.");
//     return false;
//   }
  
//   var titles = doc.getRange("4:4").getValues()[0];  // Płaska tablica
//   var maxCol = titles.length;
//   var rowData = new Array(maxCol).fill('');  // Bufor [ '', '', '', ... ]
  
//   // Znajdź kolumny i zapisz do bufora OJCZAS
//   for (const [k, v] of Object.entries(entries)) {
//     var col = titles.indexOf(k) + 1;  // indexOf zamiast findIndex
//     if (col > 0) {
//       rowData[col - 1] = v;  // Bufor w pamięci
//     }
//   }
  
//   // JEDEN setValues() na koniec!
//   if (rowData.some(cell => cell !== '')) {
//     doc.getRange(row, 1, 1, maxCol).setValues([rowData]);
//   }
  
//   return true;
// }


// function logAllTheThings(doc,row,entries){
//   var col;
//   var titles = doc.getRange("4:4").getValues();
//   if(entries==null || entries == undefined){
//     console.log("No column names. Abort.");
//     return false;
//   }
//   for (const [k, v] of Object.entries(entries)) {
//     col = titles[0].findIndex(e=>{return e==k})+1;
//     if(col>0){
//       doc.getRange("R"+row+"C"+col).setValue(v);
//       // console.log("  logged:"+k);
//     } else {
//       // console.log("unlogged:"+k);
//     }
//   }
//   return true;
// }



/*
*******************

*/
function firstRun(){
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var contentHTML='<html>'+"\n"+
	'<head>'+"\n"+
	'	<style>'+"\n"+
	'	label, input {'+"\n"+
	'		width:95%;'+"\n"+
	'	}'+"\n"+
	'    .radio {'+"\n"+
	'    	width:initial;'+"\n"+
	'    }'+"\n"+
	'    .box {'+"\n"+
	'    	border-style: solid;'+"\n"+
	'        padding: 5px;'+"\n"+
	'        margin-bottom: 10px;'+"\n"+
	'    }'+"\n"+
	'    #hidden {'+"\n"+
	'    	display: none;'+"\n"+
	'    }'+"\n"+
	'	</style>'+"\n"+
	'</head>'+"\n"+
	'  <body>'+"\n"+
	'      Go to <a href="https://dev.fitbit.com/apps/new">https://dev.fitbit.com/apps/new</a></br></br>'+"\n"+
	'      Login and register a new app using the following details:</br></br>'+"\n"+
	'    <div class="box" id="hider">'+"\n"+
	'        Only the options that must have specific values are shown below.</br>'+"\n"+
	'        <a href="#" onclick="document.getElementById(\'hidden\').style.display=\'block\';document.getElementById(\'hider\').style.display=\'none\';return false;">Click here</a> for example data you can copy and paste into the other fields.'+"\n"+
	'    </div>'+"\n"+
	'    <div class="box" id="hidden">'+"\n"+
	'        These options can be filled with different data. This is only an example.</br>'+"\n"+
	'        You can <a href="#" onclick="document.getElementById(\'hider\').style.display=\'block\';document.getElementById(\'hidden\').style.display=\'none\';return false;">hide these options</a> if you want.'+"\n"+
	'        </br></br>'+"\n"+
	'        <label>Application Name: </label></br><input type="text" value="Export to Google Spreadsheet" readonly></br></br>'+"\n"+
	'        <label>Description: </label></br><input type="text" value="Exports to Google Spreadsheet" readonly></br></br>'+"\n"+
	'        <label>Application Website URL: </label></br><input type="text" value="https://docs.google.com/" readonly></br></br>'+"\n"+
	'        <label>Organization: </label></br><input type="text" value="Me" readonly></br></br>'+"\n"+
	'        <label>Organization Website URL: </label></br><input type="text" value="https://docs.google.com/" readonly></br></br>'+"\n"+
	'        <label>Terms of Service URL: </label></br><input type="text" value="https://docs.google.com/" readonly></br></br>'+"\n"+
	'        <label>Privacy Policy URL: </label></br><input type="text" value="https://docs.google.com/" readonly></br></br>'+"\n"+
	'    </div>'+"\n"+
	'    <div class="box">'+"\n"+
	'        These options <b>must</b> be filled with the following data.</br></br>'+"\n"+
	'        <label>OAuth 2.0 Application Type: </label></br>'+"\n"+
	'        <input class="radio" type="radio" name="appType" id="Server" disabled>'+"\n"+
	'        <label class="radio" for="Server">Server</label>'+"\n"+
	'        <input class="radio" type="radio" name="appType" id="Client" disabled>'+"\n"+
	'        <label class="radio" for="Client">Client</label>'+"\n"+
	'        <input class="radio" type="radio" name="appType" id="Personal" checked>'+"\n"+
	'        <label class="radio" for="Personal">Personal</label></br></br>'+"\n"+
	'        <label>Redirect URL: </label></br><input type="text" value="https://script.google.com/macros/d/'+ScriptApp.getScriptId()+'/usercallback" readonly></br></br>'+"\n"+
	'        <label>Default Access Type: </label></br>'+"\n"+
	'        <input class="radio" type="radio" name="accessType" id="RWr" checked>'+"\n"+
	'        <label class="radio" for="RWr">Read & Write</label>'+"\n"+
	'        <input class="radio" type="radio" name="accessType" id="ROn" disabled>'+"\n"+
	'        <label class="radio" for="ROn">Read-Only</label></br></br>'+"\n"+
	'    </div>'+"\n"+
	'    Once you have accepted the terms and conditions and clicked "register", make a note of the following details on the next page:</br>'+"\n"+
	'    <ul>'+"\n"+
	'    <li><b>OAuth 2.0 Client ID</b></li>'+"\n"+
	'    <li><b>Client Secret</b></li>'+"\n"+
	'    </ul>'+"\n"+
	'    Then click the button below to move on to the next step:'+"\n"+
	'    <form id="form">'+"\n"+
	'    <input type="hidden" id="task" name="task" value="FitBitAPI">'+"\n"+
	'    <input class="normWid" type="button" value="Next" onclick="'+"\n"+
	'    google.script.run.withSuccessHandler(function(value){'+"\n"+
	'    }).submitData(form);document.getElementById(\'done\').style.display = \'block\';">'+"\n"+
	'    </form>'+"\n"+
	'    <p id="done" style="display:none;">Please wait!</p>'+"\n" +
  '   </br></br><div style="text-align: right;font-style: italic;">By <a href="https://jkybett.uk" target="_blank">JKybett</a></div>' +
	'  </body>'+"\n"+
	'</html>';
  var app= HtmlService.createHtmlOutput().setTitle("Setup: FitBit App").setContent(contentHTML);
  doc.show(app);
}               

/*
*******************

*/
function setup() {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var selected;
  var sheets = doc.getSheets();
  var selectSheet = doc.getActiveSheet();
  var earliestDate = new Date();
  if(getSheet()!=null){
    selectSheet = getSheet();
    earliestDate = getSheet().getRange("R2C2").getValue();
  }
  var contentHTML =''+
  '<!DOCTYPE html>'+"\n"+
  '<html>'+"\n"+
	' <head>'+"\n"+
  '   <style>'+"\n"+
  '     label, input, select {'+"\n"+
  '       width: 45%;'+"\n"+
  '       display: inline-block;'+"\n"+
  '       vertical-align: top;'+"\n"+
  '     }'+"\n"+
  '     label{'+"\n"+
  '     }'+"\n"+
  '     input, select {'+"\n"+
  '       text-align: right;'+"\n"+
  '     }'+"\n"+
  '     .half {'+"\n"+
  '       width: 50%;'+"\n"+
  '     }'+"\n"+
  '     .full {'+"\n"+
  '       width: 100%;'+"\n"+
  '     }'+"\n"+
  '     .right {'+"\n"+
  '       text-align: right;'+"\n"+
  '       margin-right: 0px;'+"\n"+
  '     }'+"\n"+
  '     .normWid {'+"\n"+
  '       width: initial;'+"\n"+
  '     }'+"\n"+
  '     .sheetName {'+"\n"+
  '       visibility: hidden;'+"\n"+
  '     }'+"\n"+
  '   </style>'+"\n"+
	' </head>'+"\n"+
	' <body>'+"\n"+
  '   <form id="backForm">'+"\n"+
  '     <input type="hidden" id="task" name="task" value="BackToFitBitAPI">'+"\n"+
  '     <center>'+
  '       <input class="normWid" type="button" value="<<< Setup FitBit App" onclick="'+
  '         google.script.run'+
  '         .withSuccessHandler(function(value){'+
  '         })'+
  '         .submitData(backForm);'+
  '">'+"\n"+
  '     </center>'+
  '   </form>'+"\n"+
	'   <form id="form">'+"\n"+
  '     <input type="hidden" id="task" name="task" value="setup">'+"\n"+
  '     <input type="hidden" id="spreadsheetID" name="spreadSheetID" value="'+doc.getId().toString()+'">'+"\n"+
	'     <label class="normWid">Script ID: </label>'+"\n"+
  '     <label class="normWid right">'+ScriptApp.getScriptId()+'</label></br></br>'+"\n\n"+
  '     <label>Fitbit OAuth 2.0 Client ID:*</label>'+"\n"+
	'     <input type="text" id="consumerKey" name="consumerKey" value="'+getConsumerKey()+'"></br>'+"\n\n"+		
	'     <label>Fitbit OAuth Consumer Secret:*</label>'+"\n"+
  '     <input type="text" id="consumerSecret" name="consumerSecret" value="'+getConsumerSecret()+'"></br></br>'+"\n\n"+
	'     <label>Earliest Date (year-month-day): </label>'+"\n"+
	'     <input class="normWid" type="text" maxlength="4" size="4" id="year" name="year" value="'+(earliestDate.getFullYear())+'">'+" -\n\n"+
	'     <input class="normWid" type="text" maxlength="2" size="2" id="month" name="month" value="'+(earliestDate.getMonth()+1)+'">'+" -\n\n"+
	'     <input class="normWid" type="text" maxlength="2" size="2" id="day" name="day" value="'+(earliestDate.getDate())+'"></br>'+"\n\n"+		
  '     <label>Data Elements to download: </label>'+"\n"+
	'     <select id="loggables" name="loggables" multiple>'+"\n";
	for (var resource in LOGGABLES) {
    selected = (LOGGABLES.indexOf(LOGGABLES[resource]) > -1)?" selected":"";
    contentHTML +='       <option value="'+LOGGABLES[resource]+'"'+selected+'>'+
                  LOGGABLES[resource]+
                  '</option>'+"\n";
  }
  contentHTML +=
	'     </select></br></br>'+"\n"+
  '     <label>Sheet to store data: </label>'+"\n"+
  '     <select id="sheets" onchange=\''+
  '       var val = this.value;'+
  '       document.getElementById("newSheet").value="1";'+
  '       document.getElementById("sheetID").value=val=="new"?"":val;'+
  '       var hiders = document.getElementsByClassName("sheetName");'+
  '       var display=val=="new"?"visible":"hidden";'+
  '       for (const item of hiders) {'+
  '         item.style.visibility = display;'+
  '       }'+
  '\'>'+"\n";
  if (sheets.length > 0) {
    for(var i =0; i <sheets.length; i++){
      selected = sheets[i].getSheetId()==selectSheet.getSheetId()?" selected":"";
      contentHTML += 
      '       <option value="'+sheets[i].getSheetId()+'"'+selected+">\n"+
      '         '+sheets[i].getName()+"\n"+
      '       </option>'+"\n";
    }
  }
  contentHTML+=
	'       <option value="new">'+"\n"+
  '       + New sheet'+"\n"+
  '       </option>'+"\n"+
	'     </select></br>'+"\n"+
  '     <label class="sheetName">Name:</label>'+"\n"+
	'     <input type="text" id="sheetID" name="sheetID" value="'+selectSheet.getSheetId()+'" class="sheetName"></br></br>'+"\n"+	
  '     <input type="hidden" id="newSheet" name="newSheet" value="0">'+"\n"+
	'     <center>'+
	'      <input class="normWid" type="button" value="Submit" onclick="'+
  '         google.script.run'+
  '         .withSuccessHandler(function(value){'+
  '         })'+
  '         .submitData(form);'+
  '         document.getElementById(\'form\').style.display === \'none\';'+
	'         document.getElementById(\'done\').style.display = \'block\';'+
  '">'+"\n"+
	'     </center>'+
  '    </form>'+"\n"+
	'	  <p id="done" style="display:none;">Please wait!</p>'+"\n"+
  '  </br></br><div style="text-align: right;font-style: italic;">By <a href="https://jkybett.uk" target="_blank">JKybett</a></div>'+
	' </body>'+"\n"+
  '</html>';
  var app= HtmlService.createHtmlOutput().setTitle("Setup Fitbit Download").setContent(contentHTML);
  doc.show(app);
}

/*
*******************

*/

function authWindow(){
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var service = getFitbitService();
  var authorizationUrl = service.getAuthorizationUrl();
  var contentHTML ='<a href="'+authorizationUrl+'" target="_blank">Click here to Authorize with Fitbit</a>'
        +  '</br></br><div style="text-align: right;font-style: italic;">By <a href="https://jkybett.uk" target="_blank">JKybett</a></div>';
  var app= HtmlService.createHtmlOutput().setTitle("Setup Fitbit Download").setContent(contentHTML);
  doc.show(app);
}

/*
*******************

*/

function authCallback(request) {
  Logger.log("authcallback");
  var service = getFitbitService();
  var isAuthorized = service.handleCallback(request);
  var app;
  var contentHTML;
  if (isAuthorized) {
    var displayContentHTML = 'Success! Please refresh the page .'
      + '</br></br><div style="text-align: right;font-style: italic;">By <a href="https://jkybett.uk" target="_blank">JKybett</a></div>';
    var displayApp= HtmlService.createHtmlOutput().setTitle("All done!").setContent(displayContentHTML);
    contentHTML = 'Success! You can close this tab.';
    app= HtmlService.createHtmlOutput().setTitle("Authorised!").setContent(contentHTML);
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    doc.show(displayApp);
  } else {
    contentHTML = 'Authorisation was denied.</br>Please check your FitBit credentials and try again!';
    app= HtmlService.createHtmlOutput().setTitle("Oh no!").setContent(contentHTML);
  }
  return app;
} 


/*
*******************

*/

function problemPrompt(problem="Undefined problem.", pTitle = "There was a problem!"){
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var contentHTML =''+
  '<!DOCTYPE html>'+"\n"+
  '<html>'+"\n"+
	' <body>'+"\n"+
	'	  <p>Something went wrong! Here\'s the message from the code:</p>'+"\n"+
	'	  <code>'+problem+'</code>'+"\n"+
	'	  <p>This is just to let you know. You can close this window if you want.</p>'+"\n"+
  '  </br></br><div style="text-align: right;font-style: italic;">By <a href="https://jkybett.uk" target="_blank">JKybett</a></div>' +
	' </body>'+"\n"+
  '</html>';
  var app= HtmlService.createHtmlOutput().setTitle(pTitle).setContent(contentHTML);
  doc.show(app);
}


/*
*******************

*/

// function signature(){
//   return "</br></br><div style='text-align: right;font-style: italic;'>By <a href='https://jkybett.uk' target='_blank'>JKybett</a></div>";
// }


/*
*******************

*/

function credits(){
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var contentHTML =''+
  '<!DOCTYPE html>'+"\n"+
  '<html>'+"\n"+
	' <body>'+"\n"+
	'	  <p>Something went wrong! Here\'s the message from the code:</p>'+"\n"+
	'	  <code>'+problem+'</code>'+"\n"+
	'	  <p>This is just to let you know. You can close this window if you want.</p>'+"\n"
  + '</br></br><div style="text-align: right;font-style: italic;">By <a href="https://jkybett.uk" target="_blank">JKybett</a></div>' +
	' </body>'+"\n"+
  '</html>';
  var app= HtmlService.createHtmlOutput().setTitle("").setContent(contentHTML);
  doc.show(app);
}
