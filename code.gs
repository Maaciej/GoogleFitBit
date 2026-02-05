/*
******************************************************************************************************************
******************************************************************************************************************
******************************************************************************************************************
*/


/*
  Used to display information to the user via cell B3 to let them know that scripts are actively running.
*/
function working(stepStr = "Initialized"){
  getSheet().getRange("R3C2").setValue(stepStr);
  console.log(stepStr)
}



/*
*******************


*/

function logJsonRecursive(obj, prefix = '') {
  var keys = Object.keys(obj);
  
  keys.forEach(key => {
    var value = obj[key];
    var displayValue;
    
    if (typeof value === 'object' && value !== null) {
      displayValue = '[OBJECT]';
      Logger.log(`| ${prefix}${key.padEnd(15)} | ${displayValue} |`);
      logJsonRecursive(value, prefix + '  ');  // Rekursja!
    } else {
      displayValue = String(value).substring(0, 50);
      Logger.log(`| ${prefix}${key.padEnd(15)} | ${displayValue} |`);
    }
  });
}

/*
*******************


*/


function test() {
  var doc = getSheet();
  var service = getFitbitService();
  
  // Lista endpointów Fitbit API do pobrania danych
  var endpoints = [
    {name: "sleep-summary", url: "https://api.fitbit.com/1/user/-/sleep/date/" + dateString + ".json"},
    {name: "sleep-stages", url: "https://api.fitbit.com/1.2/user/-/sleep/stages/date/" + dateString + ".json"},
    {name: "sleep-day-details", url: "https://api.fitbit.com/1/user/-/sleep/date/" + dateString + "?summary=shortData"},
    {name: "activities-summary", url: "https://api.fitbit.com/1/user/-/activities/date/" + dateString + ".json"},
    {name: "heart-rate", url: "https://api.fitbit.com/1/user/-/activities/heart/date/" + dateString + "/1d/1sec.json"},
    {name: "body-weight", url: "https://api.fitbit.com/1/user/-/body/log/weight/date/" + dateString + ".json"},
    {name: "steps", url: "https://api.fitbit.com/1/user/-/activities/steps/date/" + dateString + "/1d.json"},
    {name: "distance", url: "https://api.fitbit.com/1/user/-/activities/distance/date/" + dateString + "/1d.json"},
    {name: "calories", url: "https://api.fitbit.com/1/user/-/activities/calories/date/" + dateString + "/1d.json"}
  ];
  
  var options = { 
    headers: { "Authorization": 'Bearer ' + service.getAccessToken() }, 
    method: "GET",
    muteHttpExceptions: true 
  };

  endpoints.forEach(function(endpoint) {
    var dateFormatted = dateString.replace(/-/g, ''); // YYYYMMDD format
    var filename = 'fitbit_' + endpoint.name + '_' + dateFormatted + '.json';
    
    var result = UrlFetchApp.fetch(endpoint.url, options);
    console.log(endpoint.name + ': ' + result.getResponseCode());
    
    if (result.getResponseCode() === 429) {
      console.log('429 Rate limit! BREAK: ' + endpoint.name);
      doc.getRange(workingRow, 1, 1, 37).clearContent();
      return; // przerwij pętlę przy rate limit
    }

    if (result.getResponseCode() === 200) {
      var jsonText = result.getContentText(); 
      var jsonObj = JSON.parse(jsonText); 
      var jsonString = JSON.stringify(jsonObj, null, 2);
      
      // Zapisz każdy endpoint do osobnego pliku z datą
      DriveApp.createFile(filename, jsonString, MimeType.PLAIN_TEXT);
      
      Logger.log('|' + endpoint.name + '| Zapisano ' + filename + '|');
      Logger.log('---');
      logJsonRecursive(jsonObj);
      Logger.log('');
    } else {
      Logger.log('|' + endpoint.name + '| BŁĄD: ' + result.getResponseCode() + '|');
    }
    
    // Opóźnienie między requestami żeby nie trafić w rate limit
    Utilities.sleep(1000);
  });
}


// function test() {

//   var doc = getSheet();

//   var dateString = '2026-01-20';

//   var service = getFitbitService();
//   var options = { 
//     headers: { "Authorization": 'Bearer ' + service.getAccessToken() }, 
//     method: "GET",
//     muteHttpExceptions: true 
//   };

//   // ACTIVITIES
// // var LOGGABLE_ACTIVITIES = [
// //                  "activeScore",
// //                  "activityCalories",
// //                  "caloriesBMR",
// //                  "caloriesOut",
// //                  "elevation",
// //                  "fairlyActiveMinutes",
// //                  "floors",
// //                  "lightlyActiveMinutes",
// //                  "marginalCalories",
// //                  "sedentaryMinutes",
// //                  "steps",
// //                  "veryActiveMinutes"
// //                  ];



//     var result = UrlFetchApp.fetch("https://api.fitbit.com/1/user/-/sleep/date/"+dateString+".json", options);
//     console.log(result.getResponseCode())
//     if (result.getResponseCode() === 429) {
//       console.log('429 Rate limit! BREAK: ' + dateString);
//       doc.getRange(workingRow, 1, 1, 37).clearContent();
//       return 2;  //ratelimit
//     }


//     if (result.getResponseCode() === 200) {
      
//       var jsonText = result.getContentText(); 
//       var jsonObj = JSON.parse(jsonText); 

//       var jsonString = JSON.stringify(jsonObj, null, 2);
//       DriveApp.createFile('fitbit_sleep.json', jsonString, MimeType.PLAIN_TEXT);

      
//       Logger.log('| Klucz | Wartość |');
//       Logger.log('|-------|---------|');
//       logJsonRecursive(jsonObj);

//     }


// }

/*
*******************

*/

function STOP_EXECUTION_AND_CLEAN() {
  var props = PropertiesService.getScriptProperties();

  // props.deleteProperty('Start');
  props.deleteProperty('End');
  // props.deleteProperty('Today');

  props.deleteProperty('Next');
  props.deleteProperty('Total');
  props.deleteProperty('daysdone');

  var triggers = ScriptApp.getProjectTriggers();
  var count = 0;
  
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'resumeSync' )  {
      ScriptApp.deleteTrigger(t);
      count++;
    }
  });
  
  working('Cleared ALL triggers ' + count + ' sync triggers');

}

/*
  function syncDate() is called to download one day data from Fitbit API to the spreadsheet
*/

function syncDate(date = null) {
  // console.log('=== syncDate(' + (date || 'today') + ') ===');
  var dateString = Utilities.formatDate(date || new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  var doc = getSheet();
  if (!doc) {
    console.error('No sheet, RUN Setup!');
    throw new Error("CRITICAL: No sheet");
  }

// checking if date is in range
  var b2DateString = Utilities.formatDate(new Date( doc.getRange("R2C2").getValue() ), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var tomorrowString = Utilities.formatDate(new Date(new Date().getTime() + 24*60*60*1000), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  if (dateString < b2DateString || dateString >= tomorrowString) {
    working('DATE out of RANGE');
    return -1; //out of range
  }
  
  testdate = date || new Date()
  var dayMil = 1000 * 60 * 60 * 24;
  var firstDay = getSheet().getRange("R2C2").getValue();
  testdate = (testdate-firstDay);
  workingRow = 5 + (testdate-(testdate%dayMil))/dayMil;

  if (workingRow > 4) {
    var range = doc.getRange(workingRow, 1, 1, 29);
    var rowValues = range.getValues()[0];
    
    var emptyCount = 0, totalChecked = 0;
    for (var j = 0; j < 29; j++) {
      if (j === 13 || j === 14) continue;
      totalChecked++;
      var cellValue = rowValues[j];
      if (cellValue === null || cellValue === undefined || cellValue === '' || 
          (typeof cellValue === 'string' && cellValue.trim() === '')) {
        emptyCount++;
      }
    }
    
    if (emptyCount <= 2) {
      console.log('Row ' + workingRow + ': filled - skipping ' + dateString);
      return 1;//skipped
    }
  }
  
  working("Sync " + dateString + " (L = " + workingRow +")");
  
  if (!(getConsumerKey() != "" && getConsumerSecret() != "")) {
    console.error('Fitbit not configured!');
    throw new Error("CRITICAL: Fitbit not configured");
  }
  
  doc.setFrozenRows(4);
  doc.getRange("R1C1").setValue("Sheet last synced: " + 
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"));
  doc.getRange("R4C1").setValue("Date");
  doc.getRange("R"+workingRow+"C1").setValue(dateString);
  
  var service = getFitbitService();
  var options = { 
    headers: { "Authorization": 'Bearer ' + service.getAccessToken() }, 
    method: "GET",
    muteHttpExceptions: true 
  };

//GAS quota try
try{

    // ACTIVITIES
    if(fetchNeeded(doc, LOGGABLE_ACTIVITIES)) {
      var result = UrlFetchApp.fetch("https://api.fitbit.com/1/user/-/activities/date/"+dateString+".json", options);

      if (result.getResponseCode() === 200) {
        var Stats = JSON.parse(result.getContentText());
        if (Stats["summary"]) {
          logAllTheThings(doc, workingRow, Stats["summary"]);  // ← summary!
        }
      }
      else if (result.getResponseCode() === 429) {
          console.log('429 Rate limit! BREAK: ' + dateString);
          doc.getRange(workingRow, 1, 1, 37).clearContent();//clean partial row
          return 2;  //ratelimit
        }

      }
    

    
    // WEIGHT
    if(fetchNeeded(doc, LOGGABLE_WEIGHT)) {
      var result = UrlFetchApp.fetch("https://api.fitbit.com/1/user/-/body/log/weight/date/"+dateString+".json", options);

      if (result.getResponseCode() === 200 && result.getContentText()) {
        var Stats = JSON.parse(result.getContentText());
        if (Stats["weight"] && Stats["weight"][0]) {
          logAllTheThings(doc, workingRow, Stats["weight"][0]);
      } else if (result.getResponseCode() === 429) {
        console.log('429 Rate limit! BREAK: ' + dateString);
        doc.getRange(workingRow, 1, 1, 37).clearContent();//clean partial row
        return 2;  //ratelimit
      }

      }
    }
    
    // SLEEP
    if(fetchNeeded(doc, LOGGABLE_SLEEP)) {
      var result = UrlFetchApp.fetch("https://api.fitbit.com/1/user/-/sleep/date/"+dateString+".json", options);
      // console.log("sleep result = " + result.getResponseCode());
      if (result.getResponseCode() === 200 && result.getContentText()) {
        // console.log("sleep ok");
        var Stats = JSON.parse(result.getContentText());
        if (Stats["sleep"] && Stats["sleep"][0]) {
          logAllTheThings(doc, workingRow, Stats["sleep"][0]);
          // console.log("sleep");
        }
        if (Stats["summary"]) {
          logAllTheThings(doc, workingRow, Stats["summary"]);
          // console.log("summary");
      } else if (result.getResponseCode() === 429) {
        console.log('429 Rate limit! BREAK: ' + dateString);
        doc.getRange(workingRow, 1, 1, 37).clearContent();//clean partial row
        return 2;  //ratelimit
      }


      }
    }

  //   //FOOD not using, decreasing API usage
  //   if(fetchNeeded(doc,LOGGABLE_FOOD)){
  //     result = UrlFetchApp.fetch(
  //       "https://api.fitbit.com/1/user/-/foods/log/date/"+dateString+".json",
  //       options
  //       );
  //       console.log("FOOD");
  //     var foodStats = JSON.parse(result.getContentText());
  //     if(!logAllTheThings(doc,workingRow,foodStats["summary"])){
  //       console.log("- food");
  //     }
  //   }

} //quota bandwidth try
catch (e) {
  
  doc.getRange(workingRow, 1, 1, 37).clearContent(); //clean partial row
  // WHATEVER EXCEPTION → CLEANUP + QUOTA MODE
  console.log('Fetch FATAL ERROR: ' + e.message);
  
  if(e.message.includes('Bandwidth quota exceeded')) {
    console.log('GAS QUOTA EXCEEDED - czekaj 24h');
  }
  
  return 99; // FATAL ERROR
}

  var timestamp = new Date(); //timestamp for every row
  doc.getRange("R"+workingRow+"C37").setValue(timestamp);
  doc.getRange("R"+workingRow+"C37").setNumberFormat("yyyy-MM-dd hh:mm:ss");
  
  return 0; // no skip
}


/****************************************

*/
function syncMany(next, end, daysProcessed = 0) {
  var props = PropertiesService.getScriptProperties();

  var totalDays = parseInt( props.getProperty('Total') );
  var next = new Date(next);
  var end = new Date(end);
  
  working('actual Period: ' + Utilities.formatDate(next, Session.getScriptTimeZone(), 'yyyy-MM-dd') + 
              ' -> ' + Utilities.formatDate(end, Session.getScriptTimeZone(), 'yyyy-MM-dd') + 
              ' total = ' + totalDays + ' days');
  
  var startTime = new Date().getTime();
  var LoopstartTime = startTime;
  var apiCalls = 0;        // Count API calls made (limit 12 for rate limiting)
  
  var wasSkipped = 0

///////////////////////////////////////////////////////
  // while (next <= end && (new Date().getTime() - startTime) < 330000 && wasSkipped < 2 ) {  
  while (next <= end 
              // && ( 360000 - ( new Date().getTime() - LoopstartTime )  ) > ( new Date().getTime() - startTime )  
        ) { 
    
    //90 apis/6min, 340 czasem Exceeded
    var LoopstartTime = new Date().getTime();

    // console.log('(while) Processing: ' + Utilities.formatDate(next, Session.getScriptTimeZone(), 'yyyy-MM-dd') 
    //       + daysProcessed
    //       + " / " + totalDays
    //       + ", APIcalls = " + apiCalls 
    //       +", working " + ( new Date().getTime() - startTime ) / 1000 + " s") ;
    
    wasSkipped = syncDate(new Date(next));

    if (wasSkipped == 0) {
      apiCalls = apiCalls + 3;
      // Utilities.sleep(500);
    } 

    now = new Date().getTime();

    if ( ( 350000 - ( now - LoopstartTime )  ) <= ( now - startTime ) ) {
      console.log( ( now - startTime ) / 1000 + " s, less then loop time ( " + ( now - LoopstartTime ) / 1000 + " s ) left to 350s, BREAK while, APIcalls = " + apiCalls )
      break;
    } else if ( wasSkipped > 1 ) { break }
     else {
      next.setDate(next.getDate() + 1);
      daysProcessed++;
    }
    
  }
  
  // working( "next = " + Utilities.formatDate(next, Session.getScriptTimeZone(), 'yyyy-MM-dd')
  //               + ", daysDone = " + daysProcessed 
  //               + ", APICalls = " + apiCalls
  //               + ", totalDays = " + totalDays );
  
  // Check if sync completed
  if (next > end) { //FINISHED
    working('FINISHED: ' + apiCalls + ' API calls (' + daysProcessed + '/' + totalDays + ' days)' );


    STOP_EXECUTION_AND_CLEAN();
    // // props.deleteProperty('Start');
    // props.deleteProperty('End');
    // // props.deleteProperty('Today');

    // props.deleteProperty('Next');
    // props.deleteProperty('Total');
    // props.deleteProperty('daysdone');
    
    // // console.log("przed pobraniem triggerow");
    // var triggers = ScriptApp.getProjectTriggers();
    // // console.log("po pobraniu triggerow");
    // triggers.forEach(function(t) {
    // if (t.getHandlerFunction() === 'resumeSync' )  {
    //     ScriptApp.deleteTrigger(t);
    //     count++;
    //   }
    // });
    // console.log("po petli");

    // ScriptApp.getProjectTriggers().forEach(function(t) {
    //   if (t.getHandlerFunction() === 'resumeSync') {
    //     ScriptApp.deleteTrigger(t);
    //   }
    // });

    return; //EXIT
    }

  //preparing to stop and resume  
  props.setProperty('Next', Utilities.formatDate(next, Session.getScriptTimeZone(), 'yyyy-MM-dd'));
  props.setProperty('daysdone', daysProcessed);

  // SET resume trigger
  var now = new Date();
  if ( wasSkipped < 2 ) {
      //Google Apps Script Limit
      var newTrigger = ScriptApp.newTrigger('resumeSync').timeBased().after(30 * 1000).create();
      // console.log('New resumeSync trigger created: ' + newTrigger.getUniqueId());

      working('AppsScript resume at : ' 
                + Utilities.formatDate(new Date(now.getTime() + 30 * 1000), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') 
                + " , day = " + Utilities.formatDate(next, Session.getScriptTimeZone(), 'yyyy-MM-dd') ) ;
      console.log( apiCalls + " API calls" +", working = " +  ( now - startTime ) / 1000 + " s");
  } else if ( wasSkipped = 2 ){
    //Fitbit RATE LIMIT

    var nextHour = new Date(now.getFullYear(), now.getMonth(), now.getDate(), now.getHours() + 1, 1, 0, 0);
    var delayMs = nextHour.getTime() - now.getTime();
    
    // console.log('Trigger at: ' + Utilities.formatDate(nextHour, Session.getScriptTimeZone(), 'HH:mm:ss'));
    console.log( "Fitbit limit: " +apiCalls + " API calls, working = " +  ( now - startTime ) / 1000 + " s")
    
    var newTrigger = ScriptApp.newTrigger('resumeSync')
      .timeBased()
      .after(delayMs)
      .create();
    
    // console.log('New resumeSync trigger created: ' + newTrigger.getUniqueId());
    working("Fitbit API Resume at " + Utilities.formatDate(nextHour, Session.getScriptTimeZone(), 'HH:mm:ss'));

  } else if ( wasSkipped == 99 ) {
    // working( "GAS quota LIMIT, run again after 24h." + apiCalls + " API calls" +", working = " +  ( now - startTime ) / 1000 + " s"  + daysProcessed + '/' + totalDays + ' days)')
    // STOP_EXECUTION_AND_CLEAN();
    // return;

    //GAS bandwidth LIMIT

    var nextHour = new Date(now.getFullYear(), now.getMonth(), now.getDate(), now.getHours() + 1, 1, 0, 0);
    var delayMs = nextHour.getTime() - now.getTime();
    
    
    console.log( "GAS bandwidth limit:" + apiCalls + "  API calls, working = " +  ( now - startTime ) / 1000 + " s")
    
    var newTrigger = ScriptApp.newTrigger('resumeSync')
      .timeBased()
      .after(delayMs)
      .create();
    
    // console.log('New resumeSync trigger created: ' + newTrigger.getUniqueId());
    working("GAS bandwidth Resume at " + Utilities.formatDate(nextHour, Session.getScriptTimeZone(), 'HH:mm:ss'));

  }

}



/**********************************************************

*/

function resumeSync() {

  // console.log("resumeSync started");

  // remove triggers resumeSync
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'resumeSync') {
      ScriptApp.deleteTrigger(t);
    }
  });

  var props = PropertiesService.getScriptProperties();

  var next = props.getProperty('Next');
  var end = props.getProperty('End');
  var total = parseInt(props.getProperty('Total') );
  var daysdone = parseInt(props.getProperty('daysdone') );
  
  console.log("Resume DATA: next = " + next
              + " , end = " + end
              + " , total = " + total
              + " , daysdone = " + daysdone
  );
    
  syncMany(next, end, daysdone);

}


/*
*******************


*/

function syncMonth() {

  console.log("sync Month query")

  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  
  // Calculate previous month
  var now = new Date();
  var prevMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  var prevMonthStr = Utilities.formatDate(prevMonth, Session.getScriptTimeZone(), 'yyyy-MM');
  
  var response = ui.prompt(
    'Year and Month', 
    'Previous month: ' + prevMonthStr + '\n\nInput Year-Month (YYYY-MM) or click OK for previous month:', 
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() == ui.Button.CANCEL) return;
  
  var input = response.getResponseText().trim();
  var year, month;
  
  // If empty or OK pressed → use previous month
  if (input === '' || input === prevMonthStr) {
    year = prevMonth.getFullYear();
    month = prevMonth.getMonth();
  } else {
    var parts = input.split('-');
    if (parts.length != 2) {
      ui.alert('Error', 'Invalid format. Use YYYY-MM.', ui.ButtonSet.OK);
      return;
    }
    year = parseInt(parts[0]);
    month = parseInt(parts[1]) - 1;
  }
  
  if (month < 0 || month > 11 || isNaN(year)) {
    ui.alert('Error', 'Invalid date.', ui.ButtonSet.OK);
    return;
  }
  
  var firstDate = new Date(year, month, 1);
  var lastDate = new Date(year, month + 1, 0);

  console.log('syncMonth variables: firstDate = ' + firstDate + " , lastDate = " + lastDate )
  
  props.setProperty('End', Utilities.formatDate(lastDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') );
  props.setProperty('Total', Math.ceil(Math.abs(lastDate.getTime() - firstDate.getTime()) / (1000 * 60 * 60 * 24)) + 1 );
  
  syncMany(firstDate, lastDate);
}


/*
*******************

*/


function syncCustom(){
  
  console.log("show form to set required day")

  var contentHTML = ''+
  '<!DOCTYPE html>'+
  '<html>'+
  ' <head>'+
  '   <script>'+
  '     function submitForm(form) {'+
  '       var formData = {'+
  '         task: "sync1day",'+
  '         year: form.year.value,'+
  '         month: form.month.value,'+
  '         day: form.day.value'+
  '       };'+
  '       google.script.run.sync1day(formData);'+
  '       google.script.host.close();' +  // ← NATYCHMIAST zamyka
  '     }'+
  '   </script>'+
  ' </head>'+
  ' <body>'+
  '   <form id="form">'+
  '     <label>Sync 1 day:</label><br>'+
  '     <input type="text" maxlength="4" size="4" id="year" name="year" value="'+(new Date().getFullYear())+'">-'+
  '     <input type="text" maxlength="2" size="2" id="month" name="month" value="'+(new Date().getMonth()+1)+'">-'+
  '     <input type="text" maxlength="2" size="2" id="day" name="day" value="'+(new Date().getDate())+'"><br><br>'+
  '     <input type="button" value="Sync & Close" onclick="submitForm(this.form)">'+
  '   </form>'+
  '</br></br><div style="text-align: right;font-style: italic;">By <a href="https://jkybett.uk" target="_blank">JKybett</a></div>' +
  ' </body>'+
  '</html>';
  
  var app = HtmlService.createHtmlOutput(contentHTML)
    .setTitle("Sync 1 Day")
    .setWidth(350)
    .setHeight(180);
  SpreadsheetApp.getUi().showModalDialog(app, "Quick Sync");
}




/*
*******************

*/


function sync1day(form) {

  var day = new Date(form.year, form.month-1, form.day);

  console.log("sync1day: " + Utilities.formatDate(day, Session.getScriptTimeZone(), 'yyyy-MM-dd'));
  
  try {      

      var result = syncDate(day);
      console.log("sync1day after syncDate" + ' , result = ' + result);
      
      switch (result) {
        case -1: 
          working("Date out of range: " + Utilities.formatDate(day, Session.getScriptTimeZone(), 'yyyy-MM-dd'));
          return
          break;
        case 0: 
          working("Synced 1 day: " + Utilities.formatDate(day, Session.getScriptTimeZone(), 'yyyy-MM-dd'));
          return
          break;
        // case false: 
        //   working("Synced 1 day: " + Utilities.formatDate(day, Session.getScriptTimeZone(), 'yyyy-MM-dd'));
        //   return
        //   break;          
        case 1: 
          working("Day already synced, clear sheet for update: " + Utilities.formatDate(day, Session.getScriptTimeZone(), 'yyyy-MM-dd'));
          return
          break;
        case 2: 
          working("Fitbit API exhausted: " + Utilities.formatDate(day, Session.getScriptTimeZone(), 'yyyy-MM-dd'));
          return
          break;
        case 99: 
          working("GAS bandwidth LIMIT, wait 24h: " + Utilities.formatDate(day, Session.getScriptTimeZone(), 'yyyy-MM-dd'));
          return
          break;        

      }
      
    } catch (error) {
      console.error("sync1day ERROR: " + error.message);
      working("ERROR: " + error.message);
      throw new Error(error.message);
    }

}

/*
*******************

*/

function syncCustomRange(){
  
  console.log("show form to set required range of days")
  
  var today = new Date();
  var startDate = new Date(today.getTime() - 30*24*60*60*1000);  // 30 dni temu
  
  // Jeśli 1szy dzień miesiąca → poprzedni miesiąc
  if(today.getDate() === 1) {
    startDate = new Date(today.getFullYear(), today.getMonth()-1, 1);
  }
  
  var endDate = new Date(startDate.getTime() + 29*24*60*60*1000);  // +29 dni
  
  var contentHTML = ''+
  '<!DOCTYPE html>'+
  '<html>'+
  ' <head>'+
  '   <script>'+
  '     function submitForm(form) {'+
  '       var formData = {'+
  '         year_start: form.year_start.value,'+
  '         month_start: form.month_start.value,'+
  '         day_start: form.day_start.value,'+
  '         year_end: form.year_end.value,'+
  '         month_end: form.month_end.value,'+
  '         day_end: form.day_end.value'+
  '       };'+
  '       google.script.run.customdates(formData);'+
  '       google.script.host.close();'+
  '     }'+
  '   </script>'+
  ' </head>'+
  ' <body>'+
  '   <form id="form">'+
  '     <label>Sync od:</label><br>'+
  '     <input type="text" maxlength="4" size="4" id="year_start" name="year_start" value="'+startDate.getFullYear()+'">-'+
  '     <input type="text" maxlength="2" size="2" id="month_start" name="month_start" value="'+(startDate.getMonth()+1)+'">-'+
  '     <input type="text" maxlength="2" size="2" id="day_start" name="day_start" value="'+startDate.getDate()+'"><br><br>'+
  '     <label>Sync do:</label><br>'+
  '     <input type="text" maxlength="4" size="4" id="year_end" name="year_end" value="'+endDate.getFullYear()+'">-'+
  '     <input type="text" maxlength="2" size="2" id="month_end" name="month_end" value="'+(endDate.getMonth()+1)+'">-'+
  '     <input type="text" maxlength="2" size="2" id="day_end" name="day_end" value="'+endDate.getDate()+'"><br><br>'+
  '     <input type="button" value="Sync Range & Close" onclick="submitForm(this.form)">'+
  '   </form>'+
  '</br></br><div style="text-align: right;font-style: italic;">By <a href="https://jkybett.uk" target="_blank">JKybett</a></div>'+
  ' </body>'+
  '</html>';
  
  var app = HtmlService.createHtmlOutput(contentHTML)
    .setTitle("Sync Date Range")
    .setWidth(380)
    .setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(app, "Quick Sync Range");
}

/*
*******************

*/

function customdates(form) {
  var startDate = new Date(form.year_start, form.month_start-1, form.day_start);
  var endDate = new Date(form.year_end, form.month_end-1, form.day_end);

  // console.log("customdates: start = " + startDate + " , end = " + endDate );

  var props = PropertiesService.getScriptProperties();
  props.setProperty('End', Utilities.formatDate(endDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') );
  props.setProperty('Total', Math.ceil(Math.abs(endDate.getTime() - startDate.getTime()) / (1000 * 60 * 60 * 24)) + 1 );
  
  syncMany(startDate, endDate);
}


// function onOpen is called when the spreadsheet is opened; adds the Fitbit menu
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var date = new Date();
  var dateString = date.getFullYear() + '-' + 
    ("00" + (date.getMonth() + 1)).slice(-2) + '-' + 
    ("00" + date.getDate()).slice(-2);
    
  var menuEntries = [{
    name: "Setup",
    functionName: "firstRun"
  }];
  
  if(getConsumerKey() != "" && getConsumerSecret() != "") {
    menuEntries = [{
      name: "Sync Month (year-month) (last month)",
      functionName: "syncMonth"
    },{
      name: "Sync custom range (last 30 days)",
      functionName: "syncCustomRange"
    },{
      name: "Sync a day (today)",
      functionName: "syncCustom"
    },{
    //   name: "test",
    //   functionName: "test"
    // }, {
      name: "Setup",
      functionName: "setup"
    }, {
      name: "Reset", 
      functionName: "clearService"
    }];
  }
  ss.addMenu("Fitbit", menuEntries);
}


