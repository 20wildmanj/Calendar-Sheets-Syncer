  // 8/27/19: currently creates calendar event from spreadsheet, no description
  //Initial code derived from: https://www.youtube.com/watch?v=MOggwSls7xQ
  //To do: remove previous calendar entries, then add new ones so no overlap, add description
  //by Joey Wildman
  //calendar being used is here: https://calendar.google.com/calendar/b/1?cid=YXNpai5hYy5qcF9mdWw2dGRpdGw2M21oMXZqbW9qOThvMzJvNEBncm91cC5jYWxlbmRhci5nb29nbGUuY29t
  // 8/28/19: Added description, added location, deletes previous events created by script while keeping manually inputted events, changed color(JW)
  //Trying to add sync, but still difficult
  //Added Sync capabilities, takes user inputted info from calendar and puts it into spreadsheet, then re-sends it as a spreadsheet event (JW)
  //Added function to remove outdate events, might be changed if a log of previous events is needed, began to make a formatting function (JW)
  //Need to add ability to detect if user deleted event and sync that to spreadsheet
  //8/29/19 started using google calendar api to search through deleted events in order to find manually deleted events (JW)
  //8/30/19 started making function to remove manual deletes from spreadsheet, not working,deleted everything lol (JW)
  //9/1/19 redid function to remove manual calendar deletes, now works properly, however one caveat is that if you use try to put the event back in the spreadsheet with the same title it will be removed (JW)
  //Originally I tried tagging every event created by script by replacing the description to Autodelete before it was delete so when the API was accessed the autodeleted events could be distinguished from the manual removals
  //However, since manually created events wouldn't have that tag in the beginning, the function would delete it anyways
  //the new solution I came up with seems much simpler, but still has the issue about identical event titles, but that can be somewhat fixed by looking at time signatures instead
  //9/1/19 update 2: just found that there are too many events cluttering up the API, which screwed up manual delete since when retrieving API data not all events were returned
//To do: optimize code so that auto deleting events and then re adding them to calendar no longer happens, need to split up event creation into already created spreadsheet events, already created calendar events, and new spreadsheet events
//9/2/19 the program is now optimized in terms of deleting and readding everything, seems to work fine but I will do some more debugging to ensure it works, also started using github
//fixed some bugs involving events not deleting and some events being incorrectly deleted
//built in edit-sync since new code works differently
//completed formatting function that automatically gets rid of empty row between rows filled in with events, just for ease of use I guess
//To do: Rigorous testing, email reminders with seperate array
//9/7/19 added basic email sending, to do: send day before only once, too tired now
//9/10/19 everything is on github, has been for a while so less updates here

var eventRange = "A4:H30";
var spreadsheet = SpreadsheetApp.getActiveSheet();
var calendarId = spreadsheet.getRange("D2").getValue();
var eventCal = CalendarApp.getCalendarById(calendarId);
var now = new Date();


function postMessageToDiscord(message) {

  message = message || "Error: Message not found";
  
  var discordUrl = 'https://discordapp.com/api/webhooks/622320640795344896/kqhn0_ibcHlI0xFcago-sQlMdPlosHx4g_WsZojuZLI53yLpBWCF2fvRrj_mYnRsX2T9';
  var payload = JSON.stringify({content: message});
  
  var params = {
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded'
    },
    method: "POST",
    payload: payload,
    muteHttpExceptions: true
  };
  
  var response = UrlFetchApp.fetch(discordUrl, params);
  
  Logger.log(response.getContentText());

}

function sendEmail2(){//fixed email function
  var allEvents = spreadsheet.getRange(eventRange).getValues();
  var allEvents2 = spreadsheet.getRange(eventRange);
  var now = new Date();
  var TwoDaysFromNow = new Date(now.getTime() + (3*24 * 60 * 60 * 1000)); //used to find events from now to 1 year in the future
  var events = eventCal.getEvents(new Date(now.getTime()), TwoDaysFromNow);
  
    if (events.length > 0){
      for (i=0;i<events.length;i++){ //all events in cal from over two weeks ago
        for(j=0;j < allEvents.length;j++){ //all events in spreadsheet
          if(allEvents[j][7] == events[i].getTag("identifier")){ //finds spreadsheet match from calendar event, removes spreadsheet data
              var timeDiff = (events[i].getStartTime() - now.getTime())/(2*24*60*60*1000);
              if (events[i].getTag("email") == "NO" && timeDiff <= 1.0){
                var recipientsTO = allEvents[j][6];
                var mailArray = recipientsTO.split(",");
                console.log(mailArray);
                var recipientsCC = "";
                var formattedStartTime = Utilities.formatDate(allEvents[j][1], Session.getScriptTimeZone(), "EEE, MMM d, h:mm a");
                var formattedEndTime = Utilities.formatDate(allEvents[j][2], Session.getScriptTimeZone(), "h:mm a");
                if ( events[i].getTag("update") == "YES" && allEvents[j][4] == "CANCELLED"){
                  var Subject = "Event Canceled: " + allEvents[j][0] + ", " + formattedStartTime;
                  var header = "An upcoming event has been CANCELED:";
                }
                else if ( events[i].getTag("update") == "YES"){
                  var Subject = "Event Updated: " + allEvents[j][0] + ", " + formattedStartTime;
                  var header = "An upcoming event has been updated:";
                  
                  events[i].setTag("update","NO");
                } else{
                  var Subject = "Event Reminder: " + allEvents[j][0] + ", " + formattedStartTime;
                  var header = "This is a reminder of an upcoming event:";
                  
                }
                var body = HtmlService.createTemplateFromFile("emailFormat");
                
                body.eventName = allEvents[j][0];
                body.header = header;
                body.eventStartDate = formattedStartTime;
                body.eventEndDate = formattedEndTime;
                body.eventLocation = allEvents[j][3];
                body.description = allEvents[j][4];
                console.log("EVENT MAILED " + body.eventName);
                
                for (h = 0; h < mailArray.length; h++)
                {
                  body.email = mailArray[h];
                  MailApp.sendEmail({
                    to: mailArray[h],
                    cc: recipientsCC,
                    subject: Subject,
                    htmlBody: body.evaluate().getContent()
                  });
                  
                }
                postMessageToDiscord(Subject + "-" + formattedEndTime);
                events[i].setTag("email","YES");
              }
          }
        }
      }
    
  }
   var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  console.log(emailQuotaRemaining);
}



function formatSheet(){ //trying to automatically format sheet if events removed, need to figure out how to make it go all the way up without getting stuck in loop
   var allEvents = spreadsheet.getRange(eventRange).getValues();
   var allEvents2 = spreadsheet.getRange(eventRange);

  if (allEvents[0][0] == ""){
    for(j=0;j < 8 ; j++){
       allEvents[0][j] = allEvents[1][j];
           allEvents[1][j] = "";}
    }
  
   for(i=1;i < allEvents.length;i++){
     if (allEvents[i][0] != ""&& i != 1 && allEvents[i-1][0] == ""){
       var w = i;
       while(w != 0 && allEvents[w-1][0] == ""){
         //allEvents[w-1][0] = allEvents[w][0];
         //allEvents[w][0] = "_";
         for(j=0;j < 8 ; j++){
         allEvents[w-1][j] = allEvents[w][j];
           allEvents[w][j] = "";}
         w--;
         
       }
     }
     }
     allEvents2.setValues(allEvents);

  }
 


function removeOutdatedEvents(){ //removes any outdated events from spreadsheet and calendar
  var allEvents = spreadsheet.getRange(eventRange).getValues(); //gets event data from spreadsheet
  var allEvents2 = spreadsheet.getRange(eventRange);//gets event data format so changed values can be put back in the spreadsheet
  var now = new Date();
  var twoWeeksBefore = new Date(now.getTime() - (7*24 * 60 * 60 * 1000));
  var oneYearBefore = new Date(now.getTime() - (180*24 * 60 * 60 * 1000)); //used to find events from one year to two weeks before now
  var events = eventCal.getEvents(oneYearBefore, twoWeeksBefore);
  for (i=0;i<events.length;i++){ //all events in cal from over two weeks ago
    for(j=0;j < allEvents.length;j++){ //all events in spreadsheet
      if(allEvents[j][7] == events[i].getTag("identifier")){ //finds spreadsheet match from calendar event, removes spreadsheet data
        allEvents[j][0] = "";
        allEvents[j][1] = "";
        allEvents[j][2] = "";
        allEvents[j][3] = "";
        allEvents[j][4] = "";
        allEvents[j][5] = "";
        allEvents[j][6] = "";
        allEvents[j][7] = "";

        events[i].setDescription("AUTODEL");
          events[i].deleteEvent(); //removes event to complete removal
          break;
        console.log("outdated: " + allEvents[j][0]);
    }
  }
}
  allEvents2.setValues(allEvents);
}

function removeManualDeletedEvents(){
    var oneYearBefore = new Date(now.getTime() - (80*24 * 60 * 60 * 1000)); //used to find events from one year to two weeks before now
    var oneYearFromNow = new Date(now.getTime() + (360*24 * 60 * 60 * 1000)); //used to find events from now to 1 year in the future
    var events = eventCal.getEvents(oneYearBefore, oneYearFromNow);
  
   var allEvents = spreadsheet.getRange(eventRange).getValues(); //gets event data from spreadsheet
   var allEvents2 = spreadsheet.getRange(eventRange); //gets event data format so changed values can be put back in the spreadsheet 
   
    var response = Calendar.Events.list( 
   calendarId, {
    showDeleted: true,
    fields: "items(summary,description,extendedProperties)",
     orderBy: "updated",
     maxResults: 500,
     
    });//gets calendar API data of events
  //console.log(response);
   var eventsAPI = response.items; //list of event api data
   var remove = true; //boolean to determine whether or not event needs to be remmoved
  
   for(j=0;j < allEvents.length;j++){ //checks all events in spreadsheet
     
     remove = true;
     
     for(i=0;i<events.length;i++){ //for all current events in calendar (not removed events)
       //console.log("row " + j + " " + allEvents[j][7] +  " " + events[i].getId() );
       if (events[i].getTag("identifier") == allEvents[j][7]){ //if current event in calendar has a counterpart on spreadsheet
         remove = false; //doesn't remove
       }
     }
     if (remove == true && allEvents[j][0] != ""){ //if item in spreadsheet does not have counter part in calendar
       for(w=0;w<eventsAPI.length;w++){ //goes through all previous events, including deleted events
         //console.log("a " + allEvents[j][0] + " j " + eventsAPI[w].summary);
         //console.log(eventsAPI[w].extendedProperties.shared.identifier);
         if(allEvents[j][7] == eventsAPI[w].extendedProperties.shared.identifier && eventsAPI[w].description != "AUTODEL"){ //if event in spreadsheet matches event in previous events, removes it from spreadsheet
           
           console.log(allEvents[j][0] +  " deleted " + allEvents[j][7] + " eventId " + eventsAPI[w].extendedProperties.shared.identifier);
           allEvents[j][0] = "";
           allEvents[j][1] = "";//&& allEvents[j][1] == eventsAPI[w].start.dateTime
           allEvents[j][2] = "";
           allEvents[j][3] = "";
           allEvents[j][4] = "";
           allEvents[j][5] = "";
           allEvents[j][6] = "";
           allEvents[j][7] = "";
           

         }
       }
     }
   }
  
  //var apiDate = Moment.moment(eventsAPI[w].start.dateTime);
         //apiDate.format('YYYY-MM-DDTHH:mm:ss');
         //var formattedDate = apiDate.toDate();
         //console.log(allEvents[j][0]);
         //console.log(formattedDate);
  

  
   allEvents2.setValues(allEvents);
}


function guidGenerator() { //random generator for id
    var S4 = function() {
       return (((1+Math.random())*0x10000)|0).toString(16).substring(1);
    };
    return (S4()+S4()+"-"+S4()+"-"+S4()+"-"+S4()+"-"+S4()+S4()+S4());
}

function sheetsToCalendar() { 
  formatSheet();
  removeOutdatedEvents();
  removeManualDeletedEvents();
   
  
  
  var allEvents = spreadsheet.getRange(eventRange).getValues();
  var allEvents2 = spreadsheet.getRange(eventRange);

  var now = new Date();
  var oneYearBefore = new Date(now.getTime() - (80*24 * 60 * 60 * 1000)); //used to find events from one year to two weeks before now
  var oneYearFromNow = new Date(now.getTime() + (360*24 * 60 * 60 * 1000)); //used to find events from now to 1 year in the future
  var events = eventCal.getEvents(oneYearBefore, oneYearFromNow);
  
  var spreadsheetVal = false;
  var eventVal = false;
  var eventUpdate = false;
  for (i=0;i<events.length;i++){ //all events in cal from over two weeks ago
    spreadsheetVal = false;
    eventUpdate = false;
    for(j=0;j < allEvents.length;j++){ //all events in spreadsheet
      
            //console.log("event list row: " + i + " " + events[i].getId());

      if(allEvents[j][7] == events[i].getTag("identifier") && allEvents[j][7] != ""){ //finds spreadsheet match from calendar event, updates spreadsheet data
          var eventStart = new Date(events[i].getStartTime());
          var sheetStart = new Date(allEvents[j][1]);
          var eventEnd = new Date(events[i].getEndTime());
          var sheetEnd = new Date(allEvents[j][2]);
        
        if (eventStart.getTime() != sheetStart.getTime() || eventEnd.getTime() != sheetEnd.getTime()){
          //allEvents[j][1] != events[i].getStartTime() || allEvents[j][1] != events[i].getEndTime()
          events[i].setTime(sheetStart,sheetEnd);
          eventUpdate = true;
          console.log("start time diff: " + allEvents[j][1] + " " + allEvents[j][2] + " " + events[i].getStartTime());
          console.log("event " + eventStart.getTime() + " end " + eventEnd.getTime());
          console.log("spread " + sheetStart.getTime() + " end " + sheetEnd.getTime());
        }
        if (events[i].getLocation() != allEvents[j][3]){
          console.log("location diff: " + allEvents[j][3]);
          events[i].setLocation(allEvents[j][3]);
          eventUpdate = true;
        }
        if (events[i].getDescription() != allEvents[j][4]){
          console.log("desc diff: " + allEvents[j][4]);
          events[i].setDescription(allEvents[j][4]);
          eventUpdate = true;
        }
         events[i].setColor("9");
        if (eventUpdate){
        events[i].setTag("update","YES");
        events[i].setTag("email","NO");
        }
         console.log("Match found: " + allEvents[j][0]);
          spreadsheetVal = true;
          break;
        
      }
    }
    if (spreadsheetVal == false){ //if calendar event not in spreadsheet
      console.log("spreadsheet val false row: " + i + " " + events[i].getId());
      if (events[i].getTag("eventId") == "spreadsheet"){
          events[i].setDescription("AUTODEL");
          events[i].deleteEvent();
      } else {
        for (j=0;j<allEvents.length;j++){ //maybe check vals later
          try{
            if (allEvents[j][0] == ""){  //finds first empty row
              allEvents[j][0] = events[i].getTitle();
              allEvents[j][1] = events[i].getStartTime();
              allEvents[j][2] = events[i].getEndTime();
              allEvents[j][3] = events[i].getLocation();
              allEvents[j][4] = events[i].getDescription();

              var unique = guidGenerator();
              events[i].setTag("identifier",unique);
              allEvents[j][7] = unique;
              
              events[i].setTag("eventId","spreadsheet");
              events[i].setTag("email","NO");
              console.log("New Event Added from calendar: " + allEvents[j][0]);
              //events[i].setDescription("AUTODEL");
              //events[i].deleteEvent(); //removes event so not duplicated when spreadsheet events are sent to calendar
              break;
            }
            
          }  catch(e){
          console.error('new cal sync yielded an error: ' + e);
        }
      }
      
    }
  }
}
  for (i=0; i<allEvents.length;i++){
    eventVal = false;
    for(j=0; j<events.length;j++){
      console.log("spreadsheet check: " + allEvents[i][7] + " " + events[j].getTag("identifier"));
      if(allEvents[i][7] == events[j].getTag("identifier")){
        eventVal = true;
        console.log("spreadsheet YES: " + allEvents[i][7] + " " + events[j].getTag("identifier"));
      }
    }
    if(allEvents[i][0] != "" && eventVal == false){ //if new event from spreadsheet
      console.log("spreadsheetEvent found: " +  allEvents[i][0]);
        try{
          var event = allEvents[i];
          var eventTitle = event[0];
          var startTime = event[1];
          var endTime = event[2];
          var location = event[3]
          var notes = event[4];
          var event2 = eventCal.createEvent(eventTitle,startTime,endTime, {description: notes});
          event2.setLocation(location);
          event2.setTag("eventId","spreadsheet");
          event2.setTag("email","NO");
          event2.setColor("9");
          var unique = guidGenerator();//new Date(now.getTime());
          event2.setTag("identifier", unique);
          allEvents[i][7] = unique;
          console.log("created new event " + event[0]);
          console.log("row: " + i + " " + event2.getId());
        }
      catch(e){
        console.error('new sheet event sync yielded an error: ' + e);
      }
    }
  
  }


    allEvents2.setValues(allEvents);

  
  
    formatSheet();
  
  
 

}

function onOpen(){ //creates button next to help that runs the function without needing to open script editor
var ui = SpreadsheetApp.getUi();
ui.createMenu("Sync")
  .addItem( "Sync calendar","sheetsToCalendar")
.addToUi();
ui.createMenu("Reminder")
  .addItem( "Manually send reminder","sendEmail2")
.addToUi();
}
