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

var eventRange = "A4:G30";
var spreadsheet = SpreadsheetApp.getActiveSheet();
var calendarId = spreadsheet.getRange("D2").getValue();
var eventCal = CalendarApp.getCalendarById(calendarId);
var now = new Date();



function sendEmail(){
  var allEvents = spreadsheet.getRange("A4:G30").getValues();
  var allEvents2 = spreadsheet.getRange("A4:G30");
  for(i=0;i < allEvents.length;i++){
    if(allEvents[i][5] == "Y" || allEvents[i][5] == "YES"){
      
      
      var recipientsTO = allEvents[i][6];
      
      var recipientsCC = "joebewildman@gmail.com";
      var formattedTime = Utilities.formatDate(allEvents[i][1], Session.getScriptTimeZone(), "EEE, MMM d, h:mm a");
      var Subject = "Event Reminder: " + allEvents[i][0] + ", " + formattedTime;
     
      var body = HtmlService.createTemplateFromFile("emailFormat");
  
      body.eventName = allEvents[i][0];
      body.eventDate = formattedTime;
      body.eventLocation = allEvents[i][3];
      MailApp.sendEmail({
        to: recipientsTO,
        cc: recipientsCC,
        subject: Subject,
        htmlBody: body.evaluate().getContent()
      });
    }
  }
  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  console.log(emailQuotaRemaining);
}

function formatSheet(){ //trying to automatically format sheet if events removed, need to figure out how to make it go all the way up without getting stuck in loop
   var allEvents = spreadsheet.getRange(eventRange).getValues();
   var allEvents2 = spreadsheet.getRange(eventRange);
  console.log("!: " + allEvents[10]);
  console.log("!!!!: " + allEvents2);
  if (allEvents[0][0] == ""){
    for(j=0;j < 7 ; j++){
       allEvents[0][j] = allEvents[1][j];
           allEvents[1][j] = "";}
    }
  
   for(i=1;i < allEvents.length;i++){
     if (allEvents[i][0] != ""&& i != 1 && allEvents[i-1][0] == ""){
       var w = i;
       while(w != 0 && allEvents[w-1][0] == ""){
         //allEvents[w-1][0] = allEvents[w][0];
         //allEvents[w][0] = "_";
         for(j=0;j < 7 ; j++){
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
  var twoWeeksBefore = new Date(now.getTime() - (14*24 * 60 * 60 * 1000));
  var oneYearBefore = new Date(now.getTime() - (180*24 * 60 * 60 * 1000)); //used to find events from one year to two weeks before now
  var events = eventCal.getEvents(oneYearBefore, twoWeeksBefore);
  for (i=0;i<events.length;i++){ //all events in cal from over two weeks ago
    for(j=0;j < allEvents.length;j++){ //all events in spreadsheet
      if(allEvents[j][0] == events[i].getTitle()){ //finds spreadsheet match from calendar event, removes spreadsheet data
        allEvents[j][0] = "";
        allEvents[j][1] = "";
        allEvents[j][2] = "";
        allEvents[j][3] = "";
        allEvents[j][4] = "";
        allEvents[j][5] = "";
        allEvents[j][6] = "";

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
    fields: "items(summary,description)",
     orderBy: "updated",
     maxResults: 500
    });//gets calendar API data of events
  //console.log(response);
   var eventsAPI = response.items; //list of event api data
   var remove = true; //boolean to determine whether or not event needs to be remmoved
  
   for(j=0;j < allEvents.length;j++){ //checks all events in spreadsheet
     
     remove = true;
     for(i=0;i<events.length;i++){ //for all current events in calendar (not removed events)
       if (events[i].getTitle() == allEvents[j][0]){ //if current event in calendar has a counterpart on spreadsheet
         remove = false; //doesn't remove
       }
     }
     if (remove == true && allEvents[j][0] != ""){ //if item in spreadsheet does not have counter part in calendar
       for(w=0;w<eventsAPI.length;w++){ //goes through all previous events, including deleted events
         //console.log("a " + allEvents[j][0] + " j " + eventsAPI[w].summary);
         if(allEvents[j][0] == eventsAPI[w].summary && eventsAPI[w].description != "AUTODEL"){ //if event in spreadsheet matches event in previous events, removes it from spreadsheet
           
           console.log(allEvents[j][0] + "deleted");
           allEvents[j][0] = "";
           allEvents[j][1] = "";//&& allEvents[j][1] == eventsAPI[w].start.dateTime
           allEvents[j][2] = "";
           allEvents[j][3] = "";
           allEvents[j][4] = "";
           allEvents[j][5] = "";
           allEvents[j][6] = "";

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
  for (i=0;i<events.length;i++){ //all events in cal from over two weeks ago
    spreadsheetVal = false;
    for(j=0;j < allEvents.length;j++){ //all events in spreadsheet
      if(allEvents[j][0] == events[i].getTitle()){ //finds spreadsheet match from calendar event, removes spreadsheet data
          var eventStart = new Date(events[i].getStartTime());
          var sheetStart = new Date(allEvents[j][1]);
          var eventEnd = new Date(events[i].getEndTime());
          var sheetEnd = new Date(allEvents[j][2]);
        if (eventStart.getTime() != sheetStart.getTime() || eventEnd.getTime() != sheetEnd.getTime()){
          //allEvents[j][1] != events[i].getStartTime() || allEvents[j][1] != events[i].getEndTime()
          events[i].setTime(allEvents[j][1],allEvents[j][2]);
          console.log("start time diff: " + allEvents[j][1] + " " + allEvents[j][2] + " " + events[i].getStartTime());
        }
        if (events[i].getLocation() != allEvents[j][3]){
          console.log("location diff: " + allEvents[j][3]);
          events[i].setLocation(allEvents[j][3]);
        }
        if (events[i].getDescription() != allEvents[j][4]){
          console.log("desc diff: " + allEvents[j][4]);
          events[i].setDescription(allEvents[j][4]);
        }
         events[i].setColor("9");
         console.log("Match found: " + allEvents[j][0]);
          spreadsheetVal = true;
          break;
        
      }
    }
    if (spreadsheetVal == false){
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
              events[i].setTag("eventId","spreadsheet");
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
      if(allEvents[i][0] == events[j].getTitle()){
        eventVal = true;
        console.log("spreadsheet check: " + allEvents[i][0] + " " + events[j].getTitle());
      }
    }
    if(allEvents[i][0] != "" && eventVal == false){
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
          event2.setColor("9");
          console.log("created new event " + event[0]);
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
}
