/**
 * @OnlyCurrentDoc
 */


function onOpen() 
{
  var configSS = SpreadsheetApp.getActive();
  //create config sheet if it doesn't exist
  var configSheet = configSS.getSheetByName("Config");
  if  (configSheet = null){ 
    configSheet = configSS.insertSheet();
    configSheet.setName("Config");
  };
  
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('SCRIPT_STARTER')
    .addItem('Click To Authorize', 'authorize')
    .addItem('Run Manually', 'ColorEvents')
    .addItem('Run Automatically every 5 minutes', 'startTimeTrigger')
    .addItem('Cancel Automatic run', 'cancelTimeTrigger')
    .addToUi();

  initializeconfigfile();

}

function initializeconfigfile (){
  var configSS = SpreadsheetApp.getActive();
  // Gets the sheet, data range, and values of the
  // spreadsheet stored in bookSS.
  var configSheet = configSS.getSheetByName("Config");
  var configRange = configSheet.getDataRange();
  var configListValues = configRange.getValues();
  configSheet.getRange(1,1,7,1).setValues([
    ["*Config File for ColorToCalendarEvents script. Recolor events on a timeframe depending on title."],
    ["*PLEASE FIRST TIME (wait 20 seconds for menu to appear): Go to SCRIPT STARTER menu and run Click to Authorize."],
    ["*On Script Starter you can run the script manually or set it to run every 5 minutes (Even if browser or sheet is closed)."],
    ["*Enter title strings to match on column C: Starting on c7 and up to c17."],
    ["*You can enter several strings for one color comma separated. No spaces before/after comma"],
    ["*Enter how many days before the moment the script is run to include. Default 7"],
    ["*Enter how many days after the moment the script is run to include. Default 7"]
    ]);
var dayfieldsrow = 6;
var firstrowtable = 10;
  if (configListValues === undefined || configListValues.length === undefined || configListValues[0].length === undefined)
  { configSheet.getRange(dayfieldsrow,6,1,1).setValues([[7]]);configSheet.getRange(dayfieldsrow+1,6,1,1).setValues([[7]]);}
  else
  { 
    //Logger.log(configListValues.length +" " + configListValues.length[0] )
  if (configListValues.length<(dayfieldsrow-1) || configListValues[0].length <6)
    { configSheet.getRange(dayfieldsrow,6,1,1).setValues([[7]]);configSheet.getRange(dayfieldsrow+1,6,1,1).setValues([[7]]);
    }
  else
  {if (configListValues[dayfieldsrow-1][5] < 0 || configListValues[dayfieldsrow-1][5] != Number)
  {configSheet.getRange(dayfieldsrow,6,1,1).setValues([[7]]);}
  if (configListValues[dayfieldsrow][5] < 0 || configListValues[dayfieldsrow][5] != Number)
  {configSheet.getRange(dayfieldsrow+1,6,1,1).setValues([[7]]);}
  }  
  }
  configSheet.getRange(dayfieldsrow,6,2,1).setBackground('#fff2cc');
  //configSheet.getRange(dayfieldsrow+2,1,1,1).setValues([["Number of events matched during the last time the script was executed"]]);


  //  Table here

  configSheet.getRange(firstrowtable,1,1,6).setValues([["Color","Color Code","Title","","","Matched"]]);
  configSheet.getRange(firstrowtable,1,1,6).setFontWeight('bold');

  configSheet.getRange(firstrowtable+1,1,11,1).setValues([["PALE_BLUE"],["PALE_GREEN"],["MAUVE"],["PALE_RED"],["YELLOW"],["ORANGE"],["CYAN"],["GRAY"],["BLUE"],["GREEN"],["RED"]]);
  configSheet.getRange(firstrowtable+1,2,11,1).setValues([["1"],["2"],["3"],["4"],["5"],["6"],["7"],["8"],["9"],["10"],["11"]]);
  configSheet.getRange(firstrowtable+1,2,11,1).setHorizontalAlignment("left");
   configSheet.getRange(firstrowtable+1,3,11,1).setBackground('#fff2cc');
    configSheet.getRange(firstrowtable+1,6,11,1).setValues([["0"],["0"],["0"],["0"],["0"],["0"],["0"],["0"],["0"],["0"],["0"]]);

}

function ColorEvents() 
{  var configSS = SpreadsheetApp.getActive();
  var configSheet = configSS.getSheetByName("Config");
  var configRange = configSheet.getDataRange();
  var configListValues = configRange.getValues();
  var firstrowtable = 10;
  var dayfieldsrow = 6;
  var bop = new Date()
  var today = new Date();
  var eop = new Date();
  bop.setDate(bop.getDate() - configListValues[dayfieldsrow-1][5]);
  eop.setDate(eop.getDate() + configListValues[dayfieldsrow][5]);
  Logger.log(bop + " >> to >> " + eop);
  var calendars = CalendarApp.getAllOwnedCalendars();
  Logger.log("found number of calendars: " + calendars.length);
//initialize counter
for (var m=firstrowtable; m<firstrowtable+11;m++) 
       {configListValues[m][5]=0;
       }

  for (var i=0; i<calendars.length; i++) 
  {
    var calendar = calendars[i];
    var nbreventsupdt = 0;
    //var calcolor = calendar.getColor();
    var calcolor = calendar;
    //Logger.log("Calendar name "+calendar.getName());
    var events = calendar.getEvents(bop, eop);
    Logger.log("Nbr events to review: " + events.length);
    for (var j=0; j<events.length; j++) 
      {
        var e = events[j];
        var title = e.getTitle();
        var organizer = e.get
        var eventcolor = -1;
      //Logger.log("Test01: " + configListValues[0][1] + " Test10: " + configListValues[1][0]);
      //Logger.log("TitleString: " + configListValues[2][2] + " Color: " + configListValues[2][1]);
      //Cycel through titles top to bottom
      for (var k=firstrowtable; k<firstrowtable+11;k++) 
       { 
          //configListValues[k][5]=0;
          //Logger.log("Clear entry " + k);
          if (typeof configListValues[k][2] !== 'undefined' && configListValues[k][2] !== "") 
          {
            mytitles = configListValues[k][2].split(",");
            var l=0;
            do {
              if (title.search(mytitles[l]) != -1 ){  
              eventcolor = configListValues[k][1];
              configListValues[k][5] = configListValues[k][5]+1;
              //Logger.log("Adding to line " + k);
              Logger.log("Match");
            }
              l++;
            } while (l < mytitles.length);
          }
       }
   
       if (eventcolor !== -1) {
         if (e.getColor() != eventcolor){
         e.setColor(eventcolor);
         }
         nbreventsupdt=nbreventsupdt+1;
       }     
      }
  }
  //configSheet.getRange(firstrowtable-2,6,1,1).setValues([[nbreventsupdt]]);
  configSheet.getRange(firstrowtable+1,6,11,1).setHorizontalAlignment("left");
  configSheet.getRange(1,1, configListValues.length, configListValues[0].length).setValues(configListValues)
}
  
function startTimeTrigger() {
  if (ScriptApp.getProjectTriggers().length===0 || ScriptApp.getProjectTriggers().length===null ){
  ScriptApp.newTrigger('ColorEvents')
  .timeBased()
  .everyMinutes(5)
  .create()
  }
}

function cancelTimeTrigger() {
   var triggers = ScriptApp.getProjectTriggers();
  
  for(var i = 0; i < triggers.length; i++){
    if(triggers[i].getTriggerSource() == ScriptApp.TriggerSource.CLOCK){
      ScriptApp.deleteTrigger(triggers[i]);
    };
  };
}

function authorize(){
  var temp1 = ScriptApp.getProjectTriggers();
  var temp2 = SpreadsheetApp.getActive();
  var temp3 = CalendarApp.getAllOwnedCalendars();
  var temp4 = temp3[0];
  var temp5 = temp4.getDescription();
}
