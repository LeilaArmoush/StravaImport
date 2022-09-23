// @ts-nocheck
const ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1uaicCeXtEvVGCw0hKcTSRRRvBJ_oQI6T6as3egPiG6Y/edit#gid=0');
var sheet = ss.getSheetByName('Sheet1'); 
var previousDate;
var totalDayMiles =0;
var completedMilesColumn;
var completedMilesRow = 2;
var timeCrossColumn;
var timeCrossRow;
var timeWorkoutColumn;
var timeWorkoutRow;
var stravaDataObject = [];
var stravaCross = [];
var stravalink = "https://www.strava.com/activities/";
var arrlink = [];
var arrlinkandname = [];
var completedColour = "#9BCEA5"
var halfCompletedColour = "#f3d271"
var quarterCompletedColour = "#e5a0a0"
var white = "#ffffff"
var today = new Date()
var wcCurrent =  (today.getMonth() + 1) + '/' + (today.getDate() - today.getDay() + 1) + '/' + today.getFullYear()
var sheetData = sheet.getDataRange();
var todayFormatted = (today.getMonth() + 1) + '/' + today.getDate() + '/' + today.getFullYear()

function onOpen()
{
  dateCells().filter(x => x[0] === todayFormatted).forEach(function(x)
  {
    sheetData.getCell(x[2],x[1]).activate();
  })
}
   function importStravaData()
   {
       getStravaActivityData().forEach(function(activity)
           { 
            var arr6 = [];
           
            var arrayCross = [];
            var date = (activity[0].getMonth() + 1) + '/' + activity[0].getDate() + '/' + activity[0].getFullYear()
            if(activity[2] !== 'Run'&& activity[2] !== 'Weight Training')
            {
                dateCells().forEach(function(y){
                if(date === y[0])
                {
                timeCrossColumn = parseInt(y[1]);
                timeCrossRow = parseInt(y[2] + 7);
                }
              }) 
              arrayCross.push(
              date, 
              activity[5],
              activity[2],
              timeCrossColumn,
              timeCrossRow,
              activity[3],
              activity[4]
              )
              stravaCross.push(arrayCross);
            }
            Logger.log(stravaCross);
             if(activity[2] === 'Run')
            {
           if(previousDate !== date)
            {
               totalDayMiles = parseFloat(activity[1]);
               arrlink = [];
            }
             arrlink.push(
               '=HYPERLINK(' + '"' + stravalink + activity[3] + '", "' + activity[4] + '")'
               );    

              if(previousDate === date)
              {
               totalDayMiles = parseFloat(totalDayMiles) +  parseFloat(activity[1]);
              }
              
              dateCells().forEach(function(y){
                if(date === y[0])
                {
                completedMilesColumn = parseInt(y[1]);
                completedMilesRow = parseInt(y[2]);
                }
                }) 
            arr6.push(
              date, 
              totalDayMiles,
              activity[2],
              completedMilesColumn,
              completedMilesRow,
              arrlink
              )
              if(previousDate !== date)
              {
              stravaDataObject.push(arr6);  
              }
              if(previousDate === date)
              {
               stravaDataObject.fill(arr6, stravaDataObject.length - 1);
              }
            }
             previousDate = date;
          });  
          Logger.log(stravaDataObject); 
           
            stravaCross.forEach(function(b)
              {
              if(b[4] === null || b[3])
              {
              var secondsInFiveMins = 300;
              var elapsedTime = new Date(b[1] * 1000).toISOString().slice(11, 19);
              Logger.log(elapsedTime);
              var elapsedTimeArray = elapsedTime.split(":");
              var plannedMilesTimeInSeconds = parseInt(elapsedTimeArray[0])*3600 + parseInt(elapsedTimeArray[1])*60 + parseInt(elapsedTimeArray[2]);
              Logger.log(plannedMilesTimeInSeconds);
              var linkToCrossStrava = '=HYPERLINK(' + '"' + stravalink + b[5] + '", "' + b[6] + '")';    
              sheetData.getCell(b[4]-1,b[3]).setValue(elapsedTime);
              sheetData.getCell(b[4],b[3]-2,).setValue(linkToCrossStrava);
              }
              if(parseFloat(b[1]) >= (plannedMilesTimeInSeconds - secondsInFiveMins))
              {
                sheet.getRange(b[4]-1,b[3]-2,2, 2).setBackground(completedColour);
              }
              else
              if(parseFloat(b[1]) >= (plannedMilesTimeInSeconds-(plannedMilesTimeInSeconds*2)))
              {
                sheet.getRange(b[4],b[3]-2,1, 2).setBackground(halfCompletedColour);
              }
              else
              if(parseFloat(b[1]) <= (plannedMilesTimeInSeconds*(plannedMilesTimeInSeconds*2)))
              {
                sheet.getRange(b[4],b[3]-2,1, 2).setBackground(quarterCompletedColour);
              }
              else
             {
               sheet.getRange(b[4],b[3]-2,1, 2).setBackground(white);
             } 
            });
             
          stravaDataObject.forEach(function(x)
          {
            var plannedMilesOneCell = sheetData.getCell(x[4],x[3]-1)
            var plannedMilesTwoCell = sheetData.getCell((x[4]+1),(x[3]-1))
            var plannedMilesOneCellValue = parseFloat(plannedMilesOneCell.getValue())
            var plannedMilesTwoCellValue = parseFloat(plannedMilesTwoCell.getValue())
            var completedMilesCell = sheetData.getCell(x[4],x[3])
            if(isNaN(plannedMilesOneCellValue)) plannedMilesOneCellValue = 0;
            if(isNaN(plannedMilesTwoCellValue)) plannedMilesTwoCellValue = 0;
            if(isNaN(parseFloat(plannedMilesTwoCell.getValue()))) plannedMilesTwoCell = 0;
            var plannedMiles = plannedMilesOneCellValue + plannedMilesTwoCellValue;
            sheetData.getCell(x[4],x[3]).setValue(Math.round(x[1],4));
            var z=0;
            (x[5]).forEach(function(link)
            { if(z<=3)
            {
              sheetData.getCell(x[4]+2+z,x[3]-1).setValue(link);
            }
              z++;
            }
            )    
            //Logger.log(parseFloat(plannedMilesOneCell.getValue()))
           
            if(isNaN(plannedMiles)|| plannedMiles === 0)
            {
              sheet.getRange(x[4],(x[3]-2),2,3).setBackground(white)
            }
            else
              if((plannedMiles - 0.25)<= parseFloat(x[1]))
            {
               sheet.getRange(x[4],(x[3]-2),2,3).setBackground(completedColour)
            } 
            else
            if(plannedMiles >= (parseFloat(x[1]))/2)
            {
              sheet.getRange(x[4],(x[3]-2),2,3).setBackground(halfCompletedColour)
            }
            else
            if(plannedMiles >= (parseFloat(x[1]))/4)
            {
              sheet.getRange(x[4],(x[3]-2),2,6).setBackground(quarterCompletedColour)
            }
          }) 
  }   

function dateCells()
{
 july27WeekNumber = 30;
   
 Date.prototype.getWeekNumber = function(){
  var d = new Date(Date.UTC(this.getFullYear(), this.getMonth(), this.getDate()));
  var dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  var yearStart = new Date(Date.UTC(d.getUTCFullYear(),0,1));
  return Math.ceil((((d - yearStart) / 86400000) + 1)/7)
};
  var weekOfYearToday = new Date().getWeekNumber();
  var weekRowToday =  2 + ((weekOfYearToday - july27WeekNumber)*8);
  var columnToday;
  if(new Date().getDay() === 0)
  {
    columnToday = 22;
  }
  else
  {
   columnToday = 1 + ((new Date().getDay())*3);
  }
  var cellDatesObject = [];

const today = new Date()
const yesterday = new Date(today.setDate(today.getDate() - 1));
const yesterdayFormatted = (yesterday.getMonth() + 1) + '/' + yesterday.getDate() + '/' + yesterday.getFullYear();
const weekOfYearYesterday = yesterday.getWeekNumber();
var weekRowYesterday = parseInt(2 + ((weekOfYearYesterday - july27WeekNumber)*8));
var columnYesterday = parseInt(1 + ((yesterday.getDay())*3));

  cellDatesObject.push(
    [todayFormatted,
    columnToday,
    weekRowToday]
  )

  cellDatesObject.push(
    [yesterdayFormatted,
    columnYesterday,
    weekRowYesterday]
  )
 return cellDatesObject;
};

/*function dateCells(){
    var sheetData = sheet.getDataRange();
    var z=2;
    var wcRow = 0;
    var rowBool = true;
     while(rowBool)
     {
       var wcToday = sheetData.getCell(z,1).getValue();
       var wcTodayDate = new Date(wcToday);
       var wcDateFormatted = (wcTodayDate.getMonth() + 1) + '/' + wcTodayDate.getDate() + '/' + wcTodayDate.getFullYear();
        if(wcDateFormatted === wcCurrent)
        {
            wcRow = z;
            break;
        }
       z=z+8;
     }
   var sheetRow = 2;
   var sheetColumn = 4;
   var k = 0;
  
   var cellDatesObject = [];

   while(sheetRow<43)
   {
      var cellDates = [];
          var sheetData = sheet.getDataRange();
          try{
          var wc = sheetData.getCell(2,1).getValue();
          }
          catch(err)
          {
            return err;
          }

          weekCommencing = new Date(wc);
          var cellDateOfWeek = new Date(weekCommencing.setDate(weekCommencing.getDate() + k))
         
          var dateOfWeekFormatted = (cellDateOfWeek.getMonth() + 1) + '/' + cellDateOfWeek.getDate() + '/' + cellDateOfWeek.getFullYear();

          cellDates.push
          (dateOfWeekFormatted,
          sheetColumn,
          sheetRow);

          cellDatesObject.push(cellDates);

          if(sheetColumn>22)
          {
           sheetColumn === 4;
           sheetRow = sheetRow + 8
          }
          else
          {
          sheetColumn = sheetColumn + 3;
          }
          k++
   }
  // Logger.log(cellDatesObject);
   return cellDatesObject;
} */
// custom menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
 
  ui.createMenu('Strava App')
    .addItem('Get data', 'getStravaActivityData')
    .addToUi();
}

// Get athlete activity data
function getStravaActivityData() {
  
 
  // call the Strava API to retrieve data
  var data = callStravaAPI();
   
  // empty array to hold activity data
  var stravaData = [];
     
  // loop over activity data and add to stravaData array for Sheet
  data.forEach(function(activity) {
    var start_date = new Date(activity.start_date_local)
    start_date.toLocaleDateString();
    var arr = [];
    arr.push(
      start_date,
      activity.distance/1609.344,
      activity.type,
      activity.id,
      activity.name,
      activity.elapsed_time
    );
    stravaData.push(arr);
  });
  // paste the values into the Sheet
  return stravaData;
}
 
// call the Strava API
function callStravaAPI() {
   
  // set up the service
  var service = getStravaService();
   
  if (service.hasAccess()) {
   
    var oneDayInSeconds = 86400;
    var unixTime = BigInt(Math.floor((today.setHours(0,0,0,0) / 1000)));
    
    var endpoint = 'https://www.strava.com/api/v3/athlete/activities';
    var params = '?after=' + unixTime +'&per_page=200';
     Logger.log('App has access to endpoint:' + endpoint);
 
    var headers = {
      Authorization: 'Bearer ' + service.getAccessToken()
    }; 
    var options = {
      headers: headers,
      method : 'GET',
      muteHttpExceptions: true,
    };
     
    var response = JSON.parse(UrlFetchApp.fetch(endpoint + params, options));
     
    return response;  
  }
  else {
    Logger.log("App has no access yet.");
     
    // open this url to gain authorization from github
    var authorizationUrl = service.getAuthorizationUrl();
     
    Logger.log("Open the following URL and re-run the script: %s",
        authorizationUrl);
  }
}       
