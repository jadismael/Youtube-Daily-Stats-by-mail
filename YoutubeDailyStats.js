/**
* Ido's YouTube Dashboard Stats example
* Fetch stats on videos and channel by using YT APIs.
  1. It using the ATOM feeds - v1.
  2. For the channel we are using the v3 version of the API.
* @Author: Jad Ismael
* @Date: April 2015
* @Website:UltGate.com
*/

/**
 *This script has two main features: 
 *the first generates a report, about your youtube uploads with Video Title, Durations, Views, ID and upload date as data
 *the second sends your stats as an email to you.
*/

/**
 * This script is a an improvemenet to the original script by Ido Green 
 * Original file: https://github.com/greenido/AppsScriptBests/blob/master/public_html/AppsScript/YouTube/ytStats.js
 */

function onOpen() {

 var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{ name : "Update Stats", functionName : "fetchAllData"},
                { name : "Send Report Now", functionName : "EmailReport"}
              
                ];
  spreadsheet.addMenu("Youtube Dash", entries);
}

function retrieveMyUploads() {
  var jad =2;
  var ids = [];  
  var titles = [];   
    var ss = SpreadsheetApp.getActiveSheet();
  var results = YouTube.Channels.list('contentDetails', {mine: true});
  for(var i in results.items) {
    var item = results.items[i];
    // Get the playlist ID, which is nested in contentDetails, as described in the
    // Channel resource: https://developers.google.com/youtube/v3/docs/channels
    var playlistId = item.contentDetails.relatedPlaylists.uploads;

    var nextPageToken = '';

    // This loop retrieves a set of playlist items and checks the nextPageToken in the
    // response to determine whether the list contains additional items. It repeats that process
    // until it has retrieved all of the items in the list.
    while (nextPageToken != null) {
      var playlistResponse = YouTube.PlaylistItems.list('snippet', {
        playlistId: playlistId,
        maxResults: 50,
        pageToken: nextPageToken
      });

      for (var j = 0; j < playlistResponse.items.length; j++) 
      {
        var playlistItem = playlistResponse.items[j];
        Logger.log('[%s] Title: %s',
                   playlistItem.snippet.resourceId.videoId,
                   playlistItem.snippet.title);
        
        
        ids[ids.length] = playlistItem.snippet.resourceId.videoId; 
       titles[titles.length] =playlistItem.snippet.title;
        
        ss.getRange("C" + jad).setValue(playlistItem.snippet.resourceId.videoId);
     //   ss.getRange("D" + jad).setValue(playlistItem.snippet.description);
        
        temp = new Date(playlistItem.snippet.publishedAt);
       var formateddate=  Utilities.formatDate(temp, "GMT", "yyyy-MM-dd")
           ss.getRange("D" + jad).setValue(formateddate);
  ss.getRange("B" + jad).setValue(playlistItem.snippet.title);
        
 
        jad=jad+1;
       
      }
      
      
      
      
    
      
      
     // ss.getRange("C:C").setValues( ids);
  
      
      nextPageToken = playlistResponse.nextPageToken;
    }
  }
}






function fetchAllData() {
  
  
  

  
  retrieveMyUploads();
  var start = new Date().getTime();
  
  var curSheet = SpreadsheetApp.getActiveSheet();
  
    
      
  var titlerange= curSheet.getRange("B1");
  var idrange= curSheet.getRange("C1");
  var uploaddaterange= curSheet.getRange("D1");
  var totalviewsrange= curSheet.getRange("E1");
  var durationrange= curSheet.getRange("F1");
  
    var emailrange= curSheet.getRange("H3");
  var emaildata=emailrange.getValue();
  if(emaildata[0][0]='') emailrange.setValue("Write your email here to receve the report");

  
   titlerange.setValue("Title"); 
    idrange.setValue("Video ID");  
uploaddaterange.setValue("Upload Date") ;  
    durationrange.setValue("Duration");
  
  titlerange.setFontWeight('bold');  idrange.setFontWeight('bold');
  uploaddaterange.setFontWeight('bold');
  totalviewsrange.setFontWeight('bold');
 durationrange.setFontWeight('bold');

  curSheet.setFrozenRows(1);
  
 
  
  var ytIds = curSheet.getRange("C:C");
  var totalRows = ytIds.getNumRows();
  var ytVal = ytIds.getValues();
  var errMsg = "<h4>Errors:</h4> <ul>";
  // let's run on the rows after the header row
  for (var i = 1; i <= totalRows - 1; i++) {
    // e.g. for a call: https://gdata.youtube.com/feeds/api/videos/YIgSucMNFAo?v=2&prettyprint=true
    if (ytVal[i] == "") {
      Logger.log("We stopped at row: " + (i+1));
      break;
    }
    var link = "https://gdata.youtube.com/feeds/api/videos/" + ytVal[i] + "?v=2&prettyprint=true";
    try {
      fetchYTdata(link, i+1);
    }
    catch (err) {
      errMsg += "<li>Line: " + i + " we could not fetch data for ID: " + ytVal[i] + "</li>";
      Logger.log("*** ERR: We have issue with " + ytVal[i] + " On line: " + i);
    }
  }
  if (errMsg.length < 24) {
    // we do not have any errors at this run
    errMsg += "<li> All good for now </li>";
  }
  var end = new Date().getTime();
  var execTime = (end - start) / 1000;
}



function fetchYTdata(url, curRow) {
   //var url = 'https://gdata.youtube.com/feeds/api/videos/Eb7rzMxHyOk?v=2&prettyprint=true';
   var rawData = UrlFetchApp.fetch(url).getContentText();
   //Logger.log(rawData);
                           
  // published <published>2014-05-09T06:22:52.000Z</published>
   var inx1 = rawData.indexOf('published>') + 10;
   var inx2 = rawData.indexOf("T", inx1);
   var publishedDate = rawData.substr(inx1, inx2-inx1);
  
   // viewCount='16592'
   var inx1 = rawData.indexOf('viewCount') + 11;
   var inx2 = rawData.indexOf("'/>", inx1);
   var totalViews = rawData.substr(inx1, inx2-inx1);
  
   // <yt:duration seconds='100'/>
   var inx1 = rawData.indexOf('duration seconds') + 18;
   var inx2 = rawData.indexOf("'/>", inx1);
   var durationSec = rawData.substr(inx1, inx2-inx1);
  
   Logger.log(curRow + ") TotalViews: " + totalViews + " durationSec: " + durationSec);
   
  // update the sheet
  var ss = SpreadsheetApp.getActiveSheet();
  ss.getRange("E" + curRow).setValue(totalViews);
  ss.getRange("F" + curRow).setValue(durationSec);
  
 }


function EmailReport() {
  fetchAllData(); //updating data before sending email
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;  // First row of data to process
  var numRows = 50;   // Number of rows to process
  var dataRange = sheet.getRange("B:F");
  var message="<table>";
  var data = dataRange.getValues();
  for (i in data) {
    var row = data[i];
  
    
  var emailrange= sheet.getRange("H3");
      var emailAddress = emailrange.getValue();  // First column
    
      
    if (typeof row[1] == 'string' && row[1]!="") 
    {
    message = message 
    + "<tr><td>"
     +  row[0] 
     + "</td><td>"
      + row[1]  
       + "</td><td>"
      + row[3] 
       + "</td><td>"
      + row[4]  
      + "</td></tr>";
    }
  
  
  }
  message= message + "</table>";
  
  
  var d = new Date();
var day = d.getDate();  //day
  var month = d.getMonth();//month
  var year = d.getFullYear(); //year
 
      var subject = "Youtube Daily Summary by Jed Ismael "+ day + " /" +month + "/ " + year;
    MailApp.sendEmail(emailAddress, subject, message,{htmlBody:message });
}