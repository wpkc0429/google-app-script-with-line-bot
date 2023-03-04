var SHEET_ID = {SHEET_ID};
var SHEET_NAME = {SHEET_NAME};
var CHANNEL_ACCESS_TOKEN = {CHANNEL_ACCESS_TOKEN};


function doPost(e) {
  var msg = JSON.parse(e.postData.contents);
  var events = msg.events[0];
  if (events) {
    
    var replyToken =  events.replyToken;
    var userMessage = events.message.text;
    
    var messages = onSearch(userMessage);
    
    
    var payload = {
      replyToken: replyToken,
      messages: messages
    };
    var option = {
      'headers': {
        'Content-Type': 'application/json; charset=UTF-8',
        'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN
      },
      'method': 'post',
      'payload': JSON.stringify(payload)
    };          
    var response = UrlFetchApp.fetch(
      'https://api.line.me/v2/bot/message/reply',
      option
    );
  }

  function onSearch(search_string)
  {
      var spreadSheet = SpreadsheetApp.openById(SHEET_ID);
      var sheet = spreadSheet.getSheetByName(SHEET_NAME);

      var textFinder = sheet.createTextFinder(search_string)
      var searchResults = textFinder.findAll()


      var row
      var replyText
      var messages = [];
      var column =1; //column Index
      for (var counter = 0; counter <= searchResults.length; counter = counter + 1) {
        if(typeof(searchResults[counter])==='undefined') continue;
        row = searchResults[counter].getRow()
        var columnValues = sheet.getRange(row, column, 1, 5).getValues()[0]; //1st is header row
        var position = '無';
        if(columnValues[1]!=''){
          position = columnValues[1];
        }
        var participant_num = 1;
        if(columnValues[2]!=''){
          participant_num = columnValues[2];
        }
        var seating = '無';
        if(columnValues[3]!=''){
          seating = columnValues[3];
        }
        var is_checkin = '是';
        if(columnValues[4]!='1'||columnValues[4]!='是'){
          is_checkin = '否';
        }
        replyText = '姓名:'+columnValues[0]+'\n'+'稱謂:'+position+'\n'+'人數:'+participant_num+'人\n'+'位置:'+seating+'\n'+'已報到:'+is_checkin;
        messages.push(replyText)
      }
      
      return [{
            'type': 'text',
            'text': messages.join('\n\n===================\n\n')
      }];
  }

 
}
