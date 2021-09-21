function add_to_cal_all() { 
    var spreadsheet = SpreadsheetApp.getActiveSheet();  
    var eventCal = CalendarApp.getCalendarById(get_your_Email()); 
    var lr = spreadsheet.getLastRow(); 
    var count = spreadsheet.getRange('A2:BJ'+lr).getValues(); 
    
    for (x=0; x<count.length; x++) { 
      var shift = count[x]; 
      var isYes = shift[61]; 

        if(isYes === 'N'){ 
          var summary_pub = shift[4]+'(하판일)'; // 소프트웨어 아키텍처(하판일)
          var startTime_pub = new Date(shift[12]); 
          var endTime_pub = new Date(startTime_pub.getFullYear(), startTime_pub.getMonth(), startTime_pub.getDate() + 1);
          spreadsheet.getRange('BM'+(x+2)).setValue(startTime_pub);
         
          var summary_1_3 = shift[4]+'(1/3)'; // 소프트웨어 아키텍처(1/3)
          var startTime_1_3 = new Date(shift[14]); 
          var endTime1_3 = new Date(startTime_1_3.getFullYear(), startTime_1_3.getMonth(), startTime_1_3.getDate() + 1);
          spreadsheet.getRange('BN'+(x+2)).setValue(startTime_1_3);
          
          var summary_3_3 = shift[4]+'(3/3)'; // 소프트웨어 아키텍처(3/3)
          var startTime_3_3 = new Date(shift[20]); 
          var endTime3_3 = new Date(startTime_3_3.getFullYear(), startTime_3_3.getMonth(), startTime_3_3.getDate() + 1);
          spreadsheet.getRange('BO'+(x+2)).setValue(startTime_3_3);
         
          var summary_cover = shift[4]+'(표지의뢰)'; // 소프트웨어 아키텍처(표지의뢰)
          var startTime_cover = new Date(shift[41]); 
          var endTime_cover = new Date(startTime_cover.getFullYear(), startTime_cover.getMonth(), startTime_cover.getDate() + 1);
          spreadsheet.getRange('BP'+(x+2)).setValue(startTime_cover);

          spreadsheet.getRange('BL'+(x+2)).setValue(new Date);


          // eventCal.createEvent(summary, new Date(startTime), new Date(endTime), event); 
          eventCal.createEvent(summary_pub, startTime_pub, endTime_pub); 
          eventCal.createEvent(summary_1_3, startTime_1_3, endTime1_3); 
          eventCal.createEvent(summary_3_3, startTime_3_3, endTime3_3); 
          eventCal.createEvent(summary_cover, startTime_cover, endTime_cover); 
          // eventCal.createAllDayEvent(summary, new Date(startTime));
        }

    } 
    spreadsheet.getRange('BJ2:BJ'+String(lr)).setValue('Y'); 
} 