function onOpen(){ 
    var ui = SpreadsheetApp.getUi(); 
    
    ui.createMenu("◈스크립트 자동화◈") 
    .addItem('캘린더_일정추가(all)', 'goCreate') 
    //.addItem('캘린더_일정추가(1/3)', 'add_to_cal_1_3') 
    .addToUi();
}