function myFunction() {
  var app = SpreadsheetApp;
  var activeSheet = app.getActiveSpreadsheet().getActiveSheet();
  var currWeek = activeSheet.getRange(3,26).getValue();
  var prevWeek = activeSheet.getRange(3,27).getValue();
  var Day = activeSheet.getRange(3,28).getValue();
  var wGains=0;
  var totalGains = 0;
  var weekStartCap = activeSheet.getRange(14,23).getValue();
  var accountGrowth = 0;
  var lastWeekGains = 0;
  
  for(var i = 3; i<13; i++){
    if(activeSheet.getRange(i,1).getValue() == ""){ //cell empty 
    
      
      
    }
    else{//cell not empty
      
      
      var dateAcquired = activeSheet.getRange(i,1).getValue(); //49
      var shares = activeSheet.getRange(i,3).getValue();  //100
      var curValue = activeSheet.getRange(i,6).getValue(); //6.60
      var initCost = activeSheet.getRange(i,5).getValue(); //600

      if(dateAcquired == currWeek){ //stock was purchased this week
   
        wGains = (shares * curValue)- initCost;
        totalGains += wGains;
        activeSheet.getRange(i,11).setValue(wGains);
  
  
      }
      else{ //stock wasn't purchased in the current week
         if(Day == "R"){
          lastWeekGains = (activeSheet.getRange(i,11).getValue());
          activeSheet.getRange(i,31).setValue((activeSheet.getRange(i,31).getValue())+lastWeekGains);
          activeSheet.getRange(3,28).setValue("");
         }
        lastWeekGains = (activeSheet.getRange(i,31).getValue());
        wGains = ((shares * curValue) - initCost) - lastWeekGains;
        totalGains += wGains;
        activeSheet.getRange(i,11).setValue(wGains);
  
      }
    
    }
    
  }
  activeSheet.getRange(12,11).setValue(totalGains); //total weekly gains
  accountGrowth = (((weekStartCap + totalGains)/ weekStartCap) - 1);
  activeSheet.getRange(12,12).setValue(accountGrowth);
  
}
