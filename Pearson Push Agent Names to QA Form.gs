function updateForm() {
  
    var agentNamesTab = SpreadsheetApp.getActive().getSheetByName("Data");
  
    var formAgentList = FormApp.openById("13NX39tsXFLPqQwVR7iJoveFx4iBTdr6SUIITzZYuQOs").getItemById("1797469837").asListItem();
  
    var activeStartRow = 2;
    var activeNameCol = 1;
        
     var namesActive = agentNamesTab.getRange(activeStartRow, activeNameCol, agentNamesTab.getMaxRows()).getValues();
  
    for(var activeCount = 0; namesActive[activeCount] != ""; activeCount++) {}
  
    namesActive = agentNamesTab.getRange(activeStartRow, activeNameCol, activeCount).getValues();
    
    formAgentList.setChoiceValues(namesActive);
    
  }  