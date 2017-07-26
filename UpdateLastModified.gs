
function onOpen() {
  var menuAdd = [{name: 'Add Issue', functionName: 'addIssue_'}, {name: 'Search Issue', functionName: 'searchIssue_'}, {name: 'Update Last Modified', functionName: 'updateLastModifiedColumn'} ];
  SpreadsheetApp.getActive().addMenu('EDM', menuAdd);
}
function updateLastModifiedColumn() { 
  //get spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];    
  
  //columns to use
  var lastModifiedColumn = "F";
  var lastModifyingUserColumn = "G"; 
  var hyperlinkColumn = "H";  
  
  //get rows where selection is made 
  var selectedRange = String(ss.getActiveRange().getA1Notation());
  
  //string manipulation on range  
  var splitRange = selectedRange.split(":");      
      
  //used to check if only one cell is selected
  var selectedRangeAsString = String(selectedRange);  
  
  if(selectedRangeAsString.indexOf(":") < 1) {       
      var row = selectedRange.slice(1);      
      singleSelectUpdate(row);
    }
  
    else { //MULTISELCT 
      var rowVerticalStart = parseInt(String(splitRange[0].slice(1)));
      var rowVerticalEnd = parseInt(String(splitRange[1]).slice(1));
      
      //check to see that selction is made in a single column 
      var columnCheck1 = String(splitRange[0]).substring(0,1);
      var columnCheck2 = String(splitRange[1]).substring(0,1); 
            
      //checks that the column is lastModifiedColumn
      if( columnCheck1 == lastModifiedColumn && columnCheck2 == lastModifiedColumn ) {                        
         for (var i = rowVerticalStart; i < rowVerticalEnd + 1; ++i) {
           singleSelectUpdate(i)            
         }        
      }                  
               
  else {
    var ui = SpreadsheetApp.getUi();
    ui.alert('Selection must be made the \'Modified\' column.');  
  }        
 }
  
// i is the row #
function singleSelectUpdate(i) {  
      //get formula of H row  
      Logger.log('i is: ' + i);
      var formula = sheet.getRange(hyperlinkColumn + i).getFormula();      
      Logger.log('i after is: ' + i);
      Logger.log('link is: ' + formula); 
  
      //check empty boxes, checks that formula starts with a =hyperlink
      if((String(formula).substring(0, 10)) == "=hyperlink") {      
        //get id from formula 
        var id = String(String(formula).split('=')[2]).slice(0,-10);
    
        //fetch file from drive
        var driveFile = DriveApp.getFileById(id);
        var lastModifiedBy = Drive.Files.get(id).lastModifyingUser.emailAddress;
        Logger.log('documentID: ' + id); 
        Logger.log('last modified by: ' + lastModifiedBy); 
    
        //get last modified date    
        var lastUpdated = driveFile.getLastUpdated();
         
        //insert last modified date 
        sheet.getRange(lastModifiedColumn + i).setValue(lastUpdated);   
        sheet.getRange(lastModifyingUserColumn + i).setValue(lastModifiedBy); 
        Logger.log('doc was last updated: ' + lastUpdated); 
        Logger.log('doc name: ' + driveFile.getName()); 
      }
  
      else {
        //turn cell red if no hyperlink is found
        sheet.getRange(hyperlinkColumn + i).setBackground("red"); 
      }  
}              
  // refresh sheet
  SpreadsheetApp.flush();    
}
