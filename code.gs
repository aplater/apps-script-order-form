var form = FormApp.openById(<formID>);
var ss = SpreadsheetApp.getActive();
var individualSheets;

// Boy blazers
var bbList = form.getItemById(<itemID>).asListItem();
var bb = ss.getSheetByName("Boy Blazers Stock");
  
// Girl blazers
var gbList = form.getItemById(<itemID>).asListItem();
var gb = ss.getSheetByName("Girl Blazers Stock");
  
// Jumpers
var jList = form.getItemById(<itemID>).asListItem();
var j = ss.getSheetByName("Jumpers Stock");
  
// Cardigans
var cList = form.getItemById(<itemID>).asListItem();
var c = ss.getSheetByName("Cardigans Stock");
    
// Ties
var tList = form.getItemById(<itemID>).asListItem();
var t = ss.getSheetByName("Tie Stock");

// Waiting List
var waitingList = ss.getSheetByName("Waiting List");

function updateForm() {
  
  // Updates the drop downs using the function below
  updateDropDowns(bbList,bb);
  updateDropDowns(gbList,gb);
  updateDropDowns(jList,j);
  updateDropDowns(cList,c);
 
  // Ties are done separately
  var item = t.getRange(2, 1, t.getMaxRows() - 1).getValues();
  var quality = t.getRange(2, 2, t.getMaxRows() - 1).getValues();
  var checkbox = t.getRange(2, 4, t.getMaxRows() - 1).getValues();
  
  var listItem
  var uniformList = [];
  
  // converts ro an array ignoring empty cells
  for(var i = 0; i < item.length; i++) {
  
    if(item[i][0] != "" && checkbox[i][0] == false) {
  
      listItem = item[i][0];
      
      if(quality[i][0] != "") {
        listItem += ", Quality: " + quality[i][0];
      }
      
      if(uniformList.slice(-1)[0] != listItem) { 
         uniformList.push(listItem);
      }
      
    }
    
  }
  
  uniformList = uniformList.sort()
  
  // populate the drop-down with the array data
  tList.setChoiceValues(uniformList);

}

function updateDropDowns(listId,sheet) {

  // gets the item and size values from spreadsheet
  var item = sheet.getRange(2, 1, sheet.getMaxRows() - 1).getValues();
  var quality = sheet.getRange(2, 2, sheet.getMaxRows() - 1).getValues();
  var size = sheet.getRange(2, 3, sheet.getMaxRows() - 1).getValues();
  var checkbox = sheet.getRange(2, 5, sheet.getMaxRows() - 1).getValues();
  
  var listItem
  var uniformList = [];
  
  // converts ro an array ignoring empty cells
  for(var i = 0; i < item.length; i++) {
  
    if(item[i][0] != "" && checkbox[i][0] == false) {
  
      if(size[i][0] != "") {
        listItem = item[i][0] + ", size " + size[i][0];
      }
           
      else {
        listItem = item[i][0];
      }
      
      if(quality[i][0] != "") {
        listItem += ", Quality: " + quality[i][0];
      }
      
      if(uniformList.slice(-1)[0] != listItem) { 
         uniformList.push(listItem);
      }
      
    }
    
  }
  
  uniformList = uniformList.sort()
  
  // populate the drop-down with the array data
  listId.setChoiceValues(uniformList);

}

function onFormSubmit(e) {

  var responseSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  lastRow = responseSheet.getLastRow(); 
  
  this.nameVal = responseSheet.getRange(lastRow, 3).getValue();

  this.emailVal = responseSheet.getRange(lastRow, 2).getValue();
  
  this.bbVal = responseSheet.getRange(lastRow, 4).getValue();
  updateSheet(bb,this.bbVal,this.nameVal);
 
  this.gbVal = responseSheet.getRange(lastRow, 5).getValue();
  updateSheet(gb,this.gbVal,this.nameVal);

  this.jVal = responseSheet.getRange(lastRow, 6).getValue();
  updateSheet(j,this.jVal,this.nameVal);
  
  this.cVal = responseSheet.getRange(lastRow, 7).getValue();
  updateSheet(c,this.cVal,this.nameVal);
  
  this.tVal = responseSheet.getRange(lastRow, 8).getValue();
  
  //tie order, so match to item in tie stock:
  if(this.tVal != ""){  //only if there is a tie order
  
  Logger.log("Getting tie stock details");
    // gets the item and size values from spreadsheet
    var items = t.getRange(2, 1, t.getMaxRows() - 1).getValues();
    var qualities = t.getRange(2, 2, t.getMaxRows() - 1).getValues();
    var checkboxes = t.getRange(2, 4, t.getMaxRows() - 1).getValues();

    var searchItem;
            Logger.log ("T Val = " + this.tVal); 
    // converts to an array ignoring empty cells
    for(var i = 0; i < items.length; i++) {
    
      if(items[i][0] != "" && checkboxes[i][0] == false) {
    
        searchItem = items[i][0];
        
        if(qualities[i][0] != "") {
          searchItem += ", Quality: " + qualities[i][0];
        }
        Logger.log ("Search Item = " + searchItem);
       
        if(this.tVal == searchItem) {  //match found in stock list
          Logger.log("snap!");
          var range = i+2
          t.getRange("D"+range).setValue('true');
          t.getRange("C"+range).setValue(this.nameVal);
          updateForm();
          break;
        }
        
      }
      
    }
    
  }
  
  Logger.log("Waiting list")
  
  // Update waiting list
  
  // Get waiting list item and size
  this.waitingItem = responseSheet.getRange(lastRow, 10).getValue();
  this.waitingSize = responseSheet.getRange(lastRow, 11).getValue();
  
  Logger.log(this.waitingItem);
      
  // Add the item and size to waiting list
   
  if (this.waitingItem != "") {
    Logger.log("Adding to waiting list")
    waitingList.appendRow([this.waitingItem, this.waitingSize, this.nameVal])
  }
    
  this.request = responseSheet.getRange(lastRow, 9).getValue();
  
  // Email New Order
   
  if (this.waitingItem != "") {
  
    GmailApp.sendEmail("example@gmail.com", "Uniform Waiting List", this.nameVal + " has added an item to the waiting list \n\nName:" + this.nameVal + "\nEmail address: "+ this.emailVal + "\n\nAdded to waiting list: " + this.waitingItem + ", size " + this.waitingSize);
  
  }
  
  if (this.bbVal != "" || this.gbVal != "" || this.jVal != "" || this.cVal != "" || this.tVal != "") {
    
    GmailApp.sendEmail("example@gmail.com", "Uniform Order", "A new order has been placed \n\nName:" + this.nameVal + "\nEmail address: "+ this.emailVal + "\n\nItem(s) Ordered:\n\nBoy Blazers: " + this.bbVal + "\nGirl Blazers: " + this.gbVal + "\nJumpers: " + this.jVal + "\nCardigans: " + this.cVal + "\nTies: " + this.tVal + "\n\nRequested trousers/shirts: " + this.request);
  
  }  
  
  // Email Receipt
  
  if (this.waitingItem != "") {
  
    GmailApp.sendEmail(this.emailVal, "Confirmation of Uniform Waiting List", "Your item has been successfully added to the waiting list with the following details: \n\nName:" + this.nameVal + "\nEmail address: "+ this.emailVal + "\n\nAdded to waiting list: " + this.waitingItem + ", size " + this.waitingSize);
  
  }
  
  if (this.bbVal != "" || this.gbVal != "" || this.jVal != "" || this.cVal != "" || this.tVal != "") {
    
    GmailApp.sendEmail(this.emailVal, "Confirmation of Uniform Order", "Your order has been placed. Details: \n\nName:" + this.nameVal + "\nEmail address: "+ this.emailVal + "\n\nItem(s) Ordered:\n\nBoy Blazers: " + this.bbVal + "\nGirl Blazers: " + this.gbVal + "\nJumpers: " + this.jVal + "\nCardigans: " + this.cVal + "\nTies: " + this.tVal + "\n\nRequested trousers/shirts: " + this.request);
  
  }  
  
}

function updateSheet(mySheet,value,name) {

  Logger.log("FUNCTION" + mySheet.getName() + " // " + value + " // " + name)
  
  if(value != ""){
  
    // gets the item and size values from spreadsheet
    var item = mySheet.getRange(2, 1, mySheet.getMaxRows() - 1).getValues();
    var quality = mySheet.getRange(2, 2, mySheet.getMaxRows() - 1).getValues();
    var size = mySheet.getRange(2, 3, mySheet.getMaxRows() - 1).getValues();
    var checkbox = mySheet.getRange(2, 5, mySheet.getMaxRows() - 1).getValues();
    
    var searchItem;
    
    // converts ro an array ignoring empty cells
    for(var i = 0; i < item.length; i++) {
    
      if(item[i][0] != "" && checkbox[i][0] == false) {
    
        if(size[i][0] != "") {
          searchItem = item[i][0] + ", size " + size[i][0];
        }
             
        else {
          searchItem = item[i][0];
        }
        
        if(quality[i][0] != "") {
          searchItem += ", Quality: " + quality[i][0];
        }
        
        if(value == searchItem) {
          Logger.log("match")
          var range = i+2
          mySheet.getRange("E"+range).setValue('true');
          mySheet.getRange("D"+range).setValue(name);
          updateForm();
          break;
        }
        
      }
      
    }
    
  }

}
