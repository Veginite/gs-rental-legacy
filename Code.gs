//doOutdatedChairCheck() is executed by an installed trigger on opening the document. UI functions must be ran by a user for security reasons.
//Running it in onOpen fails because it's not ran by a user.

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Lekia Meny')
  .addItem('Uthyrning', 'createSearchForm')
  .addItem('Återlämna/Avboka', 'createTerminateCustomerForm')
  .addItem('Hyreskontrakt', 'createRentalAgreementForm')
  //.addItem('Redigera Stol', 'createEditChairForm')
  .addItem('Utgångna Stolar', 'doOutdatedChairCheck')
  //.addItem('Förläng', 'createExtendRentalForm')
  //.addItem('Ny stol', 'createAddChairForm')
  //.addItem('Radera stol', 'createDeleteChairForm')
  //.addItem('testFunction', 'testFunction')
  .addToUi();
}

function testFunction(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Bilbarnstol");
  var chairNumber = getChairNumber();
  var lastChairIndex = 3 + ((chairNumber - 1) * 9);
  
  var row = sheet.getRange(lastChairIndex, 2);
  
  var chairNameSegments = row.getCell(1, 1).getValue().toString().split(" ");
  
  var lastChairNumber = parseInt(chairNameSegments[chairNameSegments.length - 1]);
  
  if(lastChairNumber == chairNumber) //Next chair number will be the last number + 1
  {
    

  }
  else //Number of chairs is lower than last chair number, at least one number in the series is missing. We will simply pick the first available number.
  {
    
  }
  
  SpreadsheetApp.getUi().alert("lastChairNumber: " + lastChairNumber + "\n\nchairNumber: " + chairNumber);
}

function deleteChair(formValue){
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Bilbarnstol").deleteRows(2 + (9 * formValue.itemListChair), 9);
}

function createSearchForm(){
  var html = HtmlService.createTemplateFromFile('form_chair_search').evaluate();
  SpreadsheetApp.getUi().showModalDialog(html, "Uthyrning");
}

function createTerminateCustomerForm(){
  var html = HtmlService.createTemplateFromFile('form_customer_terminate').evaluate();
  SpreadsheetApp.getUi().showModalDialog(html, "Avsluta uthyrning/bokning");
}

function createRentalAgreementForm(){
  var html = HtmlService.createTemplateFromFile('form_chair_agreement').evaluate().setHeight(325);
  SpreadsheetApp.getUi().showModalDialog(html, "Hyreskontrakt");
}

function createExtendRentalForm(){
  var html = HtmlService.createTemplateFromFile('form_chair_extend_rent').evaluate().setHeight(325);
  SpreadsheetApp.getUi().showModalDialog(html, "Förläng");
}

function createAddChairForm(){
  var html = HtmlService.createHtmlOutputFromFile('form_chair_add');
  SpreadsheetApp.getUi().showModalDialog(html, "Ny stol");
}

function createDeleteChairForm(){
  var html = HtmlService.createTemplateFromFile('form_chair_delete').evaluate();
  SpreadsheetApp.getUi().showModalDialog(html, "Radera stol");
}
function createEditChairForm(){
  var html = HtmlService.createTemplateFromFile('form_chair_edit').evaluate();
  SpreadsheetApp.getUi().showModalDialog(html, "Redigera stol");
}

function searchChair(formValues) { //Function call from "form_chair_search", scan entire document for available chairs meeting the criteria
  var childAge = formValues.childAge;
  var dateOut = new Date(formValues.dateOut);
  dateOut.setHours(0, 0, 0, 0);
  var dateIn = new Date(formValues.dateIn);
  dateIn.setHours(0, 0, 0, 0);
  var rentalObjectType = formValues.itemListRentalObjectType;
  var rentalObjectIndex = formValues.itemListRentalObject;
  
  var rentalData = getChairAndCustomerIndex(dateOut, dateIn, rentalObjectType, rentalObjectIndex);
  
  var carrierListNotFull = -1;
  
  if(rentalObjectType == 1)
  {
    carrierListNotFull = getCarrierListStatus(rentalObjectIndex);
  }
  
  if( (rentalData.chairIndex > -1 || rentalObjectType == 1 ) && carrierListNotFull != 0)
  {
    var rentalTimeInDays = getDays(dateOut, dateIn);
    var rentalPrice = getRentalPrice(dateOut.getDay(), dateIn.getDay(), rentalTimeInDays, rentalObjectType, rentalObjectIndex);
    
    registerRent(dateOut, dateIn, childAge, rentalObjectType, rentalData.chairIndex, rentalData.customerIndex, rentalPrice);
  }
  else
  {
    var reason = "Ingen ledigt hyresobjekt funnet!\nAnledning: ";
    
    if(rentalObjectType == 0) //Chair
    {
      reason += "-";
    }
    else if(carrierListNotFull == 0) //Carrier
    {
      reason += "Listan för vald uthyrningstyp av babyskydd är full.";
    }
    
    SpreadsheetApp.getUi().alert(reason);
  }
}

function getCarrierListStatus(carrierMargin){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Babyskydd");
  var range = sheet.getRange(49, 6 + (10*carrierMargin), 1, 1);
  var listNotFull = 0;
  
  if(range.getCell(1, 1).isBlank() || range.getCell(1, 1).getValue() == "")
  {
    listNotFull = 1;
  }
  
  return listNotFull;
}

function getChairAndCustomerIndex(dateOut, dateIn, rentalObjectType, rentalObjectIndex)
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getSheetName(rentalObjectType));
  var totalChairs = getChairNumber();
  
  var chairIndex = -1; //When this changes to an index above or equal to 0 a chair has been found
  var customerIndex = -1; //Used for rental agreement
  
  if(rentalObjectType == 1) //Carrier
  {
    chairIndex = rentalObjectIndex;
  }
  
  var rentalObjectInText = "";
  if(rentalObjectIndex == 0)
  {
    rentalObjectInText = "Bilbarnstol";
  }
  else
  {
    rentalObjectInText = "Bältesstol";
  }
  
  //Scan through all chairs in increments of 8 rows with an offset of 3 rows and spacing of 1 row that meet the chair type criteria. This will not run if a carrier is to be rented.
  for(var i = 0; i < totalChairs && chairIndex == -1 && rentalObjectType == 0; i++)
  {
    var row = sheet.getRange("F" + (3 + (9*i) ) + ":G" + (3 + (9*i) ) );
    var nextAvailOutDate = new Date(); //Always begin with checking from today to first customer out date
    nextAvailOutDate.setHours(0, 0, 0, 0);
    /*var lastReturnDate = row.getValues()[0][0];*/
    var lastReturnDate = new Date(row.getValues()[0][0].toString());
    lastReturnDate.setHours(0, 0, 0, 0);
    var customerDateOutOfRange = false;
    var chairType = sheet.getRange("C" + (3 + (9*i) + 6 )).getCell(1, 1).getValue();
    
    for(var j = 0; j < 8 && chairIndex == -1 && !customerDateOutOfRange && chairType == rentalObjectInText && sheet.getRange("F" + (3 + (9*i) + 7 ) ).isBlank(); j++) //Maximum number of slots for a given chair is 8. Last condition is to make sure there is a slot available
    {
      row = sheet.getRange("F" + (3 + (9*i) + j ) + ":G" + (3 + (9*i) + j ));
      
      if(!row.getCell(1, 1).isBlank() && j < 7) //Don't check last chair, if it's not empty it cannot be rented
      {
        lastReturnDate = new Date(row.getValues()[0][0].toString());
        lastReturnDate.setHours(0, 0, 0, 0);
        /*lastReturnDate = row.getValues()[0][0];*/
        lastReturnDate.setDate(lastReturnDate.getDate() - 1);
        
        
        if(dateOut >= nextAvailOutDate && dateIn <= lastReturnDate) //Check valid date range
        {
          
          chairIndex = i;
          customerIndex = j;
        }
        nextAvailOutDate = new Date(row.getValues()[0][1].toString());
        nextAvailOutDate.setHours(0, 0, 0, 0);
        /*nextAvailOutDate = row.getValues()[0][1];*/
        nextAvailOutDate.setDate(nextAvailOutDate.getDate() + 1);
        
      }
      
      else if(row.getCell(1, 1).isBlank())
      { 
        if(dateOut < nextAvailOutDate) //If the cell is blank and the customer's desired date of starting a rent is before or the same day as the previous renter's return date it's logically impossible to rent the chair
        {
          customerDateOutOfRange = true;
        }
        else
        {
          chairIndex = i;
          customerIndex = j;
        }
      }
    }
  }
  
  var chairData = {chairIndex:chairIndex, customerIndex:customerIndex};
  
  return chairData;
}

function registerRent(dateOut, dateIn, childAge, rentalObjectType, rentalObjectIndex, customerIndex, rentalPrice){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getSheetName(rentalObjectType));
  var ui = SpreadsheetApp.getUi();
  var customerName = ui.prompt("Registrering", "Kundnamn", ui.ButtonSet.OK_CANCEL).getResponseText();
  var customerTel = ui.prompt("Registrering", "Telefonnummer", ui.ButtonSet.OK_CANCEL).getResponseText();
  var range;
  
  //SpreadsheetApp.getUi().alert("rentalObjectType: " + rentalObjectType + "\n\nRentalObjectIndex: " + rentalObjectIndex);
  
  if(rentalObjectType == 1) //Carrier
  {
    range = sheet.getRange(sheet.getMaxRows() - 1, 4 + (rentalObjectIndex * 10), 1, 6);
    range.setValues([[customerName,childAge,dateOut,dateIn,customerTel,rentalPrice]]);
    range = sheet.getRange(sheet.getMaxRows() - 1, 4 + (rentalObjectIndex * 10) + 6, 1, 1);
    range.check(); //The index will be lost on sorting so this will be used to locate the customer
  }
  else
  {
    range = sheet.getRange("D" + (3 + (9*rentalObjectIndex) + 7 ) + ":I" + (3 + (9*rentalObjectIndex) + 7 ) );
    range.setValues([[customerName,childAge,dateOut,dateIn,customerTel,rentalPrice]]);
  }
  
  if(rentalObjectType == 1) //Carrier
  {
    range = sheet.getRange(3, 4 + (rentalObjectIndex * 10), sheet.getMaxRows() - 3, 7);
    range.sort(6 + (rentalObjectIndex * 10)); //Sort using column "Utlämnas"
    
  }
  else
  {
    range = sheet.getRange("D" + (3 + (9*rentalObjectIndex) ) + ":I" + (3 + (9*rentalObjectIndex) + 7 ) );
    range.sort(6); //Sort using column "Utlämnas"
  }
  
  //Insert the new rental data at the end of the list and sort it
  
  //-------------------------------------------
  
  var response = ui.prompt("Hyreskontrakt", "Vill du skapa hyreskontrakt? Utlämnas av:", ui.ButtonSet.YES_NO);
  
  if(response.getSelectedButton() == ui.Button.YES)
  { 
    if(rentalObjectType == 1) //Carrier
    {
      //Fetches the row index of a checked cell. The index is lost on sorting above.
      customerIndex = getCarrierIndex(rentalObjectIndex);
    }
    
    //Re-create a named array looking like a form response, for readability
    var formValues = {itemListRentalObjectType:rentalObjectType, itemListRentalObject:rentalObjectIndex, itemListCustomer:customerIndex, employee:response.getResponseText()};
    
    rentalAgreement(formValues);
  }
}

function addChair(formValues){
  var templateSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MALL STOL");
  var chairSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Bilbarnstol");
  var rowIndex = chairSheet.getMaxRows();
  chairSheet.insertRowsAfter(rowIndex, 9);
  templateSheet.getRange("A:J").copyTo(chairSheet.getRange("A" + (rowIndex + 1) ) );
}

function terminateCustomerRental(formValues) {

  var globalRentalObjectChair = +getGlobal("rentalObjectChair");

  var rentalObjectType = parseInt(formValues.itemListRentalObjectType);
  var rentalObjectIndex = parseInt(formValues.itemListRentalObject);
  var customerIndex = parseInt(formValues.itemListCustomer);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getSheetName(rentalObjectType));
  var range;
  
  if(rentalObjectType == globalRentalObjectChair) //Chair
  {
    range = sheet.getRange("D" + (3 + (rentalObjectIndex * 9) + customerIndex) + ":I" + (3 + (rentalObjectIndex * 9) + customerIndex) );
  }
  else //Carrier
  {
    range = sheet.getRange(3 + customerIndex, 4 + (rentalObjectIndex * 10), 1, 6);
  }
  
  var rentalData = range.getValues();
  
  range.clearContent();
  
  if(rentalObjectType == globalRentalObjectChair) //Chair
  {
    range = sheet.getRange("D" + (3 + (rentalObjectIndex * 9) ) + ":I" + (3 + (rentalObjectIndex * 9) + 7) );
    range.sort(6);
  }
  else //Carrier
  {
    range = sheet.getRange(3, 4 + (rentalObjectIndex * 10), 47, 6);
    range.sort(6 + (rentalObjectIndex * 10) ); //Column varies because of design
  }
  
  var ui = SpreadsheetApp.getUi();
  ui.alert("Återlämning/avbokning genomförd", "Kund: " + rentalData[0][0] + "\nUtlämnas: " + Utilities.formatDate(new Date(rentalData[0][2].toString()), "Europe/Stockholm", "yyyy-MM-dd") + "\nÅterlämnas: " + Utilities.formatDate(new Date(rentalData[0][3].toString()), "Europe/Stockholm", "yyyy-MM-dd") + "\nSumma som ska returneras: " + getCustomerReturnAmount(rentalData[0][5]) + ":-", ui.ButtonSet.OK);
}

function rentalAgreement(formValues){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("KONTRAKTPARAMETRAR");
  
  var rentalObjectType = parseInt(formValues.itemListRentalObjectType);
  var rentalObjectIndex = parseInt(formValues.itemListRentalObject);
  
  var ui = SpreadsheetApp.getUi();
  var customerId = ui.prompt("Registrering", "Legitimation", ui.ButtonSet.OK_CANCEL).getResponseText();
  
  var html = HtmlService.createHtmlOutput("<html><body><script>window.open('https://docs.google.com/document/d/" + getRentalDocKey(rentalObjectType, rentalObjectIndex) + "/edit', '_blank');google.script.host.close();</script></body></html>");
  SpreadsheetApp.getUi().showModalDialog(html, "Öppnar hyreskontrakt...");
  
  //5 rows between parameter ranges for chairs and carriers
  sheet.getRange(10, 2 + (rentalObjectType * 5), 1, 1).setValue(rentalObjectIndex);
  sheet.getRange(11, 2 + (rentalObjectType * 5), 1, 1).setValue(formValues.itemListCustomer);
  sheet.getRange(4, 4 + (rentalObjectType * 5), 1, 1).setValue(formValues.employee);
  sheet.getRange(5, 2 + (rentalObjectType * 5), 1, 1).setValue(customerId);
  
  if(rentalObjectType == 1) //Carrier, uncheck the "to rent" checkbox
  {
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Babyskydd");
    sheet.getRange(3 + formValues.itemListCustomer, 4 + (rentalObjectIndex * 10) + 6, 1, 1).uncheck();
  }
}

//TO DO, INCOMPLETE
function extendRent(formValues){
  var rentalObjectType = +formValues.itemListRentalObjectType;
  var rentalObjectIndex = +formValues.itemListRentalObject;
  var customerIndex = +formValues.itemListCustomer;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getSheetName(rentalObjectType));
  var range;
  
  //CONDITIONS: 
  //If the row below is empty extending is guaranteed
  //If not, do an extensive check if all customers after can be moved to other chairs
  //Customer number can be determined with the one 
  
  if(rentalObjectType == 0) //Chair
  {
    range = sheet.getRange("F" + (3 + (rentalObjectIndex * 9) + customerIndex) + ":G" + (3 + (rentalObjectIndex * 9) + customerIndex) );
  }
  else //Carrier
  {
    range = sheet.getRange(3 + customerIndex, 6 + (rentalObjectIndex * 9), 1, 2);
  }
  
  SpreadsheetApp.getUi().alert("Out: " + range.getCell(1, 1).getValue() + " In: " + range.getCell(1, 2).getValue() + "\n\n New date: " + formValues.dateIn + "\n\nrentalObjectType: " + rentalObjectType + "\n\nrentalObjectIndex: " + rentalObjectIndex + "\n\ncustomerIndex: " + customerIndex + "\n\nRange: " + range.getA1Notation());
  
}

function getRentalPrice(dateDayOne, dateDayTwo, days, rentalObjectType, rentalObjectIndex){
  var price = 0;
  
  //SpreadsheetApp.getUi().alert("rentalObjectType: " + rentalObjectType + "rentalObjectIndex: " + rentalObjectIndex);
  
  if(days > 28) //Max rental time is 4 weeks
  {
    price = 800;
  }
  //Day out must be fri-sun and day in must be sat-mon. Total rental time is either 3 or 4 days
  //otherwise it's the price of a day/week.
  else if(days <= 4 && days >= 3 && ( (dateDayOne >= 5 || dateDayOne == 0) && (dateDayTwo >= 6 || dateDayTwo == 0 || dateDayTwo == 1) ) ) //Weekend rent
  {
    if(rentalObjectType == 0 && rentalObjectIndex == 1) //Bigger chair
    {
      price = 150;
    }
    else //Carrier short time & regular chair
    {
      price = 250;
    }
  }
  else if(days == 1 || days == 2) //Bigger chair
  {
    if(rentalObjectType == 0 && rentalObjectIndex == 1)
    {
      price = 100;
    }
    else //Carrier short time & regular chair
    {
      price = 150;
    }
  }
  else
  {
    //SpreadsheetApp.getUi().alert("days: " + days); 
    price = (Math.ceil(days / 7) * 100); //Base price is 100 * weeks (rounded up) plus a constant
    if(rentalObjectType == 0 && rentalObjectIndex == 1) //Bigger chair
    {
      price += 100;
    }
    else //Carrier short time & regular chair
    {
      price += 200;
    }
  }
  
  return price;
}

function getCustomerReturnAmount(rentalPrice){
  
  var returnAmount = 0;
  if(rentalPrice == 800) //Long time carrier
  {
    returnAmount = 100;
  }
  else if(rentalPrice >= 500) //Extended rental scenario
  {
    returnAmount = 0;
  }
  else //Regular scenario
  {
    returnAmount = 500 - rentalPrice;
  }
  return returnAmount;
}

function getDays(dateOne, dateTwo){
  var oneDay = 24 * 60 * 60 * 1000;
  var diffDays = Math.round(Math.abs((dateOne.getTime() - dateTwo.getTime()) / (oneDay)));
  return diffDays + 1;
}

function createListFromSheetColumn(sheetName, columnNumber, amount, margin, spacing){
  
  if(isNaN(amount)) //amount param is either a number or a function name. If function name, fetch return value from server and use for iteration
                    //"this" refers to the spreadsheetapp object, and the index is a named array of function names
  {
    amount = this[amount](); //Call the function
  }
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var itemList = [];
  
  for(var i = 0; i < amount; i++){
    var cell = sheet.getRange( (margin + (spacing*i) ), columnNumber, amount);

    if(!cell.isBlank() && cell.getValue() != "")
    {
      itemList.push(cell.getValue());
    }
  }
  
  return itemList;
}

function getChairNumber(){
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Bilbarnstol");
  var totalChairs = (sheet.getMaxRows() - 2) / 9; //-2 is the frozen top row and the excess row below it. Each chair has a column height count of 8 plus one for spacing for a total of 9
  return totalChairs;
}

function getSheetName(id){
  var sheetNames = ["Bilbarnstol", "Babyskydd", "Bilbarnstol"];
  return sheetNames[id];
}

function getCarrierIndex(rentalObjectIndex)
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Babyskydd");
  
  var carrierListLength = sheet.getMaxRows() - 3;
  
  var range = sheet.getRange(3, 4 + (rentalObjectIndex * 10) + 6, carrierListLength);
  
  var index = -1;
  for(var i = 0; i < carrierListLength && index < 0; i++)
  {
    if(range.getCell(i+1, 1).isChecked())
    {
      index = i;
    }
  }
  
  return index;
}

function getRentalDocKey(rentalObjectType, rentalObjectIndex){
  var key;
  if(rentalObjectType == 0) //Chair
  {
    key = "1cQB4kjsh2NsLhMvNHH4w3cN_F0tpHPdnh-B-PYDUZ50";
  }
  else //Carrier
  {
    if(rentalObjectIndex == 0) //Short time
    {
      key = "161aA0-nG_UN-3m-2rsYWfNHpktV_9OtxakuN_gCSwcM";
    }
    else //Long time
    {
      key = "1el22hrTcph9jSOawZLYvJ67SDtKfc2tvll7rYZqiTK8";
    }
  }
  return key;
}

function getCustomerRentalPeriod(rentalObjectType, rentalObjectIndex, customerIndex){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getSheetName(rentalObjectType));
  var range;
  if(rentalObjectType == 0) //Chair
  {
    range = sheet.getRange("F" + (3 + (rentalObjectIndex * 9) + customerIndex) + ":G" + (3 + (rentalObjectIndex * 9) + customerIndex) );
  }
  else //Carrier
  {
    range = sheet.getRange(3 + customerIndex, 6 + (rentalObjectIndex * 10), 1, 2);
  }
  
  return range.getDisplayValues()[0]; //HTMLService WILL NOT accept a spreadsheet date object.
}

function doOutdatedChairCheck(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Bilbarnstol");
  var outdatedList = "";
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  
  for(var i = 0; i < getChairNumber(); i++)
  {
    var customerReturnDate = sheet.getRange(3+(9*i), 7); //Return date
    if(!customerReturnDate.isBlank() && customerReturnDate.getValue() < today)
    {
      outdatedList += sheet.getRange(3+(9*i), 2).getValue(); //Chair name
      outdatedList += "<br>";
    }
  }
  
  //sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Babyskydd");
  /*babyskydd här*/
  
  if(outdatedList.length > 0)
  {
    var html = HtmlService.createHtmlOutput(outdatedList).setTitle("Utgångna Stolar och Babyskydd");
    SpreadsheetApp.getUi().showSidebar(html);
    SpreadsheetApp.getUi().alert("En eller fler stolar/skydd är utgångna. Använd Lekia Menyn för att se listan.");
  }
}

function getGlobal(key)
{
  return PropertiesService.getScriptProperties().getProperty(key);
}

function include(File) {
  return HtmlService.createHtmlOutputFromFile(File).getContent();
}