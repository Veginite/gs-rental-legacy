//SpreadsheetApp.getUi().alert("Chair: " + i + "\nPass: " + j + "\nnextAvailOutDate: " + nextAvailOutDate + "\n" + "lastReturnDate: " + lastReturnDate + "\n" + "Customer Out: " + Out + "\n" + "Customer In: " + In + "\n"
//+ "Cell 1: " + row.getValues()[0][0] + "\n" + "Cell 2: " + row.getValues()[0][1] + "\n");

//ui.alert("Price: " + rentalPrice + "\n" + "Rental days: " + rentalTimeInDays);

//var nextAvailOutDate = Utilities.formatDate(new Date, "GMT", "yyyy-MM-dd");

//SpreadsheetApp.getUi().alert("chair number " + (i + 1) + " slot " + (j + 1) + "\n" + nextAvailOutDate + ", " + lastReturnDate );

//SpreadsheetApp.getUi().alert("Pass: " + j + "\nnextAvailOutDate: " + nextAvailOutDate + "\n" + "lastReturnDate: " + lastReturnDate + "\n" + "Customer Out: " + Out + "\n" + "Customer In: " + In + "\n"
//+ "Cell 1: " + row.getValues()[0][0] + "\n" + "Cell 2: " + row.getValues()[0][1] + "\n");

//SpreadsheetApp.getUi().alert("s: " + s);

//SpreadsheetApp.getUi().alert(nextAvailOutDate + " to " + lastReturnDate + " is not available \n" + ("F" + (3 + (9*i) + j ) + ":G" + (3 + (9*i) + j )));

//SpreadsheetApp.getUi().alert("Chair: " + i + "\nSlot index: " + chairSlotIndex);

//SpreadsheetApp.getUi().alert("search function finished successfully");

function searchColumnForString(searchString, column)
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Bilbarnstol");
  sheet.getRange(row, column, numRows)
  var columnValues = sheet.getRange(2, column, sheet.getLastRow()).getValues(); //1st is header row
  var searchResult = columnValues.findIndex(searchString); //Row Index - 2

  if(searchResult != -1)
  {
    //searchResult + 2 is row index
    SpreadsheetApp.getActiveSpreadsheet().setActiveRange(sheet.getRange(searchResult + 2, 1))
  }
}

Array.prototype.findIndex = function(searchString){
  if(searchString == "")
  {
    return false;
  }
  
  for (var i = 0; i < this.length; i++)
  {
    if (this[i] == searchString)
    {
      return i;
    }
  }

  return -1;
}