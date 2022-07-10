function ParseVanguardTransactions()
{
  var endDate = (new Date()).getTime() + MinutesToMs(5);
  var DATE_COL = 2;
  var ACTION_COL = 4;
  var SYMBOL_COL = 7;
  var SHARE_COL = 8;
  var PRICE_COL = 9;
  var AMOUNT_COL = 10;
  var sheet = SpreadsheetApp.getActiveSheet();
  var row = 1;
  var nextRow = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VAR_SHEET_NAME).getRange(NEXT_PARSE_ROW, PARSE_VAR_COL).getValue();
  var nextDate = new Date(sheet.getRange(nextRow, DATE_COL).getValue());
  if (nextDate.toString() != "Invalid Date")
  {
    row = nextRow;
  }
  else
  {
    var maxRows = sheet.getMaxRows();
    for (var i = 1; i <= maxRows; i++)
    {
      if (sheet.getRange(i, 1).getValue() != "")
      {
        row = i + 1;
        break;
      }
    }
  }

  var depositArray = new Array(0);
  var buyArray = new Array(0);
  var sellArray = new Array(0);
  var divArray = new Array(0);

  do
  {
    if ((new Date()).getTime() > endDate)
    {
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VAR_SHEET_NAME).getRange(NEXT_PARSE_ROW, PARSE_VAR_COL).setValue(row);
      UpdateSheetAndShowUI(buyArray, sellArray, divArray, depositArray, true);
      return 0;
    }
    var rowValues = sheet.getRange(row, 1, 1, 15).getDisplayValues()[0];
    var action = rowValues[ACTION_COL - 1];
    if (action.includes("Buy") || action.includes("Reinvestment"))
    {
      buyArray[buyArray.length] = new Transaction(nextDate, rowValues[SYMBOL_COL - 1], Math.abs(rowValues[SHARE_COL - 1]), 
      rowValues[PRICE_COL - 1], Math.abs(rowValues[AMOUNT_COL - 1]));

    }
    else if (action.includes("Contribution"))
    {
      depositArray[depositArray.length] = new Transaction(nextDate, "CASH", Math.abs(rowValues[AMOUNT_COL - 1]), 
      1, Math.abs(rowValues[AMOUNT_COL - 1]));
    }
    else if (action.includes("Sell"))
    {
      sellArray[sellArray.length] = new Transaction(nextDate, rowValues[SYMBOL_COL - 1], Math.abs(rowValues[SHARE_COL - 1]), 
      rowValues[PRICE_COL - 1], Math.abs(rowValues[AMOUNT_COL - 1]));
      
    }
    else if (action.includes("Dividend"))
    {
      var totalDiv = Number(rowValues[AMOUNT_COL - 1]);
      var ticker = rowValues[SYMBOL_COL - 1];
      Wait(12100);
      var dividend = GetDivMatchingDate(ticker, nextDate);
      var shares = 0;
      if (dividend != 0)
      {
        shares = totalDiv / dividend;
      }  
      divArray[divArray.length] = new Transaction(nextDate, ticker, shares, 
      dividend, totalDiv);
    }
    row++;
    nextDate = new Date(sheet.getRange(row, DATE_COL).getValue());
  } while (nextDate.toString() != "Invalid Date");

  UpdateSheetAndShowUI(buyArray, sellArray, divArray, depositArray, false);
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VAR_SHEET_NAME).getRange(NEXT_PARSE_ROW, PARSE_VAR_COL).setValue(1);
}
