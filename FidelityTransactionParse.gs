function ParseFidelityTransactions()
{
  var endDate = (new Date()).getTime() + MinutesToMs(5);
  var DATE_COL = 1;
  var ACTION_COL = 2;
  var SYMBOL_COL = 3;
  var QUANTITY_COL = 6;
  var PRICE_COL = 7;
  var AMOUNT_COL = 11;
  var row = 1;
  var sheet = SpreadsheetApp.getActiveSheet();
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
  
  var emptyCounter = 0;
  var depositArray = new Array(0);
  var buyArray = new Array(0);
  var sellArray = new Array(0);
  var divArray = new Array(0);

  var nextDate = new Date(sheet.getRange(row, DATE_COL).getValue());
  do
  {
    if ((new Date()).getTime() > endDate)
    {
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VAR_SHEET_NAME).getRange(NEXT_PARSE_ROW, PARSE_VAR_COL).setValue(row);
      UpdateSheetAndShowUI(buyArray, sellArray, divArray, depositArray, true);
      return 0;
    }
    if (nextDate.toString() == "Invalid Date")
    {
      emptyCounter+= 1;
    }
    else
    {
      emptyCounter = 0;
      var rowValues = sheet.getRange(row, 1, 1, 15).getDisplayValues()[0];
      if (rowValues[ACTION_COL - 1].includes("BOUGHT") || rowValues[ACTION_COL - 1].includes("REINVESTMENT"))
      {
        if (rowValues[SYMBOL_COL - 1] != "SPAXX")
        {
          buyArray[buyArray.length] = new Transaction(nextDate, rowValues[SYMBOL_COL - 1], 
          rowValues[QUANTITY_COL - 1], rowValues[PRICE_COL - 1], Math.abs(rowValues[AMOUNT_COL - 1]));
        }
      }
      else if (rowValues[ACTION_COL - 1].includes("SOLD"))
      {
        sellArray[sellArray.length] = new Transaction(nextDate, rowValues[SYMBOL_COL - 1], 
        rowValues[QUANTITY_COL - 1], rowValues[PRICE_COL - 1], Math.abs(rowValues[AMOUNT_COL - 1]));
      }
      else if (rowValues[ACTION_COL - 1].includes("DIVIDEND"))
      {
        var totalDiv = Number(rowValues[AMOUNT_COL - 1]);
        var ticker = rowValues[SYMBOL_COL - 1];
        if (ticker == "SPAXX")
        {
          ticker = "CASH";
          dividend = 1.0;
          shares = totalDiv;
        }
        else
        {
          Wait(12100);
          var dividend = GetDivMatchingDate(ticker, nextDate);
          if (dividend != 0)
          {
            shares = totalDiv / dividend;
          }
        }
        divArray[divArray.length] = new Transaction(nextDate, ticker, shares, 
        dividend, totalDiv);
      }
      else if (rowValues[ACTION_COL - 1].includes("Elect") || rowValues[ACTION_COL - 1].includes("ELAN")
      || rowValues[ACTION_COL - 1].includes("IN LIEU"))
      {
        depositArray[depositArray.length] = new Transaction(nextDate, "CASH", Number(rowValues[AMOUNT_COL - 1]), 1, Number(rowValues[AMOUNT_COL - 1]));
      }
    }
    row++;
    nextDate = new Date(sheet.getRange(row, DATE_COL).getValue());
  }
  while (emptyCounter < 3);

  UpdateSheetAndShowUI(buyArray, sellArray, divArray, depositArray, false);
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VAR_SHEET_NAME).getRange(NEXT_PARSE_ROW, PARSE_VAR_COL).setValue(1);
}

function GetTotalTransactions(arr)
{
  var totalArr = 0;
  for (var i = 0; i < arr.length; i++)
  {
    totalArr += arr[i].amount;
  }
  return Number.parseFloat(totalArr).toFixed(2);
}

function GetDivMatchingDate(ticker, date)
{
  var divData = GetDividend(ticker);
  for(var i = 0; i < divData.length; i++)
  {
    var diffDays = (new Date(divData[i].pay_date) - date) / (1000*60*60*24);
    if (diffDays < 5 && diffDays > -5)
    {
      return divData[i].cash_amount;
    }
  }
  return 0;
}

function AddTransactionsToSheet(sheet, array, varSheetCol, varSheetRow, varSheet, firstCol)
{
  var nextRow = varSheet.getRange(varSheetRow, varSheetCol).getValue();
  for (var i = array.length - 1; i >= 0; i--)
  {
    var entry = array[i];
    var arr = [entry.ticker, entry.date, entry.shares, entry.price, entry.amount];
    sheet.getRange(nextRow, firstCol, 1, 5).setValues([arr]);
    nextRow++;
  }
  varSheet.getRange(varSheetRow, varSheetCol).setValue(nextRow);
}

function UpdateSheetAndShowUI(buyArr, sellArr, divArr, depositArr, notDone)
{
  var totalDeposits = GetTotalTransactions(depositArr);
  var totalBuys = GetTotalTransactions(buyArr);
  var totalSells = GetTotalTransactions(sellArr);
  var totalDivs = GetTotalTransactions(divArr);

  var spreadSheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  var variableSheet = spreadSheet.getSheetByName(VAR_SHEET_NAME);
  var firstBuyCol = variableSheet.getRange(FIRST_BUY_ROW, PARSE_VAR_COL).getValue();
  var firstSellCol = variableSheet.getRange(FIRST_SELL_ROW, PARSE_VAR_COL).getValue();
  var firstDivCol = variableSheet.getRange(FIRST_DIV_ROW, PARSE_VAR_COL).getValue();

  AddTransactionsToSheet(spreadSheet.getSheetByName(TRANSACTION_SHEET), buyArr, PARSE_VAR_COL, NEXT_BUY_ROW, variableSheet, firstBuyCol);
  AddTransactionsToSheet(spreadSheet.getSheetByName(TRANSACTION_SHEET), depositArr, PARSE_VAR_COL, NEXT_BUY_ROW, variableSheet, firstBuyCol);
  AddTransactionsToSheet(spreadSheet.getSheetByName(TRANSACTION_SHEET), sellArr, PARSE_VAR_COL, NEXT_SELL_ROW, variableSheet, firstSellCol);
  AddTransactionsToSheet(spreadSheet.getSheetByName(DIVIDEND_SHEET), divArr, PARSE_VAR_COL, NEXT_DIV_ROW, variableSheet, firstDivCol);

  var toShow = "Total Deposits: $" + totalDeposits;
  toShow += "\nTotal Buys: $" + totalBuys;
  toShow += "\nTotal Sells: $" + totalSells;
  toShow += "\nTotal Dividends: $" + totalDivs;
  if (notDone)
  {
    toShow += "\n\nNot all transactions were parsed, run the script again.";
  }

  SpreadsheetApp.getUi().alert(toShow);
}
