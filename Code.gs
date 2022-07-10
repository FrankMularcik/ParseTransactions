function onOpen() 
{
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Scripts')
  .addItem("Parse Fidelity Transactions", 'ParseFidelityTransactions')
  .addItem("Parse Vanguard Transactions", 'ParseVanguardTransactions')
  .addToUi();
}
