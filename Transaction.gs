class Transaction
{
  constructor(date, ticker, shares, price, amount)
  {
    this.date = date;
    this.ticker = ticker;
    this.shares = shares;
    this.price = price;
    this.amount = amount;
  }
}

function MinutesToMs(min)
{
  return min*60*1000;
}

function GetDividend(ticker) 
{
  var response = UrlFetchApp.fetch("https://api.polygon.io/v3/reference/dividends?ticker=" + 
  ticker + "&limit=24&apiKey=" + POLYGON_KEY);
  var data = JSON.parse(response);
  if (data.results.length == 0)
  {
    return 0;
  }
  return data.results;
}

function Wait(ms)
{
  var d1 = new Date();
  do
  {
    d2 = new Date();
  } while (d2 - d1 < ms);
}
