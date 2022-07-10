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
