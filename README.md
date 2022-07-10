# ParseTransactions
Java script project for parsing investment transactions from a brokerage and inserting them into a google sheet.

# View a full tutorial on how to set this up [here]()

# Step By Step Instructions:

1.	You must have 3 sheets in your dividend tracking spreadsheet that look like the sheets in this [spreadsheet](https://docs.google.com/spreadsheets/d/1uinoDVIIjm4r4hpkKdcxsKj4hKYiOWwPz_j3FFAM0MI/edit?usp=sharing): https://docs.google.com/spreadsheets/d/1uinoDVIIjm4r4hpkKdcxsKj4hKYiOWwPz_j3FFAM0MI/edit?usp=sharing
2.	In your spreadsheet go to Extensions  Apps Script.
3.	If you have no scripts in your spreadsheet yet you should just have one function called Code.gs.  Delete the contents of this file and replace it with the contents of Code.gs in this repo.
4.	If you have a Code.gs file, copy the contents of Code.gs into your Code.gs file.
5.	Create 4 new files: Constants.gs, Transaction.gs, FidelityTransactionParse.gs, and VanguardTransactionParse.gs using the plus on the left hand side.  Copy the contents of the respective files in this repo to the files you just created (if you only use one of those two brokerages you don’t need the file or functions for the other brokerage).
6.	In the Constants.gs file, add your spreadsheet id to the SPREADSHEET_ID variable.
7.	If you don't have a polygon account, watch [this video](https://www.youtube.com/watch?v=p1ZSY8YkGeM&t=1s) (go to timestamp 05:10) to set one up.  Create a KEYS.gs file and copy the polygon key into that file as shown in that video.
8.	Save all the files. 
9.	Close the script editor and refresh your spreadsheet.

# To Import Your Transactions:
1.	Download your transaction history from your brokerage account (currently only Fidelity and Vanguard are supported at this time).
2.	In your spreadsheet go to File  Import and select the Upload tab then click the blue button and select the transaction history csv file from your local computer.
3.	Once it uploads makes sure to change the “Import location” drop down to “Insert new sheet(s)” then click “Import data.”
4.	Once the data is imported, make sure that you are viewing the transaction history sheet then go to the Scripts menu along the top (if you don’t see this, refresh the page) and choose whichever brokerage you have.
5.	Wait a few minutes for the scripts to run.
6.	If you have a lot of transactions (especially a lot of different dividend payments) you may have to run the script multiple times.  The script can only run for a maximum of six minutes at once.

### If you are stuck on any of the steps or have any questions/issues feel free to check out my [YouTube video]() where I walk through each step or you can contact me directly on [Instagram](https://www.instagram.com/frankmularcik/), [Twitter](https://twitter.com/FrankMularcik) or through email (frank.mularcik.investing@gmail.com).

### If you want to support me make sure to follow me on the above social medias and subscribe to my [YouTube](https://www.youtube.com/c/FrankMularcik) channel.
