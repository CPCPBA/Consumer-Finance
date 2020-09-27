# Consumer-Finance
This is a basic consumer finance app.  It was inspired by the apparent lack of quicken or other financial management software for the Mac during the PC vs Mac wars in the 2010-2014 time frame.  
There was a Quicken for Mac available in 2009 release but it did the transaction download so bad, I learned the internals and realized most of the functionality I wanted was available
in Excel. QBank for Mac later came out as a replacement for Quicken for Mac, they ran into some naming issues, changed their name, licensing got messed up, and I chose not to chase them. 
I've done so many macros, its time to make a productionable version

So here is the data flow
-------------------------
1. Download Transactions from banks daily
2. Load in a data storage medium.  We will start with the existing "All Detailed Expenses" spreadsheet for now
3. Identify categories
4. Create combined budget/actuals - income/expenses report. For now, the report is still the master for available categories and budgets
5. Create immediate budget / actuals report for a single category for mobile phones

Initial environment
-------------------
1. Manually download transactions from banks daily
2. Manually import data into "All detailed expenses" spreadsheet
3. Manually categorize transactions
4. Budget/Actuals income/expenses report already exists as a spreadsheet, budgets and avaialble categories are defined static on the spreadsheet.
5. Immediae budget / actuals for mobile not started

Thoughts
--------
1. Use plaid to get transactions from banks.  Cost is $0.30 per institution/month.  Unlimited transactions.  If you have multiple accounts at the same instiution 
there is no additional cost for multiple accounts at the same institution. Thus my cost is $0.90 per month.  Web crawling does not work well since banks and credit 
card web pages are changing constantly for marketing
2. Use QFX for import.  Other formats are not strict enough between vendors.  QFX is available everywhere.  Where both QFX and OFX is avialble, QFX is only a very small variant 
of OFX.  Regarding Bank ID, QFX has a apparently required FID field that is universal for any financial institute. OFX allows a variety of bankIDs but not a required a specific 
bankID
3. Use Azure SQL database for storage.  I need to learn some type of web programming, Microsoft is world wide and other Microsoft features are everywhere. Microsoft is #2 for web programming primarily 
for the government backing
4. Use Power Automate for the mobile app because it is quick and easy
5.  I found the source code for gnuCash, a FSF version of Quicken.  It may be beneficial to review it to see if there is something I can add.
