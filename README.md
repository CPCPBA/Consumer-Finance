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
5. On-Demand budget / actuals report for iPhone/mobile not started

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

1 October 2020
The initial version of import transactions using VBA code in Excel is working and being used in internal production.  There are issues with the ribbon buttons absolute reference and the code being stored in the same Excel Workbook as the data.  It is easily resolvable however there exists a prefrence to implment this code in Azure SQL.  Therefore no more development time will be spent on this branch

2 November 2020
2 new GIT repositories were created.  The existing Consumer-Finance git was moved one folder lower to Consumer-Finance/import-Transactions-VBA.  A new branch was created called Consumer-Finance/get-Transactions

4 December 2020
Fixed no Duplicate transaction rules - 
    1.  No longer use FITID
    2.  UUID = Source & date & Description & Amount. 
    3.  If original is in database/detailed transactions table already, dup is ignored
    4.  If original is in different transaction file and not in table, dup is ignored
    5.  If original and dup are in the same file, description of dup is appended with " -i{n} where
        {n} is a sequential integer starting at 1
Regular Expressions now used to identify categories. 
    1. Can use regular expression patterns
    2. Exact match no longer required
    3. Keywords no longer have to begin at beginning of description
    4. Whole process is much smaller
Regular Expressions now used to locate transactions entries in QFX file
    1. Location within file can be long integers
    2. No longer need workaround for instr requiring short integers
    3. Whole process is much smaller
Colorize
    1. Added column for category match
    2. now used constant called firstCol and last col
New formulas
    1. Month-Category in detailed transactions table is now a formula, allows for easier manual edits
    2. UUID field is now a formula in detailed transaction table.  Allows for easiuer manual edits