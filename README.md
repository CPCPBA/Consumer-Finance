# Consumer-Finance
This is a basic consumer finance app.  It was inspired by the apparent lack of quicken or other financial management software for the Mac during the PC vs Mac wars in the 2010-2014 time frame.  
There was a Quicken for Mac available in 2009 release but it did the transaction download so bad, I learned the internals and realized most of the functionality I wanted was available
in Excel. QBank for Mac later came out as a replacement for Quicken for Mac, they ran into some naming issues, changed their name, changed the licensing, and I chose not to follow. 
I've done so many macros, its time to make a productionable version

 Author    : Christopher Prost, CP Business Analysis LLC. (9/21/2020)
 Website   : http://www.cpbusinessanalysis.com
 Copyright : 2020 CP Business Analysis LLC.  All Rights Reserved.


 Usage:
 ------
 Main
     input : QFX files in ~/downloads directory
    output : 'Expense Detail' Sheet in 'Actuals Analysis' Workbook

---------------------------------------------------------------------------------------
# History
# -------
 This script was written in 3 days as a fairly quick means to read transaction files from FIs
 into a spreadsheet based "Budget & Actuals" / "Income Expense" system using a
 language I was comfortable with. I'm in the process of moving this to an SQL database in the cloud
 using Microsoft Azure.  I chose Azure because it is heavily used in the government
 A spreadsheet was used manually and previously, it was visually more attractive, but the structure
 was different for every vendor, thus it can not be automated

# Why automation?
# ---------------
 For a large file with more than 1 month of transactions, repetition, tedium, and time to complete
 categorization became issues. More discipline is required, discipline is easy with automation

# Budgeting
#  ---------
 Maintaining a good budget manually requires discipline.  There is a manual process for budgeting
 called envelopes that works well.  But it is only cash based, people don't write checks so much
 anymore and ability to work with credit cards is then required

 Knowing how much money is available in each category all the time is a required.
 For this to work, Actual data needs to be recorded daily not monthly.  This is a burden that requires
 actuals to be processed automatically and available all the time.

# What format to choose?
# ----------------------
 I've witnessed 4-5 methods supported by a few of the top financial institutions I am working with, AMEX,
 FI of America, and CitiFI.  The methods are QIF, QFX, OFX, Excel, and CSV.
 -------------------------------------------------------------------------------------------------
 QIF            |  is a tab based Quicken defined format although still supported, not the        |
                |  preferred choice by Quicken in the last 20 years                               |
 -------------------------------------------------------------------------------------------------|
 Excel          |  This format appears ideal based on the name, is not offered by many of the     |
                |  vendors I am working it.  Although still supported, the XLS format was         |
                |  abandoned by Microsoft 17 years ago in 2003 in favor of the XML based XLSX     |
                |  format. When attempting to use it, it could not be read by Excel without       |
                |  errors.                                                                        |
 -------------------------------------------------------------------------------------------------|
 CSV            |  as stated, this became the ideal for manual processing but structure and       |
                |  columns used is different for every vendor and can not be automated as no      |
                |  structure is defined                                                           |
 -------------------------------------------------------------------------------------------------|
 OFX            |  is the preferred method as it is open, well documented, available to everyone  |
                |  at no cost, but not offered by all the vendors without request                 |
 -------------------------------------------------------------------------------------------------|
 QFX            |  is available by my choice of vendors, fairly well defined not strict easy to   |
                |  understand machine based on OFX and readable XML.  So I chose QFX              |
 -------------------------------------------------------------------------------------------------|
# So here is the data flow
# -------------------------
1. Download Transactions from banks daily
2. Load in a data storage medium.  We will start with the existing "All Detailed Expenses" spreadsheet for now
3. Identify categories
4. Create combined budget/actuals - income/expenses report. For now, the report is still the master for available categories and budgets
5. Create immediate budget / actuals report for a single category for mobile phones

# Initial environment
# -------------------
1. Manually download transactions from banks daily
2. Manually import data into "All detailed expenses" spreadsheet
3. Manually categorize transactions
4. Budget/Actuals income/expenses report already exists as a spreadsheet, budgets and avaialble categories are defined static on the spreadsheet.
5. Immediae budget / actuals for mobile not started

# Directions
# --------
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
