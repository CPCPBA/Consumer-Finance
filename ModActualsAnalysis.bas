Attribute VB_Name = "ModActualsAnalysis"
Option Explicit
Option Base 1

' Global Variables and Constants

Const DEBUGSTATUS = True

Dim categoryNotFound As Boolean
Dim expensesSheet As Object ' the sheet we are studying
Dim expensesLastRow As Long
Dim lookupSheet As Object ' Contains table of descriptions and categories
Dim lookupLastRow As Integer ' last row of Lookup sheet
Dim rw As Integer ' current row of expense sheet

Enum errCriticality
  FTLERR = 0
  WRNERR = 1
  INFOERR = 2
End Enum

Sub readTransactions()
'---------------------------------------------------------------------------------------
' Procedure : readTransactions
' Author    : Christopher Prost, CP Business Analysis LLC. (9/21/2020)
' Website   : http://www.cpbusinessanalysis.com
' Copyright : 2020 CP Business Analysis LLC.  All Rights Reserved.
' Purpose   : read actual income and expenses from bank transactions QFX files
'
' Usage:
' ------
' readTransactions
'     input : QFX files in ~/downloads directory
'    output : 'Expense Detail' Sheet in 'Actuals Analysis' Workbook
'
' entry module
'---------------------------------------------------------------------------------------
' History
' This script was written in 3 days as a fairly quick means to read transaction files from banks and
' credit card companies into a spreadsheet based "Budget & Actuals" / "Income Expense" system using a
' language I was comfortable with. I'm in the process of moving this to an SQL database in the cloud
' using Microsoft Azure.  I chose Azure because it is heavily used in the government
' A spreadsheet was used manually and previously, it was visually more attractive, but the structure
' was different for every vendor, thus it can not be automated
'
' Why automation?
' ---------------
' For a large file with more than 1 month of transactions, repetition, tedium, and time to complete
' categorization became issues. More discipline is required
'
' Budgeting
' ---------
' Maintaining a good budget manually requires discipline.  There is a manual process for budgeting
' called envelopes that works well.  But it is only cash based, people don't write checks so much
' anymore and ability to work with credit cards is then required
'
' Knowing how much money is available in each category all the time is a required.
' For this to work, Actual data needs to be recorded daily not monthly.  This is a burden that requires
' actuals to be processed automatically and available all the time.
'
' What format to choose?
' ----------------------
' I've witnessed 4-5 methods supported by a few of the top financial institutions I am working with, AMEX,
' Bank of America, and CitiBank.  The methods are QIF, QFX, OFX, Excel, and CSV.
' -------------------------------------------------------------------------------------------------
' QIF            |  is a tab based Quicken defined format although still supported, not the        |
'                |  preferred choice by Quicken in the last 20 years                               |
' -------------------------------------------------------------------------------------------------|
' Excel          |  This format appears ideal based on the name, is not offered by many of the     |
'                |  vendors I am working it.  Although still supported, the XLS format was         |
'                |  abandoned by Microsoft 17 years ago in 2003 in favor of the XML based XLSX     |
'                |  format. When attempting to use it, it could not be read by Excel without       |
'                |  errors.                                                                        |
' -------------------------------------------------------------------------------------------------|
' CSV            |  as stated, this became the ideal for manual processing but structure and       |
'                |  columns used is different for every vendor and can not be automated as no      |
'                |  structure is defined                                                           |
' -------------------------------------------------------------------------------------------------|
' OFX            |  is the preferred method as it is open, well documented, available to everyone  |
'                |  at no cost, but not offered by all the vendors without request                 |
' -------------------------------------------------------------------------------------------------|
' QFX            |  is available by my choice of vendors, fairly well defined not strict easy to   |
'                |  understand machine based on OFX and readable XML.  So I chose QFX              |
' -------------------------------------------------------------------------------------------------|
  ' ModProc = "0101"

  Dim modProcErr As String
  Dim supportedFileTypes As String                ' A few 3 letter file types delimited by spaces
  Dim fileList As Collection                      ' list of file object of given format
  Dim fileStr As String                           ' contents of QFX file as a whole
  Dim banks As Collection
  Dim numTrans As Integer
  Dim FI As oBank
  Dim bank As oBank
  Dim fikey As String
  Dim fListCounter As Integer
  
  Set banks = loadBankInfo
  getFileList supportedFileTypes, fileList
  
  For fListCounter = fileList.count To 1 Step -1
    fileStr = fileList(fListCounter)
    Set FI = getBankInfo(fileStr)
    fikey = FI.FID & " " & FI.AccountNumber
    Set bank = banks(fikey)
    getNewTransactions fileStr, bank
  Next
  Set fileList = Nothing
  writeRecords banks
  
  
  
  
  GoTo theEnd

ErrorHandle:
  displayError Err.Number, Err.Description, "There was a system error. Contact User Support", FTLERR

theEnd:
End Sub

