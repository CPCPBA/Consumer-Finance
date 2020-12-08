Attribute VB_Name = "Main"
Option Explicit

' Global Variables and Constants

Enum errCriticality
  FATALERR = 0
  WARNERR = 1
  INFOERR = 2
End Enum

' The following variables and constants are used in a few divergent places
' placed here instead of carrying them everywhere and rarely using them

' Transaction Processing
' ----------------------
Public Const EXPENSESFIRSTCOL = 1                  ' source
Public Const EXPENSESLASTCOL = 13                  ' CATEGORY GROUP
Public Const EXPENSESSOURCECOL = 1                 ' source
Public Const EXPENSESMONTHCOL = 2                  ' Month
Public Const EXPENSESDATECOL = 3                   ' Date
Public Const EXPENSESDESCRIPTIONCOL = 4            ' description
Public Const EXPENSESMONTHCATEGORYCOL = 5          ' month category
Public Const EXPENSESCATEGORYCOL = 6               ' Category
Public Const EXPENSESAMOUNTCOL = 7
Public Const EXPENSESRUNNINGTOTALCOL = 8
Public Const EXPENSERECONCILIATIONDATECOL = 9
Public Const EXPENSERECONCILIATIONTRANSIDCOL = 10
Public Const EXPENSESCLEAREDBALANCECOL = 11
Public Const EXPENSESTRANSIDCOL = 12                   'Financial Institute Transaction UUID
Public Const EXPENSESCATEGORYGROUP = 13                   'Grouping


Sub categorize()

  Dim expensesSheet As Worksheet
  Dim rw As Long
  Dim lastrow As Long
  
  Set expensesSheet = ThisWorkbook.Sheets(3)
  lastrow = expensesSheet.Cells(Rows.Count, EXPENSESDESCRIPTIONCOL).End(xlUp).Row
  getExistingCategoryDescriptions
  
  For rw = 2 To lastrow
    If (expensesSheet.Cells(rw, EXPENSESCATEGORYCOL).value = "N/F") Or (expensesSheet.Cells(rw, EXPENSESCATEGORYCOL).value = "") Then
      expensesSheet.Cells(rw, EXPENSESCATEGORYCOL).value = findCategory(expensesSheet.Cells(rw, EXPENSESDESCRIPTIONCOL).value)
'      expensesSheet.Cells(rw, EXPENSESMONTHCATEGORYCOL).value = expensesSheet.Cells(rw, EXPENSESMONTHCOL).value & " " & expensesSheet.Cells(rw, EXPENSESCATEGORYCOL).value
    End If
  Next rw
End Sub
Sub colorize()
  Dim expensesSheet As Worksheet
  Dim rw As Long
  Dim lastrow As Long
  Dim FIs As Collection     ' Collection of financial Institutions
  
  Dim NeedToLoadTransactions As Boolean
  Dim FI As oFI
  
  NeedToLoadTransactions = False
  Set FIs = loadFinancialInstitutions(NeedToLoadTransactions)    ' all supported financial institute accounts
  
  For Each FI In FIs
    colorRecords FI
  Next FI
  

End Sub

Sub main()
'---------------------------------------------------------------------------------------
' Procedure : Main
' Author    : Christopher Prost, CP Business Analysis LLC. (9/21/2020)
' Website   : http://www.cpbusinessanalysis.com
' Copyright : 2020 CP Business Analysis LLC.  All Rights Reserved.
' Purpose   : read actual income and expenses from FI transactions QFX files
'
' Usage:
' ------
' Main
'     input : QFX files in ~/downloads directory
'    output : 'Expense Detail' Sheet in 'Actuals Analysis' Workbook
'
' entry module
'---------------------------------------------------------------------------------------

  Dim Filelist As Collection                                ' list of file object of given format
  Dim FIs As Collection                                     ' Collection of financial Institutions
  Dim FIKey As String                                       ' FI and Acct in transaction file
  Dim fListCounter As Integer                               ' counter for Filelist
  Dim filestr As String
  Dim org As String
  Dim FI As oFI
  Dim categories As Collection
  Dim MaxNumberKeyWords As Integer
  Dim NeedToLoadExistingTransactions As Boolean
  
  
  On Error GoTo errorHandleMain
  getExistingCategoryDescriptions
  NeedToLoadExistingTransactions = True
  Set FIs = loadFinancialInstitutions(NeedToLoadExistingTransactions)    ' all supported financial institute accounts
  Set Filelist = getFileList("QFX")                        ' collection of supported file contents, supported extentions are delimited by a space
  FIKey = "N/D"
  For fListCounter = 1 To Filelist.Count
    filestr = Filelist(fListCounter).fileContents
    FIKey = getFIInfo(filestr)
    On Error Resume Next
    Set FI = FIs(FIKey)
    If Err.Number = 0 Then
      Debug.Print fListCounter & ". Processing Transactions for " & FI.name & " from " & Filelist(fListCounter).filename
      getNewTransactions Filelist(fListCounter), FIs(FIKey)
    Else
      org = xmlfieldvalue(Filelist(fListCounter), "<ORG>", 1)
      displayError Err.Number, Err.Description, "Error Reading Transaction Files.  " & org & " is not a supported Financial Institution", WARNERR
    End If
    Err.Clear
    On Error GoTo errorHandleMain
  Next
  Set Filelist = Nothing
  writeRecords FIs
  
GoTo theEnd
errorHandleMain:
  displayError Err.Number, Err.Description, "Error: Source: Main, FIKey= " & FIKey & ", FI= " & FI.name, FATALERR

theEnd:
End Sub
