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
Public Const EXPENSESSOURCECOL = 1                 ' source
Public Const EXPENSESMONTHCOL = 2                  ' Month
Public Const EXPENSESDATECOL = 3                   ' Date
Public Const EXPENSESDESCRIPTIONCOL = 4            ' description
Public Const EXPENSESMONTHCATEGORYCOL = 5          ' month category
Public Const EXPENSESCATEGORYCOL = 6               ' Category
Public Const EXPENSESCATEGORYTYPECOL = 7
Public Const EXPENSESAMOUNTCOL = 8
Public Const EXPENSESRUNNINGTOTALCOL = 9
Public Const EXPENSESCLEAREDCOL = 10
Public Const EXPENSESCLEAREDBALANCECOL = 11
Public Const EXPENSESFITIDCOL = 12                   'Financial Institute Transaction UUID





Sub Main()
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
  
  
  On Error GoTo errorHandleMain
  getExistingCategoryDescriptions "onThisBeautifulDay"
  Set FIs = loadFinancialInstitutions()            ' all supported financial institute accounts
  Set Filelist = getFileList("QFX")                ' collection of supported file contents, supported extentions are delimited by a space

  FIKey = "N/D"
  For fListCounter = 1 To Filelist.Count
    filestr = Filelist(fListCounter)
    FIKey = getFIInfo(filestr)
    On Error Resume Next
    Set FI = FIs(FIKey)
    If Err.Number = 0 Then
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

Sub writeRecords(FIs As Collection)
'---------------------------------------------------------------------------------------
' Procedure : writeRecords
' Author    : Christopher Prost, CP Business Analysis LLC. (9/21/2020)
' Website   : http://www.cpbusinessanalysis.com
' Copyright : 2020 CP Business Analysis LLC.  All Rights Reserved.
' Purpose   : Write the transaction records to the data management system and color them
'
' Usage:
' ------
' writeTransRecord
'     input : FI name and Transactions
'    output : this Workbook sheet 2 "Detailed Transactions"
'
' Called From
' ------------
' writeRecords
'---------------------------------------------------------------------------------------
  
  
  Dim FI As oFI
  Dim FIName As String
    
  On Error GoTo errorWriteRecords
 
  For Each FI In FIs
    FIName = FI.name
    writeTransRecords FI
    colorRecords FI
  Next FI

GoTo theEnd
errorWriteRecords:
  displayError Err.Number, Err.Description, "Error: Source: write Records, FI= " & FIName, FATALERR

theEnd:
End Sub

Sub writeTransRecords(FI As oFI)
'---------------------------------------------------------------------------------------
' Procedure : writeTransRecord
' Author    : Christopher Prost, CP Business Analysis LLC. (9/21/2020)
' Website   : http://www.cpbusinessanalysis.com
' Copyright : 2020 CP Business Analysis LLC.  All Rights Reserved.
' Purpose   : Write the transaction records to the data management system
'
' Usage:
' ------
' writeTransRecord
'     input : FI name and Transactions
'    output : this Workbook sheet 2 "Detailed Transactions"
'
' Called From
' ------------
' writeRecords
'---------------------------------------------------------------------------------------
  
 Dim FIName As String
 Dim trans As oTransaction
 Dim TransID As String
 Dim expensesSheet As Worksheet
 Dim lastrow As Long
 Dim rw As Long
 
 On Error GoTo ErrorHandlewriteTransRecord
  
  FIName = FI.name
  For Each trans In FI.Transactions
  
    lastrow = expensesSheet.Cells(Rows.Count, EXPENSESDESCRIPTIONCOL).End(xlUp).Row
    rw = lastrow
    If trans.Existing = False Then
      FIName = FI.name
      TransID = trans.FITID
      rw = rw + 1
      expensesSheet.Cells(rw, EXPENSESSOURCECOL).value = trans.Source
      expensesSheet.Cells(rw, EXPENSESMONTHCOL).value = Format(trans.postedDate, "mmm")
      expensesSheet.Cells(rw, EXPENSESDATECOL).value = trans.postedDate
      expensesSheet.Cells(rw, EXPENSESDESCRIPTIONCOL).value = trans.Description
      expensesSheet.Cells(rw, EXPENSESCATEGORYCOL).value = trans.category
      expensesSheet.Cells(rw, EXPENSESMONTHCATEGORYCOL).value = expensesSheet.Cells(rw, EXPENSESMONTHCOL).value & " " & expensesSheet.Cells(rw, EXPENSESCATEGORYCOL).value
      expensesSheet.Cells(rw, EXPENSESDESCRIPTIONCOL).value = trans.Description
      expensesSheet.Cells(rw, EXPENSESAMOUNTCOL).value = trans.amount
      expensesSheet.Cells(rw, EXPENSESFITIDCOL).value = trans.FITID
    End If
  Next

GoTo theEnd
ErrorHandlewriteTransRecord:
  displayError Err.Number, Err.Description, "Error: Source: write Trans Record, FI= " & FIName & ",TransID = " & TransID & ",Row= " & rw & ", lastRow = " & lastrow, FATALERR

theEnd:

End Sub



Sub colorRecords(FI As oFI)
'---------------------------------------------------------------------------------------
' Procedure : colorRecords
' Author    : Christopher Prost, CP Business Analysis LLC. (9/21/2020)
' Website   : http://www.cpbusinessanalysis.com
' Copyright : 2020 CP Business Analysis LLC.  All Rights Reserved.
' Purpose   : Color the records in spreadsheet by FI for easier reference
'
' Usage:
' ------
' getTransInfo
'     input : FI.BGColor and FI.FGColor and FI.Name
'    output : This workbook sheet 2 Transaction Detail
'
' Called From
' ------------
' writeRecords
'---------------------------------------------------------------------------------------
  
  Dim FIName As String
  Dim rw As Long
  Dim lastrow As Long
 
  On Error GoTo ErrorHandlecolorRecords
 
  lastrow = expensesSheet.Cells(Rows.Count, EXPENSESDESCRIPTIONCOL).End(xlUp).Row
  For rw = 2 To lastrow
    If (expensesSheet.Cells(rw, EXPENSESSOURCECOL).value = FI.name) Then
      expensesSheet.Range(expensesSheet.Cells(rw, EXPENSESSOURCECOL), expensesSheet.Cells(rw, EXPENSESFITIDCOL)).Interior.ColorIndex = FI.BGColorIndex
      expensesSheet.Range(expensesSheet.Cells(rw, EXPENSESSOURCECOL), expensesSheet.Cells(rw, EXPENSESFITIDCOL)).Font.ColorIndex = FI.FGColorIndex
    End If
  Next rw

GoTo theEnd
ErrorHandlecolorRecords:
  displayError Err.Number, Err.Description, "Error: Source: color Records, FI= " & FI.name & ", Row = " & rw & ", Lastrow = " & lastrow, FATALERR

theEnd:
End Sub
