Attribute VB_Name = "ModWriteRecords"
Option Explicit

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
' Main
'---------------------------------------------------------------------------------------
  
  
  Dim fi As oFI
  Dim FIName As String
    
  On Error GoTo errorWriteRecords
 
  For Each fi In FIs
    FIName = fi.name
    writeTransRecords fi
    colorRecords fi
  Next fi

GoTo theEnd
errorWriteRecords:
  displayError Err.Number, Err.Description, "Error: Source: write Records, FI= " & FIName, FATALERR

theEnd:
End Sub

Sub writeTransRecords(fi As oFI)
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
  
 Set expensesSheet = ThisWorkbook.Worksheets(2)
 lastrow = expensesSheet.Cells(Rows.Count, EXPENSESDESCRIPTIONCOL).End(xlUp).Row

  rw = lastrow

  FIName = fi.name
  For Each trans In fi.Transactions
  
    If trans.Existing = False Then
      FIName = fi.name
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


Sub colorRecords(fi As oFI)
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
  Dim expensesSheet As Worksheet
 
  On Error GoTo ErrorHandlecolorRecords
 
  Set expensesSheet = ThisWorkbook.Worksheets(2)
  lastrow = expensesSheet.Cells(Rows.Count, EXPENSESDESCRIPTIONCOL).End(xlUp).Row
  For rw = 2 To lastrow
    If (expensesSheet.Cells(rw, EXPENSESSOURCECOL).value = fi.name) Then
      expensesSheet.Range(expensesSheet.Cells(rw, EXPENSESSOURCECOL), expensesSheet.Cells(rw, EXPENSESFITIDCOL)).Interior.ColorIndex = fi.BGColorIndex
      expensesSheet.Range(expensesSheet.Cells(rw, EXPENSESSOURCECOL), expensesSheet.Cells(rw, EXPENSESFITIDCOL)).Font.ColorIndex = fi.FGColorIndex
    End If
  Next rw

GoTo theEnd
ErrorHandlecolorRecords:
  displayError Err.Number, Err.Description, "Error: Source: color Records, FI= " & fi.name & ", Row = " & rw & ", Lastrow = " & lastrow, FATALERR

theEnd:
End Sub

