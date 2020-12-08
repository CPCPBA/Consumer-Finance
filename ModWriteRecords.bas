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
  
  
  Dim FI As oFI
  Dim finame As String
    
  On Error GoTo errorWriteRecords
 
  For Each FI In FIs
    finame = FI.name
    writeTransRecords FI
    colorRecords FI
  Next FI

GoTo theEnd
errorWriteRecords:
  displayError Err.Number, Err.Description, "Error: Source: write Records, FI= " & finame, FATALERR

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
  
 Dim finame As String
 Dim trans As oTransaction
 Dim transID As String
 Dim expensesSheet As Worksheet
 Dim lastrow As Long
 Dim rw As Long
 
 On Error Resume Next ' ErrorHandlewriteTransRecord
  
 Set expensesSheet = ThisWorkbook.Worksheets(3)
 lastrow = expensesSheet.Cells(Rows.Count, EXPENSESDESCRIPTIONCOL).End(xlUp).Row

  rw = lastrow

  finame = FI.name
  For Each trans In FI.Transactions
  
    If trans.Existing = False Then
      finame = FI.name
      transID = trans.transID
      rw = rw + 1
      expensesSheet.Cells(rw, EXPENSESSOURCECOL).value = trans.Source
      expensesSheet.Cells(rw, EXPENSESMONTHCOL).value = Format(trans.postedDate, "mmm")
      expensesSheet.Cells(rw, EXPENSESDATECOL).value = trans.postedDate
      expensesSheet.Cells(rw, EXPENSESDESCRIPTIONCOL).value = trans.Description
      expensesSheet.Cells(rw, EXPENSESCATEGORYCOL).value = trans.category
      expensesSheet.Cells(rw, EXPENSESMONTHCATEGORYCOL).FormulaR1C1 = "=RC[" & (EXPENSESMONTHCOL - EXPENSESMONTHCATEGORYCOL) & "] & " & Chr(34) & " " & Chr(34) & " & " & "RC[" & (EXPENSESCATEGORYCOL - EXPENSESMONTHCATEGORYCOL) & "]"
      expensesSheet.Cells(rw, EXPENSESDESCRIPTIONCOL).value = trans.Description
      expensesSheet.Cells(rw, EXPENSESAMOUNTCOL).value = trans.amount
      expensesSheet.Cells(rw, EXPENSESTRANSIDCOL).FormulaR1C1 = "=RC[" & (EXPENSESSOURCECOL - EXPENSESTRANSIDCOL) & "] &" & _
                                                                "TEXT(RC[" & (EXPENSESDATECOL - EXPENSESTRANSIDCOL) & "]," & Chr(34) & "MMDDYYYY" & Chr(34) & ") & " & _
                                                                 "RC[" & (EXPENSESDESCRIPTIONCOL - EXPENSESTRANSIDCOL) & "] &" & _
                                                                 "RC[" & (EXPENSESAMOUNTCOL - EXPENSESTRANSIDCOL) & "]"
                                                                

    End If
  Next

GoTo theEnd
ErrorHandlewriteTransRecord:
  displayError Err.Number, Err.Description, "Error: Source: write Trans Record, FI= " & finame & ",TransID = " & transID & ",Row= " & rw & ", lastRow = " & lastrow, FATALERR

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
  
  Dim finame As String
  Dim rw As Long
  Dim lastrow As Long
  Dim expensesSheet As Worksheet
 
  On Error GoTo ErrorHandlecolorRecords
 
  Set expensesSheet = ThisWorkbook.Worksheets(3)
  lastrow = expensesSheet.Cells(Rows.Count, EXPENSESDESCRIPTIONCOL).End(xlUp).Row
  For rw = 2 To lastrow
    If (expensesSheet.Cells(rw, EXPENSESSOURCECOL).value = FI.name) Then
      expensesSheet.Range(expensesSheet.Cells(rw, EXPENSESFIRSTCOL), expensesSheet.Cells(rw, EXPENSESLASTCOL)).Interior.ColorIndex = FI.BGColorIndex
      expensesSheet.Range(expensesSheet.Cells(rw, EXPENSESFIRSTCOL), expensesSheet.Cells(rw, EXPENSESLASTCOL)).Font.ColorIndex = FI.FGColorIndex
    End If
  Next rw

GoTo theEnd
ErrorHandlecolorRecords:
  displayError Err.Number, Err.Description, "Error: Source: color Records, FI= " & FI.name & ", Row = " & rw & ", Lastrow = " & lastrow, FATALERR

theEnd:
End Sub

