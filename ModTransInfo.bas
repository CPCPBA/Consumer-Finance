Attribute VB_Name = "ModTransInfo"
Option Explicit

Sub getExistingFITransactions(FIName As String, Transactions As Collection)
'---------------------------------------------------------------------------------------
' Procedure : getExistingFITransactions
' Author    : Christopher Prost, CP Business Analysis LLC. (9/21/2020)
' Website   : http://www.cpbusinessanalysis.com
' Copyright : 2020 CP Business Analysis LLC.  All Rights Reserved.
' Purpose   : Read existing transactions for each FI from data management system copy into collection of transactions
'             for faster processing
'
' Usage:
' ------
' getExistingFITransactions
'     input : Financial Institute Name, string to determine which records are desired
'     input : This workbook, sheet 2 detailed transactions
'    output : collection of transactions
'
' Called From
' ------------
' main
'---------------------------------------------------------------------------------------
  
'  Dim transactions As Collection
  Dim rw As Long
  Dim lastrw  As Long
  Dim trans As oTransaction
  Dim expensesSheet As Worksheet
  Dim transIndex As Long
  
  
 On Error GoTo ErrorHandleGetExistingFITransactions
 
  
  transIndex = 0
  rw = 0
  
  Set expensesSheet = ThisWorkbook.Worksheets(2)
  lastrw = expensesSheet.Cells(Rows.Count, EXPENSESDESCRIPTIONCOL).End(xlUp).Row
  
'  Set transactions = New Collection
  Set trans = New oTransaction
  
  For rw = 2 To lastrw
    If FIName = expensesSheet.Cells(rw, EXPENSESSOURCECOL).value Then
      transIndex = transIndex + 1
      trans.Index = transIndex
      trans.Source = expensesSheet.Cells(rw, EXPENSESSOURCECOL).value
      trans.postedDate = expensesSheet.Cells(rw, EXPENSESDATECOL).value
      trans.Description = expensesSheet.Cells(rw, EXPENSESDESCRIPTIONCOL).value
      trans.category = expensesSheet.Cells(rw, EXPENSESCATEGORYCOL).value
      trans.amount = expensesSheet.Cells(rw, EXPENSESAMOUNTCOL).value
      trans.FITID = CStr(expensesSheet.Cells(rw, EXPENSESFITIDCOL).value)
      trans.Existing = True
      Transactions.Add trans, trans.FITID
    End If
  Next rw


GoTo theEnd
ErrorHandleGetExistingFITransactions:
  displayError Err.Number, Err.Description, "Error: Source: get Existing FI Transactions, FI= " & FIName & ", transindex = " & transIndex & ", Row = " & rw, FATALERR

theEnd:
End Sub



Sub getNewTransactions(str As String, FI As oFI)
'---------------------------------------------------------------------------------------
' Procedure : getNewTransactions
' Author    : Christopher Prost, CP Business Analysis LLC. (9/21/2020)
' Website   : http://www.cpbusinessanalysis.com
' Copyright : 2020 CP Business Analysis LLC.  All Rights Reserved.
' Purpose   : read only new transaction data from QFX file put into an instance of class transaction data
'
' Usage:
' ------
' getNewTransactions
'     input : str = entire contents of QFX or OFX file from FI
'     input : FI = FI information, about this FI
'    output : trans,all transactions read from QFX file
'
' Called From
' ------------
' main
'---------------------------------------------------------------------------------------
  Dim char30000 As String                 ' str broken up into 30000 char chunks do to limitation of integers used in instr function
  Dim theRest As String                   ' the rest of str = str - char30000
  Dim theBigData As String                ' Temp holding spot for char30000 data
  Dim FIName As String                    ' FI name Gives me some idea what file I'm working with
  
  On Error GoTo errorHandleGetNewTransactions
 
  FIName = FI.name
  createChar30000 str, char30000, theRest
   
  While (Len(char30000) > 0)
    getActualTransactions char30000, FI
    theBigData = theRest
    createChar30000 theBigData, char30000, theRest
  Wend

GoTo theEnd
errorHandleGetNewTransactions:
  displayError Err.Number, Err.Description, "Error: Source: get New Transactions, FI= " & FI.name, FATALERR

theEnd:
End Sub
Sub createChar30000(BiggerData As String, char30000 As String, theRest As String)
'---------------------------------------------------------------------------------------
' Procedure : createChar3000
' Author    : Christopher Prost, CP Business Analysis LLC. (9/21/2020)
' Website   : http://www.cpbusinessanalysis.com
' Copyright : 2020 CP Business Analysis LLC.  All Rights Reserved.
' Purpose   : pull the first 30,000 chars from Bigger data and put the rest in TheRest
'
' Usage:
' ------
' getTransInfo
'     input : BiggerData, remainder of QFX file
'    output : char30000, the next 30000 char of bigger data + the remainder of the transaction
'    output : theRest, the remainder of biggerData after char30000
'
' Called From
' -----------
' getNewTransactions
'---------------------------------------------------------------------------------------
  Dim strEnd As Long                         ' the desired end of char30000
    
  On Error GoTo errorHandleCreateChar3000
   
  If Len(BiggerData) > 30000 Then
    strEnd = InStr(30000, BiggerData, "<STMTTRN>")           ' instr function inStrStartparameter is an integer and limited to 2^15 (32000+)
    theRest = Mid(BiggerData, strEnd, Len(BiggerData) - strEnd)
  Else
    strEnd = Len(BiggerData)
    theRest = ""
  End If
  char30000 = Mid(BiggerData, 1, strEnd)
  

GoTo theEnd
errorHandleCreateChar3000:
  displayError Err.Number, Err.Description, "Error: Source: create Char3000, size of Big data = " & Len(BiggerData) & ", size of char30000 = ," & Len(char30000) & ", size of theRest = " & Len(theRest), FATALERR

theEnd:
End Sub
Sub getActualTransactions(str As String, FI As oFI)
'---------------------------------------------------------------------------------------
' Procedure : getActualTransaction
' Author    : Christopher Prost, CP Business Analysis LLC. (9/21/2020)
' Website   : http://www.cpbusinessanalysis.com
' Copyright : 2020 CP Business Analysis LLC.  All Rights Reserved.
' Purpose   : Copy data from STMTRNS section of transactionfile
'
' Usage:
' ------
' getActualTransactions
'     input : str = some section of QFX or OFX file
'     input : FI = FI information, about this FI
'    output : transactions,as part of FI
'
' Called From
' ------------------
' getNewTransactions
'---------------------------------------------------------------------------------------
  Dim newTransIndex As Long                '
  Dim strEnd As Long                      ' end of char30000
  Dim trans As oTransaction               ' A new transaction
  Dim FITemp As String                    ' Temporary FID holder
  Dim DtTemp As String                    ' tempoary PDate holder
  Dim inStrStart As Long                  ' start index for instr function
  Dim FIName As String                    ' FI.name
  Dim TransID As Integer                  ' current transaction index

 On Error GoTo errorHandleGetActualTransactions

  FIName = FI.name
  inStrStart = 1
  newTransIndex = InStr(inStrStart, str, "<STMTTRN>")
  While (newTransIndex > 0)
    inStrStart = newTransIndex
    FITemp = xmlfieldvalue(str, "<FITID>", inStrStart)
    If Not IsRepeatedTrans(FI, FITemp) Then
      Set trans = New oTransaction
      TransID = FI.Transactions.Count
      trans.Index = TransID
      trans.FITID = FITemp
      DtTemp = xmlfieldvalue(str, "<DTPOSTED>", inStrStart)
      trans.postedDate = CDate(Mid(DtTemp, 5, 2) & "/" & Mid(DtTemp, 7, 2) & "/" & Mid(DtTemp, 1, 4))
      trans.amount = CCur(xmlfieldvalue(str, "<TRNAMT>", inStrStart)) * FI.DBCRdirection
      trans.Description = xmlfieldvalue(str, "<NAME>", inStrStart)
      trans.category = findCategory(trans.Description)
      trans.Source = FIName
      trans.Existing = False
      FI.Transactions.Add trans, Key:=trans.FITID
      ' Debug.Print "Posted " & trans.Index & ". " & trans.Source & " " & trans.postedDate & " " & trans.amount & " " & trans.Description
    End If
    inStrStart = inStrStart + 1
    newTransIndex = InStr(inStrStart, str, "<STMTTRN>")
  Wend
  
  
GoTo theEnd
errorHandleGetActualTransactions:

  displayError Err.Number, Err.Description, "Error: Source:Get Actual Transactions, FI = " & FIName & ",Transaction ID = " & TransID & ", FITID = ," & FITemp, FATALERR

theEnd:
End Sub
Function IsRepeatedTrans(FI As oFI, FITID As String) As Boolean
'---------------------------------------------------------------------------------------
' Procedure : IsRepeatedTrans
' Author    : Christopher Prost, CP Business Analysis LLC. (9/21/2020)
' Website   : http://www.cpbusinessanalysis.com
' Copyright : 2020 CP Business Analysis LLC.  All Rights Reserved.
' Purpose   : read only new transaction data from QFX file put into an instance of class transaction data
'
' Usage:
' ------
' IsRepeatedTrans
'     input : FI object reprsenting entire Financial Instituion instance - whole object is desired for debugging
'     input : FITID, string Transactions's unique ID
'    output : Boolean value if record already exists in database
'
' Called From
' -------------
' get Actual Transactions
'---------------------------------------------------------------------------------------

  Dim trans As oTransaction               ' A new transaction
  Dim foundTrans As oTransaction          ' A reference to a repeated transaction
  Dim repeated As Boolean                 ' Found transaction flag
  Dim Transactions As Collection          ' A collection of transaction found
  Dim FIName As String                    ' Name of FI

  On Error GoTo errorHandleIsRepeatedTrans
 
  If FI.Transactions.Count > 0 Then
    repeated = False
    On Error Resume Next
    Set foundTrans = FI.Transactions.Item(FITID)
    If Err.Number = 0 Then   ' record is found
      ' Debug.Print "Found repeated transaction " & foundTrans.Index & ". " & foundTrans.postedDate & " " & foundTrans.amount & " " & foundTrans.Description
      repeated = True
    End If
    Err.Clear
    On Error GoTo errorHandleIsRepeatedTrans
  End If

GoTo theEnd
errorHandleIsRepeatedTrans:
  displayError Err.Number, Err.Description, "Error: Source: Is Repeated Trans, FI= " & FIName & ", FITID = " & FITID, FATALERR
theEnd:
End Function


