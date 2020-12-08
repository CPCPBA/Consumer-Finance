Attribute VB_Name = "ModTransInfo"
Option Explicit

Sub getExistingFITransactions(finame As String, Transactions As Collection)
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
  Dim lastRw  As Long
  Dim trans As oTransaction
  Dim expensesSheet As Worksheet
  Dim transIndex As Long
  
  
 On Error GoTo ErrorHandleGetExistingFITransactions
 
  
  transIndex = 0
  rw = 0
  
  Set expensesSheet = ThisWorkbook.Worksheets(3)
  lastRw = expensesSheet.Cells(Rows.Count, EXPENSESDESCRIPTIONCOL).End(xlUp).Row
  
'  Set transactions = New Collection
  Set trans = New oTransaction
  
  For rw = 2 To lastRw
    If finame = expensesSheet.Cells(rw, EXPENSESSOURCECOL).value Then
      transIndex = transIndex + 1
      trans.index = transIndex
      trans.Source = expensesSheet.Cells(rw, EXPENSESSOURCECOL).value
      trans.postedDate = expensesSheet.Cells(rw, EXPENSESDATECOL).value
      trans.Description = expensesSheet.Cells(rw, EXPENSESDESCRIPTIONCOL).value
      trans.category = expensesSheet.Cells(rw, EXPENSESCATEGORYCOL).value
      trans.amount = expensesSheet.Cells(rw, EXPENSESAMOUNTCOL).value
      trans.transID = trans.Source & trans.postedDate & trans.Description & trans.amount

      trans.Existing = True
      Transactions.Add trans, trans.transID
    End If
  Next rw


GoTo theEnd
ErrorHandleGetExistingFITransactions:
  displayError Err.Number, Err.Description, "Error: Source: get Existing FI Transactions, FI= " & finame & ", transindex = " & transIndex & ", Row = " & rw, FATALERR

theEnd:
End Sub



Sub getNewTransactions(f As myFile, FI As oFI)
'---------------------------------------------------------------------------------------
' Procedure : getNewTransactions
' Author    : Christopher Prost, CP Business Analysis LLC. (9/21/2020)
' Website   : http://www.cpbusinessanalysis.com
' Copyright : 2020 CP Business Analysis LLC.  All Rights Reserved.
' Purpose   : read only new unique transaction data from QFX file put into an instance of class transaction data
'             only compare new transactions to existing transactions.  Assume no duplicates in QFX transactions and
'             existing transactions are unique amongst each other
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
  Dim finame As String                    ' FI name Gives me some idea what file I'm working with
  Dim regEx As Object
  Dim theMatches As MatchCollection
  Dim matchStr As match
  Dim i As Integer
  Dim transStart As Long
  Dim transEnd As Long
  Dim transSource As String
  Dim numTrans As Integer
  Dim fname As String
  Dim str As String
  
  On Error GoTo errorHandleGetNewTransactions
 
  fname = f.filename
  str = f.fileContents
  
  Set regEx = New RegExp
  regEx.Pattern = "<STMTTRN>"
  regEx.Global = True
  regEx.IgnoreCase = True

  Set theMatches = regEx.Execute(str)
  numTrans = theMatches.Count
  Debug.Print "Found " & numTrans & " Transactions"
  i = 0
  While i < theMatches.Count
    transStart = theMatches.Item(i).FirstIndex + 1
    If i = (theMatches.Count - 1) Then
      transEnd = Len(str)
    Else
      transEnd = theMatches.Item(i + 1).FirstIndex - 1
    End If
    transSource = Mid(str, transStart, (transEnd - transStart) + 1)
    If Not setTransaction(transSource, FI, fname) Then
      numTrans = numTrans - 1
      Debug.Print "Expected new added transactions is now " & numTrans
    End If
    i = i + 1
  Wend

GoTo theEnd
errorHandleGetNewTransactions:
  displayError Err.Number, Err.Description, "Error: Source: get New Transactions, FI= " & FI.name, FATALERR

theEnd:
End Sub

Function setTransaction(str As String, FI As oFI, filename As String) As Boolean
'---------------------------------------------------------------------------------------
' Procedure : setTransaction
' Author    : Christopher Prost, CP Business Analysis LLC. (11/21/2020)
' Website   : http://www.cpbusinessanalysis.com
' Copyright : 2020 CP Business Analysis LLC.  All Rights Reserved.
' Purpose   : Copy data from STMTRNS section of transactionfile
'
' Usage:
' ------
' setTransaction
'     input : str = some section of QFX or OFX file
'     input : FI = FI information, about this FI
'    output : transaction,as part of FI
'    output : boolean
'
' Called From
' ------------------
' getNewTransactions
'---------------------------------------------------------------------------------------
  Dim newTransIndex As Long                '
  Dim trans As oTransaction               ' A new transaction
  Dim transID As String                    ' Temporary FID holder
  Dim DtTemp As String                    ' tempoary PDate holder
  Dim inStrStart As Long                  ' start index for instr function
  Dim finame As String                    ' FI.name
  Dim index As Integer                  ' current transaction index
  Dim amount As Currency
  Dim descr As String
  Dim postedDate As Date
  
  Dim foundtrans As oTransaction
  Dim repeatedTrans As Boolean
  Dim lastWordDescr As String
  Dim increment As Integer
  Dim descrprefix As String

  On Error GoTo errorHandleGetActualTransactions

  finame = FI.name
  inStrStart = 1
  DtTemp = xmlfieldvalue(str, "<DTPOSTED>", inStrStart)
  postedDate = CDate(Mid(DtTemp, 5, 2) & "/" & Mid(DtTemp, 7, 2) & "/" & Mid(DtTemp, 1, 4))
  amount = CCur(xmlfieldvalue(str, "<TRNAMT>", inStrStart))
  descr = xmlfieldvalue(str, "<NAME>", inStrStart)
  transID = finame & postedDate & descr & amount         ' does not protect against same FI, similar product, same day, same description, same cost
                                                         ' if repeated trans based on this ID are found in same trans file, an increment will be applied to additional trans

  If FI.Transactions.Count > 0 Then
    repeatedTrans = False              ' assumption. now for the test
    On Error Resume Next
    Set foundtrans = FI.Transactions.Item(transID)   ' Financial Institution's Tranaction ID
    If Not foundtrans Is Nothing Then                  ' previous transaction same descr, same cost, same day is found
      repeatedTrans = True
      If (foundtrans.Existing) Or (foundtrans.transFile <> filename) Then                      ' this is a real repeated transaction
        Debug.Print "Found repeated transaction in " & foundtrans.transFile & ": " & foundtrans.index & ". " & foundtrans.postedDate & " " & foundtrans.amount & " " & foundtrans.Description & vbLf & _
                    "                              " & filename & ": " & postedDate & " " & amount & " " & descr
        Set foundtrans = Nothing
      Else
      
      ' repeated trans using app transid but not using Financial Institute Trans ID.  change description and transid until there is no repeated trans
      ' financial institute trans id (FITID) usually contains a transaction index of the number of transactions within the file and thus only unique to the file
        
        While repeatedTrans
          lastWordDescr = Right(descr, Len(descr) - InStrRev(descr, " "))
          If (Left(lastWordDescr, 2) = "-i") And (IsNumeric(Right(lastWordDescr, Len(lastWordDescr) - 1))) Then
            increment = Right(lastWordDescr, Len(lastWordDescr) - 1)
            descrprefix = Left(foundtrans.Description, InStr(foundtrans.Description, lastWordDescr))
          Else
            descrprefix = descr
            increment = 0
          End If
          descr = descrprefix & " -i" & (increment + 1)
          transID = finame & postedDate & descr & amount
          Set foundtrans = Nothing
          Set foundtrans = FI.Transactions.Item(transID)
          If foundtrans Is Nothing Then                  ' previous transaction same descr, same cost, same day is found
            repeatedTrans = False
          End If   ' foundtrans nothing
        Wend       ' repeatedtrans
      End If       ' existingtrans in spreadsheet
    End If         ' not foundtrans nothing
  End If           ' transactions.count > 0
  
  
  
  If Not repeatedTrans Then
    Set trans = New oTransaction
    trans.index = index
    trans.transID = transID
    trans.postedDate = CDate(Mid(DtTemp, 5, 2) & "/" & Mid(DtTemp, 7, 2) & "/" & Mid(DtTemp, 1, 4))
    trans.Description = descr
    trans.category = findCategory(trans.Description)
    trans.amount = amount
    trans.Source = finame
    trans.transFile = filename
    trans.Existing = False
    FI.Transactions.Add trans, Key:=trans.transID
    ' Debug.Print "Posted " & trans.Index & ". " & trans.Source & " " & trans.postedDate & " " & trans.amount & " " & trans.Description
    setTransaction = True
  Else
    setTransaction = False
  End If
  
  GoTo theEnd
  
errorHandleGetActualTransactions:

displayError Err.Number, Err.Description, "Error: Source:Get Actual Transactions, FI = " & finame & ",Transaction Index = " & index & ", transID = ," & transID, FATALERR

theEnd:
End Function



