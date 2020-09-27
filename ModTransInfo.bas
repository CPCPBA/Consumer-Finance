Attribute VB_Name = "ModTransInfo"
Private Const EXPENSESSOURCECOL = 1                 ' source
Private Const EXPENSEFITIDCOL = 12                   ' fid
Private Const EXPENSESMONTHCOL = 2                  ' Month
Private Const EXPENSESDATECOL = 3                   ' Date
Private Const EXPENSESDESCRIPTIONCOL = 4            ' description
Private Const EXPENSESMONTHCATEGORYCOL = 5          ' month category
Private Const EXPENSESCATEGORYCOL = 6               ' Category
Private Const EXPENSESCATEGORYTYPECOL = 7
Private Const EXPENSESAMOUNTCOL = 8
Private Const EXPENSESRUNNINGTOTALCOL = 9
Private Const EXPENSESCLEAREDCOL = 10
Private Const EXPENSESCLEAREDBALANCECOL = 11

Function getExistingBankTransactions(BankName As String) As Collection
  
  Dim coll As Collection
  Dim rw As Long
  Dim lastrw  As Long
  Dim trans As oTransaction
  Dim expensesSheet As Worksheet
  
  Set expensesSheet = ThisWorkbook.Worksheets(2)
  lastrw = expensesSheet.Cells(Rows.count, EXPENSESDESCRIPTIONCOL).End(xlUp).Row
  
  Set coll = New Collection
  
  For rw = 2 To lastrw
    If BankName = expenseSheet.Cells(rw, EXPENSESSOURCECOL).Value Then
      trans.Source = expenseSheet.Cells(rw, EXPENSESSOURCECOL).Value
      trans.FITID = expensesSheet.Cells(rw, EXPENSEFIDCOL).Value
      trans.postedDate = expensesSheet.Cells(rw, EXPENSESDATECOL).Value
      trans.Description = expensesSheet.Cells(rw, EXPENSESDESCRIPTIONCOL).Value
      trans.amount = expensesSheet.Cells(rw, EXPENSESAMOUNTCOL).Value
      trans.Existing = True
      coll.Add trans, trans.FITID
    End If
  Next rw
    
  Set getExistingBankTransactions = coll
End Function

Sub writeRecords(banks As Collection)
  
  Dim rw As Long
  Dim lastrw  As Long
  Dim trans As oTransaction
  Dim expensesSheet As Worksheet
  Dim bank As oBank
  Dim trans As oTransaction
  
  Set expensesSheet = ThisWorkbook.Worksheets(2)
  lastrw = expensesSheet.Cells(Rows.count, EXPENSESDESCRIPTIONCOL).End(xlUp).Row
    
  For Each bank In banks
    For Each trans In bank.TransactionList
      If trans.Existing = False Then
        expensesSheet.Cells(rw, EXPENSESSOURCECOL).Value = trans.Source
        expensesSheet.Cells(rw, EXPENSEFIDCOL).Value = trans.FITID
        expensesSheet.Cells(rw, EXPENSESDATECOL).Value = trans.postedDate
        expensesSheet.Cells(rw, EXPENSESDESCRIPTIONCOL).Value = trans.Description
        expensesSheet.Cells(rw, EXPENSESAMOUNTCOL).Value = trans.amount
      End If
    Next
  Next
  Set getExistingBankTransactions = coll
End Sub


Sub getNewTransactions(str As String, bank As oBank)
'---------------------------------------------------------------------------------------
' Procedure : getNewTransactions
' Author    : Christopher Prost, CP Business Analysis LLC. (9/21/2020)
' Website   : http://www.cpbusinessanalysis.com
' Copyright : 2020 CP Business Analysis LLC.  All Rights Reserved.
' Purpose   : read only new transaction data from QFX file put into an instance of class transaction data
'
' Usage:
' ------
' getTransInfo
'     input : str = entire contents of QFX or OFX file from bank
'     input : bank = bank information, about this FI
'    output : trans,all transactions read from QFX file
'
' readTransactions
'---------------------------------------------------------------------------------------Function strCount(str As String, searchStr As String) As Long
  
  Dim modProcErr As String
  Dim trans As oTransaction             ' new transaction
  Dim existingTrans As oTransaction
  Dim transList As Collection
  Dim start As Long                            ' starting point for next search
  Dim pDate As Date
  Dim newTransIndex As Long
  
  Dim char30000 As String
  Dim theRest As String
  Dim strStart As Long
  Dim strEnd As Long
  Dim count As Long
  
  ' ModProc = "0301"
  
  strStart = 1
  strEnd = InStr(30000, str, "<STMTTRN>")
  newTransIndex = InStr(char30000, "<STMTTRN>")
  char30000 = Mid(str, strStart, strEnd)
  theRest = Mid(str, strEnd, Len(str) - strEnd)
  count = 0
  While (Len(char30000) > 0)
    newTransIndex = InStr(char30000, "<STMTTRN>")
    While newTransIndex > 0
      start = newTransIndex
      Set trans = New oTransaction
      tmpstr = xmlfieldvalue(char30000, "<FITID>", start)
      count = count + 1
      If Not bank.TransactionList(tmpstr) Is Nothing Then
        trans.FITID = tmpstr
        tempStr = xmlfieldvalue(char30000, "<DTPOSTED>", start)
        trans.postedDate = CDate(Mid(tempStr, 5, 2) & "/" & Mid(tempStr, 7, 2) & "/" & Mid(tempStr, 1, 4))
        trans.amount = CCur(xmlfieldvalue(char30000, "<TRNAMT>", start)) * bank.DBCRdirection
        trans.Description = xmlfieldvalue(char30000, "<NAME>", start)
        trans.Source = bank.Name
        trans.Existing = False
        bank.TransactionList.Add trans, key:=trans.FITID
        Debug.Print "Posted " & count & ". " & trans.Source & " " & trans.FITID & " " & trans.postedDate & " " & trans.amount & " " & trans.Description
      Else
        Debug.Print "Duplicate transaction " & trans.Source & " " & trans.FITID & " " & trans.postedDate & " " & trans.amount & " " & trans.Description
      End If
      start = start + 1
      newTransIndex = InStr(start, char30000, "<STMTTRN>")
    Wend
    If (Len(theRest) <= 32000) Then
      char30000 = theRest
      theRest = ""
    Else
      strEnd = InStr(30000, theRest, "<STMTTRN>")
      char30000 = Mid(theRest, 1, strEnd)
      theRest = Mid(theRest, strEnd, Len(theRest) - strEnd)
    End If
  Wend
GoTo theEnd
ErrorHandle:
  displayError Err.Number, Err.Description, "There was a system error. Contact User Support", FTLERR

theEnd:

End Sub
'Function getTransInfoSaved(str As String, bank As BankInfo) As BankTransaction
''---------------------------------------------------------------------------------------
'' Procedure : getTransInfo
'' Author    : Christopher Prost, CP Business Analysis LLC. (9/21/2020)
'' Website   : http://www.cpbusinessanalysis.com
'' Copyright : 2020 CP Business Analysis LLC.  All Rights Reserved.
'' Purpose   : read transaction data from QFX file put into an instance of class transaction data
''
'' Usage:
'' ------
'' getTransInfo
''     input : str = entire contents of QFX or OFX file from bank
''     input : bank = bank information, about this FI
''    output : trans,all transactions read from QFX file
''
'' readTransactions
''---------------------------------------------------------------------------------------Function strCount(str As String, searchStr As String) As Long
'
'  Dim modProcErr As String
'  Dim singleTrans As tTrans                    ' The 4 values in each transaction
'  Dim transID As Long                          ' index to allTrans
'  Dim transExist As Boolean
'  Dim expensesSheet As Excel.Worksheet         ' sheet to be filled
'  Dim expensesLastRow As Long                  ' last row of 'Expenses Detail' sheet
'  Dim rw As Long                               ' current spreadsheet row
'  Dim dtStr As String                          ' date string in mm/dd/yy format for translation
'
'  Dim transAlreadyProcessed As Boolean         ' QFX file with same FIT & account number was exist
'  ' ModProc = "0301"
'
'  Set expensesSheet = ThisWorkbook.Sheets(2)
'
'  ' SIMULATE Ctrl + Shift + End FOR BOTH SHEETS
'  expensesLastRow = expensesSheet.Cells(Rows.Count, EXPENSESDESCRIPTIONCOL).End(xlUp).Row
'  rw = 2
'  While rw <= expensesLastRow
'    allTrans(rw - 1).Source = expensesSheet.Cells(rw, EXPENSESSOURCECOL).value
'    allTrans(rw - 1).fid = expensesSheet.Cells(rw, EXPENSEFIDCOL).value
'    allTrans(rw - 1).Date = expensesSheet.Cells(rw, EXPENSESDATECOL).value
'    allTrans(rw - 1).Description = expensesSheet.Cells(rw, EXPENSESDESCRIPTIONCOL).value
'    allTrans(rw - 1).Amount = expensesSheet.Cells(rw, EXPENSESAMOUNTCOL).value
'  Wend
'
'  While elementID <= UBound(elements)
'    parseTransInfo elements, elementID, singleTrans, transFound, dbcr
'
'  ' identify transaction was already processed
'    If transFound.fid And transFound.PostedDate And transFound.Description And transFound.Amount Then
'      transID = LBound(allTrans)
'      transExist = False
'      While (transID <= UBound(allTrans)) And (Not transExist)
'        If (allTrans(transID).Source & allTrans(transID).fid) = Source & singleTrans.fid Then
'          transExist = True
'        Else
'          transID = transID + 1
'        End If
'      Wend
'      If Not transExist Then
'        rw = rw + 1
'        allTrans(transID).Source = Source
'        allTrans(transID).fid = singleTrans.fid
'        allTrans(transID).Date = singleTrans.Date
'        allTrans(transID).Description = singleTrans.Description
'        allTrans(transID).Amount = singleTrans.Amount
'        expensesSheet.cell(rw.EXPENSESSOURCECOL).value = Source
'        expensesSheet.cell(rw.EXPENSEFIDCOL).value = singleTrans.fid
'        expensesSheet.cell(rw.EXPENSESDATECOL).value = singleTrans.Date
'        expensesSheet.cell(rw.EXPENSESDESCRIPTIONCOL).value = singleTrans.Description
'        expensesSheet.cell(rw.EXPENSESAMOUNTCOL).value = singleTrans.Amount
'      End If
'    End If
'  Wend
'
'
''GoTo theEnd
'ErrorHandle:
'  displayError err.Number, err.Description, "There was a system error. Contact User Support", FTLERR
'
'theEnd:
'
'End Function
'Sub categorize()
''---------------------------------------------------------------------------------------
'' Procedure : categorize
'' Author    : Christopher Prost, CP Business Analysis LLC. (7/9/2020)
'' Website   : http://www.cpbusinessanalysis.com
'' Copyright : 2020 CP Business Analysis LLC.  All Rights Reserved.
'' Purpose   : key phrases are subsets of transaction description.  find key phrase in table assign category to transaction
''
'' Usage:
'' ------
'' categorize
''     input : worksheet with expense detail
''     input : table of key phrases and categories,
''    output : transaction category
''
'' Called From:
'' ------------
'' TBD        : TBD
''---------------------------------------------------------------------------------------
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' This subroutine will lookup a known subset of the description and apply the pre-defined category.  It is likely many
'' descriptions will not have a category but this routine gets rid of most of the repetition
'' Written by Christopher Prost 7/9/2020
''
'' description is typically up to 1-5 keywords followed by a string consisting of unique vendor/location qualifier.
'' It could vary store location and/or date. We assume the first few words are related to the vendor
'' keyphrase = 1 to 5 keywords currently defined manually on lookup sheet
'' example description: "COSTCO GAS #03 06/11 PURCHASE ROSEVILLE MI".keyphrase "COSTCO GAS"
'' example description: "COSTCO WHSE #0 07/03 PURCHASE ROSEVILLE MI".keyphrase "COSTCO WHSE"
'' example descripiton: "KROGER #710 06/08 PURCHASE HARPER WOODS MI".keyphrase "KROGER"
''
'' method:start with the largest known key phrase for the description and reduce words until there is a match or left blank for manual fill
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Date     Who       Description of Change
'' -------- --------- --------------------------------------------------------------------------------------------------------
'' 07/09/20 Prost     Initial Release
'' 07/10/20 Prost     Colorized each row based on data source, easier to read
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  Dim modProcErr As String
'  Dim Source As String ' contents of current row, column 1
'  Dim Description As String ' contents of current row, column 4
'  Dim descrArray() As String ' broken up description
'  Dim descrSize As Integer ' number of words in description
'  Dim existingCategory As String ' existing contents of column 6
'  Dim expensesLastRow As Integer ' last row of sheet 2
'  Dim largestKeyPhraseCount As Integer ' max number of words used key Phrases
'  Dim keyPhraseSize As Integer ' current working keyPhrase number of words
'  Dim keyPhraseCount As Integer ' min of descrsize and largest key phrase count
'  Dim keyPhrase As String ' current keyphrase
'  Dim descrDelimitersArray() As Variant ' List of possible delimiters in description.Could be " ", "*" currently
'  Dim delimSize As Integer ' number of delimiters
'  Dim delimCount As Integer ' current Delimiter
'
'
'  Const LOOKUPCOLORSOURCECOL = 1
'  Const LOOKUPCOLORCOLORCOL = 2
'  Const LOOKUPCOLORFIRSTROW = 17
'  Const LOOKUPCOLORLASTROW = 22
'
'  ' ModProc = "0303"
'
'#If Not DEBUGSTATUS Then
'  On Error GoTo ErrorHandle
'#End If
'
'  Set expensesSheet = ThisWorkbook.Sheets(2)
'  Set lookupSheet = ThisWorkbook.Sheets(3)
'
'  ' SIMULATE Ctrl + Shift + End FOR BOTH SHEETS
'  expensesLastRow = expensesSheet.Cells(Rows.Count, EXPENSESDESCRIPTIONCOL).End(xlUp).Row
'  lookupLastRow = lookupSheet.Cells(lookupSheet.Rows.Count, LOOKUPKEYWORDSCOL).End(xlUp).Row
'
'  descrDelimitersArray = Array(" ", "*", "-")
'  delimSize = UBound(descrDelimitersArray)
'
'
'  ' Turn on error trapping
'  On Error Resume Next
'  err.Clear
'
'  largestKeyPhraseCount = WorksheetFunction.Max(lookupSheet.Range(lookupSheet.Cells(2, LOOKUPWORDCOUNTCOL), lookupSheet.Cells(lookupLastRow, LOOKUPWORDCOUNTCOL)))
'  For rw = 2 To expensesLastRow
'
'    ' Format Interior Cell Color For Ease Of Read.  Always gets messed up during manual edits
'    expensesSheet.Range(expensesSheet.Cells(rw, EXPENSESSOURCECOL), expensesSheet.Cells(rw, EXPENSESCLEAREDCOL)).Interior.ColorIndex = _
'        WorksheetFunction.VLookup(expensesSheet.Cells(rw, EXPENSESSOURCECOL).value, _
'        lookupSheet.Range(lookupSheet.Cells(LOOKUPCOLORFIRSTROW, LOOKUPCOLORSOURCECOL), lookupSheet.Cells(LOOKUPCOLORLASTROW, LOOKUPCOLORCOLORCOL)), 2, False)
'
'    Source = Cells(rw, EXPENSESSOURCECOL).value
'    Description = Cells(rw, EXPENSESDESCRIPTIONCOL).value
'
'
'
'    existingCategory = Cells(rw, EXPENSESCATEGORYCOL).value
'    If existingCategory = "" Then
'      categoryNotFound = True
'      delimCount = 1
'      While (delimCount <= delimSize) And categoryNotFound
'        Erase descrArray
'        descrArray = Split(Trim(Description), descrDelimitersArray(delimCount))
'        descrSize = UBound(descrArray)
'        keyPhraseCount = WorksheetFunction.Min(descrSize + 1, largestKeyPhraseCount + 1)
'
'        ' wierd delimiters are rare by vendor and only on some web purchases, an exception is coded only to pull off the first field
'        ' so we know and therefore set the max number of words in the keyphrase immediately to 1
'
'        If ((descrDelimitersArray(delimCount) = "*") Or (descrDelimitersArray(delimCount) = "-")) Then
'          keyPhraseCount = 1
'        End If
'
'        While (keyPhraseCount > 0) And categoryNotFound
'          keyPhraseCount = keyPhraseCount - 1
'          keyPhrase = ""
'          For keyPhraseSize = 0 To keyPhraseCount
'            keyPhrase = keyPhrase & descrArray(keyPhraseSize)
'            If (keyPhraseSize < keyPhraseCount) Then
'              keyPhrase = keyPhrase & " "
'            End If
'          Next ' keyphrase size
'          autoCategorize (keyPhrase)
'        Wend ' keyphrasecount > 0 and categoryNotFound
'
'        delimCount = delimCount + 1
'      Wend ' delimcount
'    End If
'  Next rw
'
'GoTo theEnd
'ErrorHandle:
'  displayError err.Number, err.Description, "There was a system error. Contact User Support", FTLERR
'
'theEnd:
'End Sub
'Sub autoCategorize(str As String)
''---------------------------------------------------------------------------------------
'' Procedure : autoCategorize
'' Author    : Christopher Prost, CP Business Analysis LLC. (7/9/2020)
'' Website   : http://www.cpbusinessanalysis.com
'' Copyright : 2020 CP Business Analysis LLC.  All Rights Reserved.
'' Purpose   : Looks up existing keyword and populates category field
''
'' Usage:
'' ------
'' autoCategorize
''     input : str,subset of transaction description
''    output : transaction category
''
'' Called From:
'' ------------
'' TBD        : TBD
''---------------------------------------------------------------------------------------
'
'
'  Dim modProcErr As String
'  Dim categoryLookItUp As Variant
'
'#If Not DEBUGSTATUS Then
'  On Error GoTo ErrorHandle
'#End If
'  ' ModProc = "0304"
'  err.Clear
'
'  categoryLookItUp = WorksheetFunction.VLookup(UCase(str), lookupSheet.Range(lookupSheet.Cells(2, LOOKUPKEYWORDSCOL), lookupSheet.Cells(lookupLastRow, LOOKUPCATEGORYCOL)), 2, False)
'
'  If (err.Number = 0) Or (err.Number = 13) Then
'    expensesSheet.Cells(rw, EXPENSESCATEGORYCOL).value = categoryLookItUp
'    categoryNotFound = False
'  End If
'
'GoTo theEnd
'ErrorHandle:
'  displayError err.Number, err.Description, "There was a system error. Contact User Support", FTLERR
'
'theEnd:
'End Sub
'
