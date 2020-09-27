Attribute VB_Name = "ModBankInfo"
Option Explicit

' ModProcErr is used in every procedure and is an attempt to quickly locate an error message within a project
' It is comprised of a 6 digit concatenation of 2 digit module index, procedure index, and err index
' Mod is a 2 digit index of modules in alphabetic order within current projectin alphabetic Order
' procedure is a 2 digit index of prodedure in current module in list order
' err is a 2 digit index of error ID within any current procedure

Function loadBankInfo() As Collection
  Dim BankList As Collection
  
  Dim bank As oBank
  Dim trans As oTransaction
  Dim FIDAcctID As String
 
  Set BankList = New Collection

  Set bank = New oBank
  Set trans = New oTransaction
  bank.Name = "Bank of America Checking"
  bank.AccountNumber = "9252"
  bank.FID = "5959"
  bank.DBCRdirection = 1
  bank.TransactionList.Add getExistingBankTransactions(bank.Name)
  FIDAcctID = bank.FID & " " & bank.AccountNumber
  BankList.Add bank, key:=FIDAcctID
  
  Set bank = New oBank
  bank.Name = "AMEX"
  bank.AccountNumber = "3006"
  bank.FID = "3101"
  bank.DBCRdirection = -1
  bank.TransactionList.Add getExistingBankTransactions(bank.Name)
  FIDAcctID = bank.FID & " " & bank.AccountNumber
  BankList.Add bank, key:=FIDAcctID
   
  Set bank = New oBank
  bank.Name = "Costco Visa"
  bank.AccountNumber = "8590"
  bank.FID = "2102"
  bank.DBCRdirection = -1
  bank.TransactionList.Add getExistingBankTransactions(bank.Name)
  FIDAcctID = bank.FID & " " & bank.AccountNumber
  BankList.Add bank, key:=FIDAcctID
   
  Set bank = New oBank
  bank.Name = "AAvantage Mastcard"
  bank.AccountNumber = "8379"
  bank.FID = "2102"
  bank.DBCRdirection = -1
  bank.TransactionList.Add getExistingBankTransactions(bank.Name)
  FIDAcctID = bank.FID & " " & bank.AccountNumber
  BankList.Add bank, key:=FIDAcctID
  

  Set loadBankInfo = BankList
  
End Function

Function getBankInfo(fileStr As String) As oBank
'---------------------------------------------------------------------------------------
' Procedure : getBankInfo
' Author    : Christopher Prost, CP Business Analysis LLC. (9/21/2020)
' Website   : http://www.cpbusinessanalysis.com
' Copyright : 2020 CP Business Analysis LLC.  All Rights Reserved.
' Purpose   : put bankInfo into bankInfo array
'
' Usage:
' ------
' getBankInfo
'     input   : elements, contents of 1 QFX file
'     output  : instance of a bank record
'
' readTransactions
'---------------------------------------------------------------------------------------
      
  Dim modProcErr As String
  Dim bankEntry As oBank
  Dim StrPos As Long                        ' liststr position
  Dim AcctID As String
  
 
  ' ModProc = "0201"
  
  Set bankEntry = New oBank
  StrPos = 1
  bankEntry.FID = xmlfieldvalue(fileStr, "<FID>", StrPos)
  AcctID = xmlfieldvalue(fileStr, "<ACCTID>", StrPos)
  bankEntry.AccountNumber = Right(AcctID, 4)
  
  Set getBankInfo = bankEntry
  
  GoTo theEnd

ErrorHandle:
  displayError Err.Number, Err.Description, "There was a system error. Contact User Support", FTLERR

theEnd:
End Function

