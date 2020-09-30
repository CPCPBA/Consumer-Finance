Attribute VB_Name = "ModFIInfo"
Option Explicit


Function loadFinancialInstitutions() As Collection

On Error GoTo errorHandleLoadFinancialInstitutions
   
  Dim FIDAcctID As String                                  ' Quicken Financial Institute ID +  FI's last 4 Acct ID
  Dim FI As oFI                                            ' Financial Institute Object
  Dim FIList As Collection                                 ' Collection of FI Objects
   '----------------------------------------------------------------------------------
  Set FIList = New Collection
  
  Set FI = New oFI
  FI.name = "Bank of America"
  FI.AccountNumber = "9252"
  FI.FID = "5959"
  FI.DBCRdirection = 1
  getExistingFITransactions FI.name, FI.Transactions
  FI.BGColorIndex = 24
  FI.FGColorIndex = 3
  FIDAcctID = FI.FID & " " & FI.AccountNumber
  FIList.Add FI, Key:=FIDAcctID
  '----------------------------------------------------------------------------------
  Set FI = New oFI
  FI.name = "AMEX"
  FI.AccountNumber = "3006"
  FI.FID = "3101"
  FI.DBCRdirection = -1
  getExistingFITransactions FI.name, FI.Transactions
  FI.BGColorIndex = 35
  FI.FGColorIndex = 47
  FIDAcctID = FI.FID & " " & FI.AccountNumber
  FIList.Add FI, Key:=FIDAcctID
  '----------------------------------------------------------------------------------
  Set FI = New oFI
  FI.name = "Costco Visa"
  FI.AccountNumber = "8590"
  FI.FID = "2102"
  FI.DBCRdirection = -1
  getExistingFITransactions FI.name, FI.Transactions
  FI.BGColorIndex = 2
  FI.FGColorIndex = 33
  FIDAcctID = FI.FID & " " & FI.AccountNumber
  FIList.Add FI, Key:=FIDAcctID
  '----------------------------------------------------------------------------------
   Set FI = New oFI
  FI.name = "AAdvantage Mastcard"
  FI.AccountNumber = "8379"
  FI.FID = "2102"
  FI.DBCRdirection = -1
  getExistingFITransactions FI.name, FI.Transactions
  FI.BGColorIndex = 15
  FI.FGColorIndex = 3
  FIDAcctID = FI.FID & " " & FI.AccountNumber
  FIList.Add FI, Key:=FIDAcctID
  

  Set loadFinancialInstitutions = FIList
  
GoTo theEnd
errorHandleLoadFinancialInstitutions:
  displayError Err.Number, Err.Description, "Error: Source: Load Financial Information, FI= " & FI.name, FATALERR

theEnd:
End Function

Function getFIInfo(filestr As String) As String
'---------------------------------------------------------------------------------------
' Procedure : getFIInfo
' Author    : Christopher Prost, CP Business Analysis LLC. (9/21/2020)
' Website   : http://www.cpbusinessanalysis.com
' Copyright : 2020 CP Business Analysis LLC.  All Rights Reserved.
' Purpose   : put FIInfo into FIInfo array
'
' Usage:
' ------
' getFIInfo
'     input   : elements, contents of 1 QFX file
'     output  : name, name of FI
'     output  : "FI ACCTID", unique ID representing FI
'
' readTransactions
'---------------------------------------------------------------------------------------
      
  Dim FIKey As String
  
 On Error GoTo errorHandleGetFIInfo
 
 
  ' ModProc = "0201"
  
  FIKey = xmlfieldvalue(filestr, "<FID>", 1)
  FIKey = FIKey & " " & Right(xmlfieldvalue(filestr, "<ACCTID>", 1), 4)
   
  getFIInfo = FIKey
   
  GoTo theEnd

GoTo theEnd
errorHandleGetFIInfo:
  displayError Err.Number, Err.Description, "Error: Source: get FI Info, FIKey = " & FIKey & ", Length of filestr = " & Len(filestr), FATALERR
theEnd:
End Function
