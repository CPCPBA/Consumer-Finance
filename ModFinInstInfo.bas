Attribute VB_Name = "ModFinInstInfo"
Option Explicit


Function loadFinancialInstitutions(loadTransactions As Boolean) As Collection

On Error GoTo errorHandleLoadFinancialInstitutions
   
  Dim FIDAcctID As String                                  ' Quicken Financial Institute ID +  FI's last 4 Acct ID
  Dim fi As oFI                                            ' Financial Institute Object
  Dim FIList As Collection                                 ' Collection of FI Objects
   '----------------------------------------------------------------------------------
  Set FIList = New Collection
  
  Set fi = New oFI
  fi.name = "Bank of America"
  fi.AccountNumber = "9252"
  fi.FID = "5959"
  fi.DBCRdirection = 1
  If loadTransactions Then
    getExistingFITransactions fi.name, fi.Transactions
  End If
  fi.BGColorIndex = 24
  fi.FGColorIndex = 3
  FIDAcctID = fi.FID & " " & fi.AccountNumber
  FIList.Add fi, Key:=FIDAcctID
  '----------------------------------------------------------------------------------
  Set fi = New oFI
  fi.name = "AMEX"
  fi.AccountNumber = "3006"
  fi.FID = "3101"
  fi.DBCRdirection = -1
  If loadTransactions Then
    getExistingFITransactions fi.name, fi.Transactions
  End If
  fi.BGColorIndex = 35
  fi.FGColorIndex = 47
  FIDAcctID = fi.FID & " " & fi.AccountNumber
  FIList.Add fi, Key:=FIDAcctID
  '----------------------------------------------------------------------------------
  Set fi = New oFI
  fi.name = "Costco Visa"
  fi.AccountNumber = "8590"
  fi.FID = "2102"
  fi.DBCRdirection = -1
  If loadTransactions Then
    getExistingFITransactions fi.name, fi.Transactions
  End If
  fi.BGColorIndex = 2
  fi.FGColorIndex = 33
  FIDAcctID = fi.FID & " " & fi.AccountNumber
  FIList.Add fi, Key:=FIDAcctID
  '----------------------------------------------------------------------------------
   Set fi = New oFI
  fi.name = "AAdvantage Mastcard"
  fi.AccountNumber = "8379"
  fi.FID = "2102"
  fi.DBCRdirection = -1
  If loadTransactions Then
    getExistingFITransactions fi.name, fi.Transactions
  End If
  fi.BGColorIndex = 15
  fi.FGColorIndex = 3
  FIDAcctID = fi.FID & " " & fi.AccountNumber
  FIList.Add fi, Key:=FIDAcctID
  

  Set loadFinancialInstitutions = FIList
  
GoTo theEnd
errorHandleLoadFinancialInstitutions:
  displayError Err.Number, Err.Description, "Error: Source: Load Financial Information, FI= " & fi.name, FATALERR

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
