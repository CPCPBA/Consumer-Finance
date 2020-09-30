Attribute VB_Name = "ModUtilities"
Option Explicit

Sub displayError(errNum As Integer, Description As String, customMsg As String, criticality As errCriticality)
'---------------------------------------------------------------------------------------
' Procedure : displayError
' Author    : Christopher Prost, CP Business Analysis LLC. (9/21/2020)
' Website   : http://www.cpbusinessanalysis.com
' Copyright : 2020 CP Business Analysis LLC.  All Rights Reserved.  All Rights Reserved.
' Purpose   : standard error message display for all error
'
' Usage:
' ------
' getBankInfo
'     input : errNum, typically err.number
'     input : description, typically err.description
'     input : customMsg, a user friendly message more specific to what caused incident locally
'     input : criticality, incident priority
'    output : debug.print of system error
'    output : msgbox of user friendly custom error message if caught
'
' Called From:
' ------------
' should be anywhere
'---------------------------------------------------------------------------------------
  Dim modProcErr As String
  Dim status As Integer
  Dim prefix As String

#If Not DEBUGSTATUS Then
  On Error GoTo ErrorHandle
#End If

  ' ModProc = "9902"
  If criticality = FTLERR Then
    status = vbCritical
    prefix = "Cannot Continue: "
  ElseIf criticality = WRNERR Then
    status = vbQuestion
    prefix = "Warning: "
  Else
    status = vbInformation
    prefix = "FYI: "
  End If

<<<<<<< Updated upstream
  MsgBox prefix & " " & errMsg, status
=======
  ' MsgBox prefix & " " & customMsg, status
>>>>>>> Stashed changes
  Debug.Print prefix & " " & customMsg & " : " & errNum & " : "; Description

GoTo theEnd
ErrorHandle:
  displayError Err.Number, Err.Description, "There was a system error. Contact User Support", FTLERR

theEnd:
End Sub


Function xmlfieldvalue(str As String, searchStr As String, startPos As Long) As String
'---------------------------------------------------------------------------------------
' Procedure : xmlfieldvalue
' Author    : Christopher Prost, CP Business Analysis LLC. (9/21/2020)
' Website   : http://www.cpbusinessanalysis.com
' Copyright : 2020 CP Business Analysis LLC.  All Rights Reserved.
' Purpose   : return substr between two tags.  There is aassumption instr was
'
' Usage:
' ------
' strCount
'     input : str, string to be searched
'     input : searchStr, string to locate within str
'     input : desired location of substr
'    output : strCount, number of times searchStr appears within str
'
' parseBankInfo, parseTransInfo
'---------------------------------------------------------------------------------------
  Dim valueStart As Long
  Dim bracketPos As Integer
  Dim found As Boolean
  
  found = False
  If InStr(startPos, str, searchStr) > 0 Then
    valueStart = InStr(startPos, str, searchStr) + Len(searchStr)
    bracketPos = InStr(valueStart, str, "<")
    If bracketPos > 0 Then
      xmlfieldvalue = Mid(str, valueStart, bracketPos - valueStart)
      found = True
    End If
  End If
  If Not found Then
    xmlfieldvalue = "N/D"
  End If
  
End Function

'*****************************************************************
'*****************************************************************
'
' Method to count substrings within a string
'
'*****************************************************************
'*****************************************************************
'
Function strCount(str As String, searchStr As String) As Integer
'---------------------------------------------------------------------------------------
' Procedure : strCount
' Author    : Christopher Prost, CP Business Analysis LLC. (9/21/2020)
' Website   : http://www.cpbusinessanalysis.com
' Copyright : 2020 CP Business Analysis LLC.  All Rights Reserved.
' Purpose   : identify number of times searchStr appears in
'
' Usage:
' ------
' strCount
'     input : str, string to be searched
'     input : searchStr, string to locate within str
'    output : strCount, number of times searchStr appears within str
'
' parseBankInfo, parseTransInfo
'---------------------------------------------------------------------------------------
  Dim subStrs() As String

  subStrs = Split(str, searchStr)
  strCount = UBound(subStrs) - LBound(subStrs)

GoTo theEnd
ErrorHandle:
  displayError Err.Number, Err.Description, "There was a system error. Contact User Support", FTLERR

theEnd:
End Function
'*****************************************************************
'*****************************************************************
'
' Example of adding a collection to a class
'
'*****************************************************************
'*****************************************************************
'
'' class cProduct
'Private pChildList As Collection
'
'Private Sub Class_Initialize()
'    Set pChildList = New Collection
'End Sub
'
'Public Property Set ChildList(Value As CProduct)
'    pChildList.Add Value
'End Property
'
'Public Property Get ChildList() As Collection
'    ChildList = pChildList
'End Property
'
'
'' The main function calling
'
'Set Pro = New CProduct
'Set Child = New CProduct
'Pro.ChildList.Add Child
