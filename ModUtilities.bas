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
' getFIInfo
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

  On Error Resume Next

  ' ModProc = "9902"
  If criticality = FATALERR Then
    status = vbCritical
    prefix = "Cannot Continue: "
  ElseIf criticality = WARNERR Then
    status = vbQuestion
    prefix = "Warning: "
  Else
    status = vbInformation
    prefix = "FYI: "
  End If

  MsgBox prefix & " " & customMsg, status
  Debug.Print prefix & " " & customMsg & " : " & errNum & " : "; Description

theEnd:
End Sub

Function xmlfieldvalue(str As String, searchStr As String, startPos As Long) As String
'---------------------------------------------------------------------------------------
' Procedure : XMLfieldvalue
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
' anywhere
'---------------------------------------------------------------------------------------
  Dim valueStart As Long
  Dim bracketPos As Integer
  Dim found As Boolean
  
  On Error GoTo errorHandleXMLieldvalue
  
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
  


GoTo theEnd
errorHandleXMLieldvalue:
  displayError Err.Number, Err.Description, "Error: Source: XML field value, len of str = " & Len(str) & ", start = " & startPos & ", searchStr = " & searchStr, FATALERR

theEnd:
End Function

'*****************************************************************
'*****************************************************************
'
' function to compare 2 tokens return max or min
'
'*****************************************************************
'*****************************************************************
Public Function max(x, y As Variant) As Variant
  
 On Error GoTo errorHandleMax
 
  max = IIf(x > y, x, y)


GoTo theEnd
errorHandleMax:

  displayError Err.Number, Err.Description, "Error: Source: Max, X = " & x & ", Y = " & y, FATALERR

theEnd:
End Function
Public Function min(x, y As Variant) As Variant
  
 On Error GoTo errorHandleMin
 
   min = IIf(x < y, x, y)


GoTo theEnd
errorHandleMin:

  displayError Err.Number, Err.Description, "Error: Source: Min, X = " & x & ", Y = " & y, FATALERR

theEnd:
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
' Called From
' -----------
' anywhere
'---------------------------------------------------------------------------------------
 On Error GoTo errorHandleStrCount
 

  strCount = UBound(Split(str, searchStr))

GoTo theEnd
errorHandleStrCount:
  displayError Err.Number, Err.Description, "Error: Source: Str Count, len of str " & Len(str), FATALERR

theEnd:
End Function



