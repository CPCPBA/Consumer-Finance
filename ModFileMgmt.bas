Attribute VB_Name = "ModFileMgmt"
Option Explicit

Function getFileList(types As String) As Collection
'---------------------------------------------------------------------------------------
' Procedure : getFileList
' Author    : Christopher Prost, CP Business Analysis LLC. (9/21/2020)
' Website   : http://www.cpbusinessanalysis.com
' Copyright : 2020 CP Business Analysis LLC.  All Rights Reserved.
' Purpose   : locate the directory, create collection of file content
'
' Usage:
' ------
' getFileList
'     input : all QFX files in provided downloads directory
'    output : collection of file contents
'
' Main
'---------------------------------------------------------------------------------------

  Const ForReading = 1, ForWriting = 2, ForAppending = 8
  Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

  Dim downloadsPath As String
  Dim filetypes() As String
  Dim ftIndex As Integer
  Dim filestr As String
  Dim noEOLfileStr As String
  Dim Filelist As Collection
  Dim fileListIndex As Integer
  
  Dim fso As Object                               ' file system object
  Dim oFolder As Object                           ' $Home/downloads directory
  Dim oFile  As Object                            ' Each file in oFolder
  Dim sourceFile As Object
  Dim path As String
  
  
 On Error GoTo errorHandlegetFileList
 
      
  downloadsPath = Environ$("USERPROFILE") & "\Downloads"
  path = downloadsPath & "and I don't know yet"
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set Filelist = New Collection
  Set oFolder = fso.GetFolder(downloadsPath)
  filetypes = Split(types, " ")
  fileListIndex = 0
  For Each oFile In oFolder.Files
    For ftIndex = LBound(filetypes) To UBound(filetypes)
      path = oFile.path
      If (LCase(fso.GetExtensionName(oFile.path)) = LCase(filetypes(ftIndex))) Then
        Set sourceFile = fso.openTextfile(oFile.path, ForReading)
        filestr = sourceFile.ReadAll
        noEOLfileStr = Replace(Replace(Replace(filestr, vbCrLf, ""), vbLf, ""), vbCr, "")
        fileListIndex = fileListIndex + 1
        Debug.Print fileListIndex & ". " & oFile.path
        Filelist.Add noEOLfileStr
        Set sourceFile = Nothing
      End If
    Next ftIndex
  Next
  Set getFileList = Filelist
  Set oFolder = Nothing
  Set fso = Nothing


GoTo theEnd
errorHandlegetFileList:
  displayError Err.Number, Err.Description, "Error: Source: get File List, path = " & oFile.path, FATALERR

theEnd:
End Function

