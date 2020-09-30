Attribute VB_Name = "ModFileMgmt"
Option Explicit

Sub getFileList(types As String, fileList As Collection)
'---------------------------------------------------------------------------------------
' Procedure : readTransacations
' Author    : Christopher Prost, CP Business Analysis LLC. (9/21/2020)
' Website   : http://www.cpbusinessanalysis.com
' Copyright : 2020 CP Business Analysis LLC.  All Rights Reserved.
' Purpose   : locate the directory, create list of files
'
' Usage:
' ------
' readTransactions
'     input : all QFX files in provided downloads directory
'    output : elements, file contents in array of elements
'    output : file contents in one large string
'
' read transactions entry module
'---------------------------------------------------------------------------------------

  Dim modProcErr As String
  Dim fso As Object                               ' file system object
  Dim errMsg As String                            ' custom error message
  
  Dim oFolder As Object                           ' $Home/downloads directory
  Dim fileTypes() As String                       ' array of supported 3 letter file types
  Dim ftIndex As Integer
  Dim oFile  As Object                            ' Each file in oFolder
  Dim sourceFile As Object
  Dim fileStr As String
  Dim noEOLfileStr As String
  Dim downloadsPath As String
  Dim supportedFileTypes As String
  Dim fileListIndex As Integer
  
      
  Const ForReading = 1, ForWriting = 2, ForAppending = 8
  Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

  downloadsPath = Environ$("USERPROFILE") & "\Downloads"
  supportedFileTypes = "QFX"
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set fileList = New Collection
  Set oFolder = fso.GetFolder(downloadsPath)
  fileTypes = Split(supportedFileTypes, " ")
  fileListIndex = 0
  For Each oFile In oFolder.Files
    For ftIndex = LBound(fileTypes) To UBound(fileTypes)
      If (LCase(fso.GetExtensionName(oFile.Path)) = LCase(fileTypes(ftIndex))) Then
        Set sourceFile = fso.openTextfile(oFile.Path, ForReading)
        fileStr = sourceFile.ReadAll
        noEOLfileStr = Replace(Replace(Replace(fileStr, vbCrLf, ""), vbLf, ""), vbCr, "")
        fileListIndex = fileListIndex + 1
        Debug.Print fileListIndex & ". " & oFile.Path
        fileList.Add noEOLfileStr
        Set sourceFile = Nothing
      End If
    Next ftIndex
  Next
  Set oFolder = Nothing
  Set fso = Nothing
End Sub
