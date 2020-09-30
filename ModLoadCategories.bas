Attribute VB_Name = "ModLoadCategories"
Option Explicit

Public categories As Collection

'Categorization
'--------------
Private Const LOOKUPKEYWORDSCOL = 1
Private Const LOOKUPVALUECOL = 2

Public maxDescriptionCategoryWordCount As Integer


Sub getExistingCategoryDescriptions(str As String)
'---------------------------------------------------------------------------------------
' Procedure : getExistingCategoryDescriptions
' Author    : Christopher Prost, CP Business Analysis LLC. (9/21/2020)
' Website   : http://www.cpbusinessanalysis.com
' Copyright : 2020 CP Business Analysis LLC.  All Rights Reserved.
' Purpose   : Read Category data from lookup table in sheet 3 of this workbook
'
' Usage:
' ------
' getExistingCategoryDescriptions
'     input : This Workbook Sheet 3 Table of Categories
'    output : A collection of Description substrings and their associated categories
'
' Called From
' ------------
' main
'
' History
' ------
' Usually the first few words of a description identify a vendor
' Typically based on the vendor, one can identify the category, like Dominos means "fast food" and Kroger means groceries
' This work has alredy been done by the categorize function, This is the result
' It is stored in a data management system and read in a structure for faster processing
'---------------------------------------------------------------------------------------
  Dim category As oCategory
  Dim rw As Long
  Dim lastrow  As Long
  Dim lookupSheet As Worksheet
  Dim numKeyPhraseWords As Integer
  Dim strTmp As String
  
  
 On Error GoTo errorHandleGetExistingCategoryDescriptions
 
  
  Set lookupSheet = ThisWorkbook.Worksheets(3)
  lastrow = lookupSheet.Cells(Rows.Count, LOOKUPKEYWORDSCOL).End(xlUp).Row
  
  Set categories = New Collection
  
  For rw = 2 To lastrow
    Set category = New oCategory
    On Error Resume Next
    category.keyPhrase = lookupSheet.Cells(rw, LOOKUPKEYWORDSCOL).value
    category.value = lookupSheet.Cells(rw, LOOKUPVALUECOL).value
    category.Existing = True
    categories.Add category, lookupSheet.Cells(rw, LOOKUPKEYWORDSCOL).value
  Next rw
  maxDescriptionCategoryWordCount = lookupSheet.Cells(2, 4).value

GoTo theEnd
errorHandleGetExistingCategoryDescriptions:
  displayError Err.Number, Err.Description, "Error: Source:get Existing Category Descriptions,  Row = " & rw & ", Lastrow = " & lastrow, FATALERR

theEnd:
End Sub
