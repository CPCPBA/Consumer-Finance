Attribute VB_Name = "ModCategories"
Option Explicit

Public categories As Collection

'Categorization
'--------------
Private Const LOOKUPKEYWORDSCOL = 1
Private Const LOOKUPVALUECOL = 2



Sub getExistingCategoryDescriptions()
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
  Dim nummRegExPhraseWords As Integer
  Dim strTmp As String
  
  
 On Error GoTo errorHandleGetExistingCategoryDescriptions
 
  
  Set lookupSheet = ThisWorkbook.Worksheets(4)
  lastrow = lookupSheet.Cells(Rows.Count, LOOKUPKEYWORDSCOL).End(xlUp).Row
  
  Set categories = New Collection
  
  For rw = 2 To lastrow
    Set category = New oCategory
    On Error Resume Next
    category.RegExPhrase = lookupSheet.Cells(rw, LOOKUPKEYWORDSCOL).value
    category.value = lookupSheet.Cells(rw, LOOKUPVALUECOL).value
    category.Existing = True
    categories.Add category, lookupSheet.Cells(rw, LOOKUPKEYWORDSCOL).value
  Next rw

GoTo theEnd
errorHandleGetExistingCategoryDescriptions:
  displayError Err.Number, Err.Description, "Error: Source:get Existing Category Descriptions,  Row = " & rw & ", Lastrow = " & lastrow, FATALERR

theEnd:
End Sub




Function findCategory(descr As String) As String
'---------------------------------------------------------------------------------------
' Procedure : findCategory
' Author    : Christopher Prost, CP Business Analysis LLC. (7/9/2020)
' Website   : http://www.cpbusinessanalysis.com
' Copyright : 2020 CP Business Analysis LLC.  All Rights Reserved.
' Purpose   : key phrases are subsets of transaction description.  find key phrase in table assign category to transaction
'
' Usage:
' ------
' findCategory
'     input : str, the transaction description
'     input : table of key phrases and categories, loaded globally as a collection
'     input : largestmRegExPhraseCount stored globally as an integer
'    output : transaction category
'
' Called From:
' ------------
' GetActualTransactions
'---------------------------------------------------------------------------------------
  
  Dim categoryNotFound
  Dim strWordCount As Integer
  Dim lookupSheet As Worksheet
  Dim regEx As Object
  Dim catNum As Integer
  Dim category As oCategory
  
  
 
  On Error GoTo errorHandleFindCategory
  
  categoryNotFound = True
 
  descr = Replace(Replace(Replace(Replace(Replace(descr, "  ", " "), "*", " "), "-", " "), "_", " "), "  ", " ")
  Set regEx = CreateObject("VBScript.RegExp")
  
  findCategory = "N/F"
  catNum = 1
  While (catNum <= categories.Count) And categoryNotFound
    Set category = categories.Item(catNum)
    regEx.Pattern = category.RegExPhrase
    regEx.IgnoreCase = True
    If regEx.Test(descr) Then
      findCategory = category.value
      categoryNotFound = False
    End If
    catNum = catNum + 1
  Wend
  
GoTo theEnd

errorHandleFindCategory:

  displayError Err.Number, Err.Description, "Error: Source:Find Category,  descr = " & descr, FATALERR

theEnd:

End Function


