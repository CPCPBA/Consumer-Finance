Attribute VB_Name = "ModCategory"
Option Explicit

Private categories As Collection

'Categorization
'--------------
Const LOOKUPKEYWORDSCOL = 1
Const LOOKUPVALUECOL = 2
Private maxDescriptionCategoryWordCount As Integer

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
    If lookupSheet.Cells(rw, LOOKUPKEYWORDSCOL).value Is Empty Then     ' if empty row assume we are done
      Exit For
    End If
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

Function findCategory(str As String) As String
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
'     input : largestKeyPhraseCount stored globally as an integer
'    output : transaction category
'
' Called From:
' ------------
' GetActualTransactions
'---------------------------------------------------------------------------------------
  
  Dim PossibleValue As String
  Dim categoryNotFound
  Dim strArray() As String
  Dim strWordCount As Integer
 
  On Error GoTo errorHandleFindCategory
  
  findCategory = "N/F"
  categoryNotFound = True
 
  str = Replace(Replace(Replace(Replace(str, "  ", " "), "*", " "), "-", " "), "_", " ")
  strArray = Split(str, " ")
  strWordCount = min(UBound(strArray), maxDescriptionCategoryWordCount)
  
  While (strWordCount > 0) And categoryNotFound
    PossibleValue = CategoryLookup(str)
    If PossibleValue <> "N/F" Then
      categoryNotFound = False
    Else
      ReDim Preserve strArray(UBound(strArray) - 1)
      str = Join(strArray)
      strWordCount = strCount(str, " ") + 1
    End If
  Wend '
  
  If categoryNotFound Then
    findCategory = "N/F"
  Else
    findCategory = PossibleValue
  End If


GoTo theEnd
errorHandleFindCategory:

  displayError Err.Number, Err.Description, "Error: Source:Find Category,  str = " & str, FATALERR

theEnd:

End Function


Function CategoryLookup(str As String) As String
'---------------------------------------------------------------------------------------
' Procedure : autoCategorize
' Author    : Christopher Prost, CP Business Analysis LLC. (7/9/2020)
' Website   : http://www.cpbusinessanalysis.com
' Copyright : 2020 CP Business Analysis LLC.  All Rights Reserved.
' Purpose   : Looks up existing keyword and populates category field
'
' Usage:
' ------
' autoCategorize
'     input : str,subset of transaction description
'    output : transaction category
'
' Called From:
' ------------
' Find Category
'---------------------------------------------------------------------------------------

  Dim category As oCategory

  On Error Resume Next
  Set category = categories(str)
  If Not category Is Nothing Then
    CategoryLookup = category.value
  Else
    CategoryLookup = "N/F"
  End If

theEnd:

End Function
