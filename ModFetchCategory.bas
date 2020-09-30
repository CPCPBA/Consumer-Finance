Attribute VB_Name = "ModFetchCategory"
Option Explicit





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
  Dim lookupSheet As Worksheet
  
 
  On Error GoTo errorHandleFindCategory
  
  findCategory = "N/F"
  categoryNotFound = True
  Set lookupSheet = ThisWorkbook.Sheets(3)
  
 
  str = Replace(Replace(Replace(Replace(Replace(str, "  ", " "), "*", " "), "-", " "), "_", " "), "  ", " ")
  strArray = Split(str, " ")
  maxDescriptionCategoryWordCount = lookupSheet.Cells(2, 4)
  strWordCount = min(UBound(strArray) + 1, maxDescriptionCategoryWordCount)
  
  While (strWordCount >= 1) And categoryNotFound
    PossibleValue = CategoryLookup(str)
    If PossibleValue <> "N/F" Then
      categoryNotFound = False
    Else
      If strWordCount > 1 Then
        ReDim Preserve strArray(UBound(strArray) - 1)
        str = Join(strArray)
        strWordCount = strCount(str, " ") + 1
      Else
        strWordCount = 0
      End If
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
