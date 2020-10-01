Attribute VB_Name = "ModAssignCategory"
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
  Dim regEx As Object
  Dim counter As Integer
  Dim category As oCategory
  
  On Error GoTo errorHandleFindCategory
  Set regEx = CreateObject("VBScript.RegExp")
  
  findCategory = "N/F"
  
  counter = 0
  For Each category In categories
    counter = counter + 1
    With regEx
      .Pattern = category.RegExPhrase
      .IgnoreCase = True
      PossibleValue = .Test(str)
      Debug.Print "Looking for " & .Pattern & " in " & str & ".  " & PossibleValue
      If PossibleValue Then
        findCategory = category.value
        Exit For
      End If
    End With
  Next

GoTo theEnd
errorHandleFindCategory:

  displayError Err.Number, Err.Description, "Error: Source:Find Category,  str = " & str, FATALERR

theEnd:

End Function

