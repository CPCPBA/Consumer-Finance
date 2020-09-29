VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "oCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mKeyPhrase As String
Private mValue As String
Private mExisting As Boolean

Public Property Get keyPhrase() As String
  keyPhrase = mKeyPhrase
End Property

Public Property Get value() As String
  value = mValue
End Property

Public Property Let keyPhrase(phrase As String)
  mKeyPhrase = phrase
End Property

Public Property Let value(val As String)
  mValue = val
End Property

Public Property Get Existing() As Boolean
  Existing = mExisting
End Property

Public Property Let Existing(doesItExist As Boolean)
  mExisting = doesItExist
End Property

