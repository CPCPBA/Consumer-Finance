VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "oTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mIndex As Long
Private mPostedDate As Date
Private mSource As String
Private mFITID As String
Private mDescription As String
Private mCategory As String
Private mAmount As Currency
Private mExisting As Boolean

Public Property Get postedDate() As Date
  postedDate = mPostedDate
End Property

Public Property Get Index() As Long
  Index = mIndex
End Property

Public Property Get Source() As String
  Source = mSource
End Property

Public Property Get FITID() As String
  FITID = mFITID
End Property

Public Property Get Description() As String
  Description = mDescription
End Property

Public Property Get category() As String
  category = mCategory
End Property

Public Property Get amount() As Currency
  amount = mAmount
End Property

Public Property Get Existing() As Boolean
  Existing = mExisting
End Property

Public Property Let postedDate(posted As Date)
  mPostedDate = posted
End Property
Public Property Let Index(i As Long)
  mIndex = i
End Property

Public Property Let Source(src As String)
  mSource = src
End Property

Public Property Let FITID(TransID As String)
  mFITID = TransID
End Property

Public Property Let Description(desc As String)
  mDescription = desc
End Property

Public Property Let category(grouping As String)
  mCategory = grouping
End Property

Public Property Let amount(amt As Currency)
  mAmount = amt
End Property

Public Property Let Existing(doesItExist As Boolean)
  mExisting = doesItExist
End Property
