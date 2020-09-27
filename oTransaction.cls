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

Public Property Get Source() As String
  Source = mSource
End Property

Public Property Get FITID() As String
  FITID = mFITID
End Property

Public Property Get Description() As String
  Description = mDescription
End Property

Public Property Get Category() As String
  Category = mCategory
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

Public Property Let Source(src As String)
  mSource = src
End Property

Public Property Let FITID(transID As String)
  mFITID = transID
End Property

Public Property Let Description(desc As String)
  mDescription = desc
End Property

Public Property Let Category(grouping As String)
  mCategory = grouping
End Property

Public Property Let amount(amt As Currency)
  mAmount = amt
End Property

Public Property Let Existing(doesItExist As Boolean)
  mExisting = doesItExist
End Property
