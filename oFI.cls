VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "oFI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mName As String
Private mFID As String
Private mAccountNumber As String
Private mDBCRdirection As Integer
Private mBalance As Currency
Private mTransactions As Collection
Private mBGColorIndex As Integer
Private mFGColorIndex As Integer

Private Sub Class_Initialize()
  Set mTransactions = New Collection
End Sub

Public Property Get name() As String
  name = mName
End Property

Public Property Get FID() As String
  FID = mFID
End Property

Public Property Get AccountNumber() As String
  AccountNumber = mAccountNumber
End Property

Public Property Get DBCRdirection() As Integer
  DBCRdirection = mDBCRdirection
End Property

Public Property Get Balance() As Currency
  Balance = mBalance
End Property
Public Property Get Transactions() As Collection
  Set Transactions = mTransactions
End Property
Public Property Get BGColorIndex() As Integer
  BGColorIndex = mBGColorIndex
End Property

Public Property Get FGColorIndex() As Integer
  FGColorIndex = mFGColorIndex
End Property

Public Property Let name(nam As String)
  mName = nam
End Property

Public Property Let FID(FI As String)
  mFID = FI
End Property

Public Property Let AccountNumber(acct As String)
  mAccountNumber = acct
End Property

Public Property Let DBCRdirection(direction As Integer)
  mDBCRdirection = direction
End Property

Public Property Let Balance(bal As Currency)
  mBalance = bal
End Property

Public Property Set Transactions(trans As oTransaction)
 Set mTransactions = trans
End Property

Public Property Let BGColorIndex(BGColor As Integer)
  mBGColorIndex = BGColor
End Property

Public Property Let FGColorIndex(FGColor As Integer)
  mFGColorIndex = FGColor
End Property
