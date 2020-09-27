VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "oBank"
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
Private mTransactionList As Collection

Private Sub Class_Initialize()
  Set mTransactionList = New Collection
End Sub

Public Property Get Name() As String
  Name = mName
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
Public Property Get TransactionList() As Collection
  Set TransactionList = mTransactionList
End Property


Public Property Let Name(nam As String)
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

Public Property Set TransactionList(trans As oTransaction)
  mTransactionList.Add trans
End Property


