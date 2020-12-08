VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "myFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mFileName As String
Private mFileContents As String

Public Property Get filename() As String
  filename = mFileName
End Property

Public Property Get fileContents() As String
  fileContents = mFileContents
End Property

Public Property Let filename(name As String)
  mFileName = name
End Property

Public Property Let fileContents(contents As String)
  mFileContents = contents
End Property

