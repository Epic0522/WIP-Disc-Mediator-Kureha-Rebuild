VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "FileEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mName As String
Private mOriginalPath As String
Private mSizeText As String
Private mModifiedText As String
Private mIsDirectory As Boolean

Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal value As String)
    mName = value
End Property

Public Property Get OriginalPath() As String
    OriginalPath = mOriginalPath
End Property

Public Property Let OriginalPath(ByVal value As String)
    mOriginalPath = value
End Property

Public Property Get SizeText() As String
    SizeText = mSizeText
End Property

Public Property Let SizeText(ByVal value As String)
    mSizeText = value
End Property

Public Property Get ModifiedText() As String
    ModifiedText = mModifiedText
End Property

Public Property Let ModifiedText(ByVal value As String)
    mModifiedText = value
End Property

Public Property Get IsDirectory() As Boolean
    IsDirectory = mIsDirectory
End Property

Public Property Let IsDirectory(ByVal value As Boolean)
    mIsDirectory = value
End Property

