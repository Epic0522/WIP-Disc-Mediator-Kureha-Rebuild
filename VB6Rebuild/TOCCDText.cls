VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TOCCDText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mLanguageName As String
Private mLanguageEnabled As Boolean
Private mTitleEnabled As Boolean
Private mPerformerEnabled As Boolean
Private mSongwriterEnabled As Boolean
Private mComposerEnabled As Boolean
Private mArrangerEnabled As Boolean
Private mMessageEnabled As Boolean
Private mTitle As String
Private mPerformer As String
Private mSongwriter As String
Private mComposer As String
Private mArranger As String
Private mMessage As String

Public Property Get LanguageName() As String
    LanguageName = mLanguageName
End Property

Public Property Let LanguageName(ByVal value As String)
    mLanguageName = value
End Property

Public Property Get LanguageEnabled() As Boolean
    LanguageEnabled = mLanguageEnabled
End Property

Public Property Let LanguageEnabled(ByVal value As Boolean)
    mLanguageEnabled = value
End Property

Public Property Get TitleEnabled() As Boolean
    TitleEnabled = mTitleEnabled
End Property

Public Property Let TitleEnabled(ByVal value As Boolean)
    mTitleEnabled = value
End Property

Public Property Get PerformerEnabled() As Boolean
    PerformerEnabled = mPerformerEnabled
End Property

Public Property Let PerformerEnabled(ByVal value As Boolean)
    mPerformerEnabled = value
End Property

Public Property Get SongwriterEnabled() As Boolean
    SongwriterEnabled = mSongwriterEnabled
End Property

Public Property Let SongwriterEnabled(ByVal value As Boolean)
    mSongwriterEnabled = value
End Property

Public Property Get ComposerEnabled() As Boolean
    ComposerEnabled = mComposerEnabled
End Property

Public Property Let ComposerEnabled(ByVal value As Boolean)
    mComposerEnabled = value
End Property

Public Property Get ArrangerEnabled() As Boolean
    ArrangerEnabled = mArrangerEnabled
End Property

Public Property Let ArrangerEnabled(ByVal value As Boolean)
    mArrangerEnabled = value
End Property

Public Property Get MessageEnabled() As Boolean
    MessageEnabled = mMessageEnabled
End Property

Public Property Let MessageEnabled(ByVal value As Boolean)
    mMessageEnabled = value
End Property

Public Property Get Title() As String
    Title = mTitle
End Property

Public Property Let Title(ByVal value As String)
    mTitle = value
End Property

Public Property Get Performer() As String
    Performer = mPerformer
End Property

Public Property Let Performer(ByVal value As String)
    mPerformer = value
End Property

Public Property Get Songwriter() As String
    Songwriter = mSongwriter
End Property

Public Property Let Songwriter(ByVal value As String)
    mSongwriter = value
End Property

Public Property Get Composer() As String
    Composer = mComposer
End Property

Public Property Let Composer(ByVal value As String)
    mComposer = value
End Property

Public Property Get Arranger() As String
    Arranger = mArranger
End Property

Public Property Let Arranger(ByVal value As String)
    mArranger = value
End Property

Public Property Get Message() As String
    Message = mMessage
End Property

Public Property Let Message(ByVal value As String)
    mMessage = value
End Property

Public Function EnabledFieldCount() As Long
    If Not mLanguageEnabled Then Exit Function
    If mTitleEnabled Then EnabledFieldCount = EnabledFieldCount + 1
    If mPerformerEnabled Then EnabledFieldCount = EnabledFieldCount + 1
    If mSongwriterEnabled Then EnabledFieldCount = EnabledFieldCount + 1
    If mComposerEnabled Then EnabledFieldCount = EnabledFieldCount + 1
    If mArrangerEnabled Then EnabledFieldCount = EnabledFieldCount + 1
    If mMessageEnabled Then EnabledFieldCount = EnabledFieldCount + 1
End Function

Public Function Summary() As String
    If Not mLanguageEnabled Then
        Summary = mLanguageName & ": <off>"
        Exit Function
    End If

    Summary = mLanguageName & ": " & _
        "Title=" & IIf(mTitleEnabled, QuoteText(mTitle), "<off>") & ", " & _
        "Performer=" & IIf(mPerformerEnabled, QuoteText(mPerformer), "<off>") & ", " & _
        "Songwriter=" & IIf(mSongwriterEnabled, QuoteText(mSongwriter), "<off>") & ", " & _
        "Composer=" & IIf(mComposerEnabled, QuoteText(mComposer), "<off>") & ", " & _
        "Arranger=" & IIf(mArrangerEnabled, QuoteText(mArranger), "<off>") & ", " & _
        "Message=" & IIf(mMessageEnabled, QuoteText(mMessage), "<off>")
End Function

Private Function QuoteText(ByVal value As String) As String
    QuoteText = """" & value & """"
End Function

