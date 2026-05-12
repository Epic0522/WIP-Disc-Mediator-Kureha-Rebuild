VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TrackEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mTrackNo As Long
Private mFilePath As String
Private mSource As String
Private mPregap As String
Private mDuration As String
Private mPostgap As String
Private mFlags As String
Private mPreviewKiB As Long
Private mEnglish As TOCCDText
Private mJapanese As TOCCDText

Private Sub Class_Initialize()
    Set mEnglish = New TOCCDText
    Set mJapanese = New TOCCDText

    mEnglish.LanguageName = "English"
    mJapanese.LanguageName = "Japanese"

    mSource = "Audio"
    mPregap = "00:02:00"
    mDuration = "04:00:00"
    mPostgap = "00:00:00"
    mFlags = "DGP"
    mPreviewKiB = 0
End Sub

Public Property Get TrackNo() As Long
    TrackNo = mTrackNo
End Property

Public Property Let TrackNo(ByVal value As Long)
    mTrackNo = value
End Property

Public Property Get FilePath() As String
    FilePath = mFilePath
End Property

Public Property Let FilePath(ByVal value As String)
    mFilePath = value
End Property

Public Property Get Source() As String
    Source = mSource
End Property

Public Property Let Source(ByVal value As String)
    mSource = value
End Property

Public Property Get Pregap() As String
    Pregap = mPregap
End Property

Public Property Let Pregap(ByVal value As String)
    mPregap = value
End Property

Public Property Get Duration() As String
    Duration = mDuration
End Property

Public Property Let Duration(ByVal value As String)
    mDuration = value
End Property

Public Property Get Postgap() As String
    Postgap = mPostgap
End Property

Public Property Let Postgap(ByVal value As String)
    mPostgap = value
End Property

Public Property Get Flags() As String
    Flags = mFlags
End Property

Public Property Let Flags(ByVal value As String)
    mFlags = value
End Property

Public Property Get PreviewKiB() As Long
    PreviewKiB = mPreviewKiB
End Property

Public Property Let PreviewKiB(ByVal value As Long)
    mPreviewKiB = value
End Property

Public Property Get EnglishText() As TOCCDText
    Set EnglishText = mEnglish
End Property

Public Property Set EnglishText(ByVal value As TOCCDText)
    Set mEnglish = value
End Property

Public Property Get JapaneseText() As TOCCDText
    Set JapaneseText = mJapanese
End Property

Public Property Set JapaneseText(ByVal value As TOCCDText)
    Set mJapanese = value
End Property

Public Property Get DisplayTitle() As String
    If mEnglish.LanguageEnabled And mEnglish.TitleEnabled And Trim$(mEnglish.Title) <> "" Then
        DisplayTitle = mEnglish.Title
    ElseIf mJapanese.LanguageEnabled And mJapanese.TitleEnabled And Trim$(mJapanese.Title) <> "" Then
        DisplayTitle = mJapanese.Title
    Else
        DisplayTitle = "(untitled)"
    End If
End Property

Public Property Get DisplayPerformer() As String
    If mEnglish.LanguageEnabled And mEnglish.PerformerEnabled And Trim$(mEnglish.Performer) <> "" Then
        DisplayPerformer = mEnglish.Performer
    ElseIf mJapanese.LanguageEnabled And mJapanese.PerformerEnabled And Trim$(mJapanese.Performer) <> "" Then
        DisplayPerformer = mJapanese.Performer
    Else
        DisplayPerformer = ""
    End If
End Property

