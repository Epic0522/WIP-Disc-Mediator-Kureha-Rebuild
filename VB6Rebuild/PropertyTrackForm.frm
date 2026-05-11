VERSION 5.00
Begin VB.Form PropertyTrackForm
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Track CD-TEXT"
   ClientHeight    =   6525
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8775
   LinkTopic       =   "PropertyTrackForm"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkCDTextEnabled
      Caption         =   "Enable CD-TEXT for this track"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   2835
   End
   Begin VB.Frame fraEnglish
      Caption         =   "English"
      Height          =   2715
      Left            =   180
      TabIndex        =   1
      Top             =   600
      Width           =   3975
      Begin VB.CheckBox chkEnglishEnabled
         Caption         =   "Use this language"
         Height          =   255
         Left            =   180
         TabIndex        =   2
         Top             =   240
         Width           =   1635
      End
      Begin VB.CheckBox chkEnglishTitle
         Caption         =   "Title"
         Height          =   255
         Left            =   180
         TabIndex        =   3
         Top             =   600
         Width           =   1155
      End
      Begin VB.TextBox txtEnglishTitle
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   540
         Width           =   2295
      End
      Begin VB.CheckBox chkEnglishPerformer
         Caption         =   "Performer"
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   960
         Width           =   1155
      End
      Begin VB.TextBox txtEnglishPerformer
         Height          =   315
         Left            =   1440
         TabIndex        =   6
         Top             =   900
         Width           =   2295
      End
      Begin VB.CheckBox chkEnglishSongwriter
         Caption         =   "Songwriter"
         Height          =   255
         Left            =   180
         TabIndex        =   7
         Top             =   1320
         Width           =   1155
      End
      Begin VB.TextBox txtEnglishSongwriter
         Height          =   315
         Left            =   1440
         TabIndex        =   8
         Top             =   1260
         Width           =   2295
      End
      Begin VB.CheckBox chkEnglishComposer
         Caption         =   "Composer"
         Height          =   255
         Left            =   180
         TabIndex        =   9
         Top             =   1680
         Width           =   1155
      End
      Begin VB.TextBox txtEnglishComposer
         Height          =   315
         Left            =   1440
         TabIndex        =   10
         Top             =   1620
         Width           =   2295
      End
      Begin VB.CheckBox chkEnglishArranger
         Caption         =   "Arranger"
         Height          =   255
         Left            =   180
         TabIndex        =   11
         Top             =   2040
         Width           =   1155
      End
      Begin VB.TextBox txtEnglishArranger
         Height          =   315
         Left            =   1440
         TabIndex        =   12
         Top             =   1980
         Width           =   2295
      End
      Begin VB.CheckBox chkEnglishMessage
         Caption         =   "Message"
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   2400
         Width           =   1155
      End
      Begin VB.TextBox txtEnglishMessage
         Height          =   315
         Left            =   1440
         TabIndex        =   14
         Top             =   2340
         Width           =   2295
      End
   End
   Begin VB.Frame fraJapanese
      Caption         =   "Japanese"
      Height          =   2715
      Left            =   4440
      TabIndex        =   15
      Top             =   600
      Width           =   4155
      Begin VB.CheckBox chkJapaneseEnabled
         Caption         =   "Use this language"
         Height          =   255
         Left            =   180
         TabIndex        =   16
         Top             =   240
         Width           =   1635
      End
      Begin VB.CheckBox chkJapaneseTitle
         Caption         =   "Title"
         Height          =   255
         Left            =   180
         TabIndex        =   17
         Top             =   600
         Width           =   1155
      End
      Begin VB.TextBox txtJapaneseTitle
         Height          =   315
         Left            =   1500
         TabIndex        =   18
         Top             =   540
         Width           =   2355
      End
      Begin VB.CheckBox chkJapanesePerformer
         Caption         =   "Performer"
         Height          =   255
         Left            =   180
         TabIndex        =   19
         Top             =   960
         Width           =   1155
      End
      Begin VB.TextBox txtJapanesePerformer
         Height          =   315
         Left            =   1500
         TabIndex        =   20
         Top             =   900
         Width           =   2355
      End
      Begin VB.CheckBox chkJapaneseSongwriter
         Caption         =   "Songwriter"
         Height          =   255
         Left            =   180
         TabIndex        =   21
         Top             =   1320
         Width           =   1155
      End
      Begin VB.TextBox txtJapaneseSongwriter
         Height          =   315
         Left            =   1500
         TabIndex        =   22
         Top             =   1260
         Width           =   2355
      End
      Begin VB.CheckBox chkJapaneseComposer
         Caption         =   "Composer"
         Height          =   255
         Left            =   180
         TabIndex        =   23
         Top             =   1680
         Width           =   1155
      End
      Begin VB.TextBox txtJapaneseComposer
         Height          =   315
         Left            =   1500
         TabIndex        =   24
         Top             =   1620
         Width           =   2355
      End
      Begin VB.CheckBox chkJapaneseArranger
         Caption         =   "Arranger"
         Height          =   255
         Left            =   180
         TabIndex        =   25
         Top             =   2040
         Width           =   1155
      End
      Begin VB.TextBox txtJapaneseArranger
         Height          =   315
         Left            =   1500
         TabIndex        =   26
         Top             =   1980
         Width           =   2355
      End
      Begin VB.CheckBox chkJapaneseMessage
         Caption         =   "Message"
         Height          =   255
         Left            =   180
         TabIndex        =   27
         Top             =   2400
         Width           =   1155
      End
      Begin VB.TextBox txtJapaneseMessage
         Height          =   315
         Left            =   1500
         TabIndex        =   28
         Top             =   2340
         Width           =   2355
      End
   End
   Begin VB.CommandButton cmdOK
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5820
      TabIndex        =   29
      Top             =   5940
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel
      Caption         =   "Cancel"
      Cancel          =   -1  'True
      Height          =   375
      Left            =   7260
      TabIndex        =   30
      Top             =   5940
      Width           =   1275
   End
End
Attribute VB_Name = "PropertyTrackForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mEnglish As TOCCDText
Private mJapanese As TOCCDText
Private mAccepted As Boolean

Public Sub LoadFromTexts(ByRef englishText As TOCCDText, ByRef japaneseText As TOCCDText, ByVal enabled As Boolean)
    Set mEnglish = englishText
    Set mJapanese = japaneseText

    chkCDTextEnabled.Value = BoolToCheck(enabled)

    PopulateEnglish
    PopulateJapanese
    ToggleInputs
    mAccepted = False
End Sub

Public Property Get Accepted() As Boolean
    Accepted = mAccepted
End Property

Private Sub chkCDTextEnabled_Click()
    ToggleInputs
End Sub

Private Sub chkEnglishEnabled_Click()
    ToggleInputs
End Sub

Private Sub chkJapaneseEnabled_Click()
    ToggleInputs
End Sub

Private Sub chkEnglishTitle_Click()
    ToggleInputs
End Sub

Private Sub chkEnglishPerformer_Click()
    ToggleInputs
End Sub

Private Sub chkEnglishSongwriter_Click()
    ToggleInputs
End Sub

Private Sub chkEnglishComposer_Click()
    ToggleInputs
End Sub

Private Sub chkEnglishArranger_Click()
    ToggleInputs
End Sub

Private Sub chkEnglishMessage_Click()
    ToggleInputs
End Sub

Private Sub chkJapaneseTitle_Click()
    ToggleInputs
End Sub

Private Sub chkJapanesePerformer_Click()
    ToggleInputs
End Sub

Private Sub chkJapaneseSongwriter_Click()
    ToggleInputs
End Sub

Private Sub chkJapaneseComposer_Click()
    ToggleInputs
End Sub

Private Sub chkJapaneseArranger_Click()
    ToggleInputs
End Sub

Private Sub chkJapaneseMessage_Click()
    ToggleInputs
End Sub

Private Sub cmdCancel_Click()
    mAccepted = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Not ValidateInputs() Then Exit Sub

    WriteBackEnglish
    WriteBackJapanese
    mAccepted = True
    Unload Me
End Sub

Private Sub PopulateEnglish()
    chkEnglishEnabled.Value = BoolToCheck(mEnglish.LanguageEnabled)
    chkEnglishTitle.Value = BoolToCheck(mEnglish.TitleEnabled)
    txtEnglishTitle.Text = mEnglish.Title
    chkEnglishPerformer.Value = BoolToCheck(mEnglish.PerformerEnabled)
    txtEnglishPerformer.Text = mEnglish.Performer
    chkEnglishSongwriter.Value = BoolToCheck(mEnglish.SongwriterEnabled)
    txtEnglishSongwriter.Text = mEnglish.Songwriter
    chkEnglishComposer.Value = BoolToCheck(mEnglish.ComposerEnabled)
    txtEnglishComposer.Text = mEnglish.Composer
    chkEnglishArranger.Value = BoolToCheck(mEnglish.ArrangerEnabled)
    txtEnglishArranger.Text = mEnglish.Arranger
    chkEnglishMessage.Value = BoolToCheck(mEnglish.MessageEnabled)
    txtEnglishMessage.Text = mEnglish.Message
End Sub

Private Sub PopulateJapanese()
    chkJapaneseEnabled.Value = BoolToCheck(mJapanese.LanguageEnabled)
    chkJapaneseTitle.Value = BoolToCheck(mJapanese.TitleEnabled)
    txtJapaneseTitle.Text = mJapanese.Title
    chkJapanesePerformer.Value = BoolToCheck(mJapanese.PerformerEnabled)
    txtJapanesePerformer.Text = mJapanese.Performer
    chkJapaneseSongwriter.Value = BoolToCheck(mJapanese.SongwriterEnabled)
    txtJapaneseSongwriter.Text = mJapanese.Songwriter
    chkJapaneseComposer.Value = BoolToCheck(mJapanese.ComposerEnabled)
    txtJapaneseComposer.Text = mJapanese.Composer
    chkJapaneseArranger.Value = BoolToCheck(mJapanese.ArrangerEnabled)
    txtJapaneseArranger.Text = mJapanese.Arranger
    chkJapaneseMessage.Value = BoolToCheck(mJapanese.MessageEnabled)
    txtJapaneseMessage.Text = mJapanese.Message
End Sub

Private Sub ToggleInputs()
    Dim enabled As Boolean
    Dim englishEnabled As Boolean
    Dim japaneseEnabled As Boolean

    enabled = (chkCDTextEnabled.Value = 1)
    englishEnabled = enabled And (chkEnglishEnabled.Value = 1)
    japaneseEnabled = enabled And (chkJapaneseEnabled.Value = 1)

    fraEnglish.Enabled = enabled
    fraJapanese.Enabled = enabled

    chkEnglishEnabled.Enabled = enabled
    chkJapaneseEnabled.Enabled = enabled

    txtEnglishTitle.Enabled = englishEnabled And (chkEnglishTitle.Value = 1)
    txtEnglishPerformer.Enabled = englishEnabled And (chkEnglishPerformer.Value = 1)
    txtEnglishSongwriter.Enabled = englishEnabled And (chkEnglishSongwriter.Value = 1)
    txtEnglishComposer.Enabled = englishEnabled And (chkEnglishComposer.Value = 1)
    txtEnglishArranger.Enabled = englishEnabled And (chkEnglishArranger.Value = 1)
    txtEnglishMessage.Enabled = englishEnabled And (chkEnglishMessage.Value = 1)

    chkEnglishTitle.Enabled = englishEnabled
    chkEnglishPerformer.Enabled = englishEnabled
    chkEnglishSongwriter.Enabled = englishEnabled
    chkEnglishComposer.Enabled = englishEnabled
    chkEnglishArranger.Enabled = englishEnabled
    chkEnglishMessage.Enabled = englishEnabled

    txtJapaneseTitle.Enabled = japaneseEnabled And (chkJapaneseTitle.Value = 1)
    txtJapanesePerformer.Enabled = japaneseEnabled And (chkJapanesePerformer.Value = 1)
    txtJapaneseSongwriter.Enabled = japaneseEnabled And (chkJapaneseSongwriter.Value = 1)
    txtJapaneseComposer.Enabled = japaneseEnabled And (chkJapaneseComposer.Value = 1)
    txtJapaneseArranger.Enabled = japaneseEnabled And (chkJapaneseArranger.Value = 1)
    txtJapaneseMessage.Enabled = japaneseEnabled And (chkJapaneseMessage.Value = 1)

    chkJapaneseTitle.Enabled = japaneseEnabled
    chkJapanesePerformer.Enabled = japaneseEnabled
    chkJapaneseSongwriter.Enabled = japaneseEnabled
    chkJapaneseComposer.Enabled = japaneseEnabled
    chkJapaneseArranger.Enabled = japaneseEnabled
    chkJapaneseMessage.Enabled = japaneseEnabled
End Sub

Private Sub WriteBackEnglish()
    mEnglish.LanguageEnabled = (chkEnglishEnabled.Value = 1)
    mEnglish.TitleEnabled = (chkEnglishTitle.Value = 1)
    mEnglish.Title = Trim$(txtEnglishTitle.Text)
    mEnglish.PerformerEnabled = (chkEnglishPerformer.Value = 1)
    mEnglish.Performer = Trim$(txtEnglishPerformer.Text)
    mEnglish.SongwriterEnabled = (chkEnglishSongwriter.Value = 1)
    mEnglish.Songwriter = Trim$(txtEnglishSongwriter.Text)
    mEnglish.ComposerEnabled = (chkEnglishComposer.Value = 1)
    mEnglish.Composer = Trim$(txtEnglishComposer.Text)
    mEnglish.ArrangerEnabled = (chkEnglishArranger.Value = 1)
    mEnglish.Arranger = Trim$(txtEnglishArranger.Text)
    mEnglish.MessageEnabled = (chkEnglishMessage.Value = 1)
    mEnglish.Message = Trim$(txtEnglishMessage.Text)
End Sub

Private Sub WriteBackJapanese()
    mJapanese.LanguageEnabled = (chkJapaneseEnabled.Value = 1)
    mJapanese.TitleEnabled = (chkJapaneseTitle.Value = 1)
    mJapanese.Title = Trim$(txtJapaneseTitle.Text)
    mJapanese.PerformerEnabled = (chkJapanesePerformer.Value = 1)
    mJapanese.Performer = Trim$(txtJapanesePerformer.Text)
    mJapanese.SongwriterEnabled = (chkJapaneseSongwriter.Value = 1)
    mJapanese.Songwriter = Trim$(txtJapaneseSongwriter.Text)
    mJapanese.ComposerEnabled = (chkJapaneseComposer.Value = 1)
    mJapanese.Composer = Trim$(txtJapaneseComposer.Text)
    mJapanese.ArrangerEnabled = (chkJapaneseArranger.Value = 1)
    mJapanese.Arranger = Trim$(txtJapaneseArranger.Text)
    mJapanese.MessageEnabled = (chkJapaneseMessage.Value = 1)
    mJapanese.Message = Trim$(txtJapaneseMessage.Text)
End Sub

Private Function BoolToCheck(ByVal value As Boolean) As Integer
    If value Then
        BoolToCheck = 1
    Else
        BoolToCheck = 0
    End If
End Function

Private Function ValidateInputs() As Boolean
    If chkCDTextEnabled.Value <> 1 Then
        ValidateInputs = True
        Exit Function
    End If

    If Not ValidateLanguage("English", chkEnglishEnabled, chkEnglishTitle, txtEnglishTitle, chkEnglishPerformer, txtEnglishPerformer, chkEnglishSongwriter, txtEnglishSongwriter, chkEnglishComposer, txtEnglishComposer, chkEnglishArranger, txtEnglishArranger, chkEnglishMessage, txtEnglishMessage) Then Exit Function
    If Not ValidateLanguage("Japanese", chkJapaneseEnabled, chkJapaneseTitle, txtJapaneseTitle, chkJapanesePerformer, txtJapanesePerformer, chkJapaneseSongwriter, txtJapaneseSongwriter, chkJapaneseComposer, txtJapaneseComposer, chkJapaneseArranger, txtJapaneseArranger, chkJapaneseMessage, txtJapaneseMessage) Then Exit Function

    ValidateInputs = True
End Function

Private Function ValidateLanguage(ByVal languageLabel As String, ByRef languageCheck As CheckBox, ByRef titleCheck As CheckBox, ByRef titleText As TextBox, ByRef performerCheck As CheckBox, ByRef performerText As TextBox, ByRef songwriterCheck As CheckBox, ByRef songwriterText As TextBox, ByRef composerCheck As CheckBox, ByRef composerText As TextBox, ByRef arrangerCheck As CheckBox, ByRef arrangerText As TextBox, ByRef messageCheck As CheckBox, ByRef messageText As TextBox) As Boolean
    If languageCheck.Value <> 1 Then
        ValidateLanguage = True
        Exit Function
    End If

    If Not ValidateField(languageLabel, "Title", titleCheck, titleText) Then Exit Function
    If Not ValidateField(languageLabel, "Performer", performerCheck, performerText) Then Exit Function
    If Not ValidateField(languageLabel, "Songwriter", songwriterCheck, songwriterText) Then Exit Function
    If Not ValidateField(languageLabel, "Composer", composerCheck, composerText) Then Exit Function
    If Not ValidateField(languageLabel, "Arranger", arrangerCheck, arrangerText) Then Exit Function
    If Not ValidateField(languageLabel, "Message", messageCheck, messageText) Then Exit Function

    ValidateLanguage = True
End Function

Private Function ValidateField(ByVal languageLabel As String, ByVal fieldLabel As String, ByRef fieldCheck As CheckBox, ByRef fieldText As TextBox) As Boolean
    If fieldCheck.Value = 1 And Trim$(fieldText.Text) = "" Then
        MsgBox languageLabel & " " & fieldLabel & " is enabled but empty.", vbExclamation, "Track CD-TEXT"
        fieldText.SetFocus
        ValidateField = False
        Exit Function
    End If

    ValidateField = True
End Function

