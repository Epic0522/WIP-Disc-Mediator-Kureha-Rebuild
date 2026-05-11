VERSION 5.00
Begin VB.Form MainForm
   Appearance      =   0  'Flat
   BackColor       =   &H8000000F&
   Caption         =   "呉葉"
   ClientHeight    =   9495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13815
   BeginProperty Font
      Name            =   "MS UI Gothic"
      Size            =   9
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "MainForm"
   ScaleHeight     =   9495
   ScaleWidth      =   13815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraToolbar
      Caption         =   "操作"
      Height          =   1035
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13815
      Begin VB.CommandButton cmdWriteDisc
         Caption         =   "書き込み"
         Height          =   495
         Left            =   150
         TabIndex        =   1
         Top             =   285
         Width           =   1215
      End
      Begin VB.CommandButton cmdSaveProject
         Caption         =   "保存"
         Height          =   495
         Left            =   1440
         TabIndex        =   2
         Top             =   285
         Width           =   1215
      End
      Begin VB.CommandButton cmdEraseDisc
         Caption         =   "消去"
         Height          =   495
         Left            =   2730
         TabIndex        =   3
         Top             =   285
         Width           =   1215
      End
      Begin VB.CommandButton cmdCopyDisc
         Caption         =   "複製"
         Height          =   495
         Left            =   4020
         TabIndex        =   4
         Top             =   285
         Width           =   1215
      End
      Begin VB.CommandButton cmdImageWrite
         Caption         =   "媒体書込"
         Height          =   495
         Left            =   5310
         TabIndex        =   5
         Top             =   285
         Width           =   1335
      End
      Begin VB.CommandButton cmdImageRead
         Caption         =   "媒体読込"
         Height          =   495
         Left            =   6720
         TabIndex        =   6
         Top             =   285
         Width           =   1335
      End
      Begin VB.CommandButton cmdReadTracks
         Caption         =   "取込み"
         Height          =   495
         Left            =   8130
         TabIndex        =   7
         Top             =   285
         Width           =   1215
      End
      Begin VB.CommandButton cmdAnalyzeDisc
         Caption         =   "分析"
         Height          =   495
         Left            =   9420
         TabIndex        =   8
         Top             =   285
         Width           =   1215
      End
   End
   Begin VB.Frame fraDisc
      Caption         =   "構成"
      Height          =   1335
      Left            =   0
      TabIndex        =   9
      Top             =   1140
      Width           =   13815
      Begin VB.Label lblDiscLabel
         Caption         =   "媒体識別子:"
         Height          =   255
         Left            =   180
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtDiscLabel
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1320
         TabIndex        =   11
         Top             =   300
         Width           =   4875
      End
      Begin VB.Label lblFileSystem
         Caption         =   "ファイル系:"
         Height          =   255
         Left            =   180
         TabIndex        =   12
         Top             =   780
         Width           =   1095
      End
      Begin VB.ComboBox cboFileSystem
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   720
         Width           =   4875
      End
      Begin VB.Frame fraUsage
         Caption         =   "使用量"
         Height          =   855
         Left            =   6360
         TabIndex        =   14
         Top             =   240
         Width           =   7230
         Begin VB.Label lblUsageValue
            Alignment       =   2  'Center
            Caption         =   "空き:648.000KiB"
            Height          =   255
            Left            =   4710
            TabIndex        =   18
            Top             =   165
            Width           =   1875
         End
         Begin VB.ComboBox cboMediaType
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5280
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   480
            Width           =   1710
         End
         Begin VB.Shape shpDiscOuter
            BorderColor     =   &H00FFFF00&
            BorderWidth     =   2
            Height          =   615
            Left            =   1290
            Shape           =   3  'Circle
            Top             =   120
            Width           =   2115
         End
         Begin VB.Shape shpDiscInner
            BorderColor     =   &H00FFFF00&
            BorderWidth     =   2
            Height          =   255
            Left            =   2010
            Shape           =   3  'Circle
            Top             =   300
            Width           =   675
         End
         Begin VB.Label lblUsageCaption
            Caption         =   "使用量:"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   240
            Width           =   795
         End
      End
   End
   Begin VB.Frame fraExplorer
      Caption         =   "構造配置"
      Height          =   4335
      Left            =   0
      TabIndex        =   20
      Top             =   2460
      Width           =   13815
      Begin VB.Frame fraDirectories
         Caption         =   "ディレクトリ"
         Height          =   3795
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   4515
         Begin VB.ListBox lstDirectories
            Appearance      =   0  'Flat
            Height          =   3180
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   4215
         End
      End
      Begin VB.Frame fraFiles
         Caption         =   "ファイル"
         Height          =   3795
         Left            =   4740
         TabIndex        =   23
         Top             =   360
         Width           =   8955
         Begin VB.Frame fraFileActions
            Caption         =   "ファイル操作"
            Height          =   615
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   8715
            Begin VB.CommandButton cmdAddFile
               Caption         =   "追加"
               Height          =   315
               Left            =   180
               TabIndex        =   25
               Top             =   210
               Width           =   915
            End
            Begin VB.CommandButton cmdAddFolder
               Caption         =   "フォルダ"
               Height          =   315
               Left            =   1155
               TabIndex        =   26
               Top             =   210
               Width           =   915
            End
            Begin VB.CommandButton cmdRenameEntry
               Caption         =   "名前変更"
               Height          =   315
               Left            =   2130
               TabIndex        =   27
               Top             =   210
               Width           =   915
            End
            Begin VB.CommandButton cmdRemoveEntry
               Caption         =   "削除"
               Height          =   315
               Left            =   3105
               TabIndex        =   28
               Top             =   210
               Width           =   915
            End
            Begin VB.CommandButton cmdProperties
               Caption         =   "設定"
               Height          =   315
               Left            =   4080
               TabIndex        =   29
               Top             =   210
               Width           =   915
            End
         End
         Begin VB.Label lblFilesHeader
            Caption         =   "名前                                             容量        更新日時"
            Height          =   255
            Left            =   180
            TabIndex        =   30
            Top             =   1140
            Width           =   8115
         End
         Begin VB.ListBox lstFiles
            Appearance      =   0  'Flat
            Height          =   2250
            Left            =   120
            TabIndex        =   31
            Top             =   1140
            Width           =   8715
         End
      End
   End
   Begin VB.Frame fraTrackArea
      Caption         =   "トラック"
      Height          =   2235
      Left            =   0
      TabIndex        =   32
      Top             =   6780
      Width           =   13815
      Begin VB.Frame fraTrackActions
         Caption         =   "トラック操作"
         Height          =   615
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   3135
         Begin VB.CommandButton cmdAddTrack
            Caption         =   "+"
            Height          =   315
            Left            =   180
            TabIndex        =   34
            Top             =   210
            Width           =   435
         End
         Begin VB.CommandButton cmdRemoveTrack
            Caption         =   "x"
            Height          =   315
            Left            =   660
            TabIndex        =   35
            Top             =   210
            Width           =   435
          End
         Begin VB.CommandButton cmdMoveTrackUp
            Caption         =   "^"
            Height          =   315
            Left            =   1140
            TabIndex        =   36
            Top             =   210
            Width           =   435
         End
         Begin VB.CommandButton cmdMoveTrackDown
            Caption         =   "v"
            Height          =   315
            Left            =   1620
            TabIndex        =   37
            Top             =   210
            Width           =   435
         End
         Begin VB.CommandButton cmdTrackProperties
            Caption         =   "CD"
            Height          =   315
            Left            =   2250
            TabIndex        =   38
            Top             =   210
            Width           =   615
         End
      End
      Begin VB.TextBox txtAlbumName
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3780
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   330
         Width           =   3675
      End
      Begin VB.Label lblAlbumNameCaption
         Caption         =   "Album Name:"
         Height          =   255
         Left            =   3780
         TabIndex        =   39
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label lblTracksHeader
         Caption         =   "トラック      曲名                           演奏者                         ソース      Pregap   時間   Postgap   属性"
         Height          =   255
         Left            =   180
         TabIndex        =   41
         Top             =   960
         Width           =   13155
      End
      Begin VB.ListBox lstTracks
         Appearance      =   0  'Flat
         Height          =   1080
         Left            =   120
         TabIndex        =   42
         Top             =   1200
         Width           =   13515
      End
   End
   Begin VB.Label lblStatus
      BackColor       =   &H8000000F&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ready"
      Height          =   255
      Left            =   0
      TabIndex        =   43
      Top             =   9180
      Width           =   13815
   End
   Begin Menu mnuFile
      Caption         =   "ファイル(&F)"
      Begin Menu mnuFileOpenImage
         Caption         =   "イメージを開く(&O)..."
      End
      Begin Menu mnuFileLoadProject
         Caption         =   "プロジェクトを開く(&R)..."
      End
      Begin Menu mnuFileSaveProject
         Caption         =   "プロジェクト保存(&S)..."
      End
      Begin Menu mnuFileSep1
         Caption         =   "-"
      End
      Begin Menu mnuFileWriteDisc
         Caption         =   "書き込み(&W)..."
      End
      Begin Menu mnuFileSaveImage
         Caption         =   "イメージ保存(&A)..."
      End
      Begin Menu mnuFileNew
         Caption         =   "新規作成(&N)"
      End
      Begin Menu mnuFileSep2
         Caption         =   "-"
      End
      Begin Menu mnuFileExit
         Caption         =   "終了(&X)"
      End
   End
   Begin Menu mnuView
      Caption         =   "表示(&V)"
      Begin Menu mnuViewExplorer
         Caption         =   "エクスプローラ(&E)"
      End
      Begin Menu mnuViewAlwaysOnTop
         Caption         =   "常に手前に表示(&T)"
      End
   End
   Begin Menu mnuComposition
      Caption         =   "構造配置(&W)"
      Begin Menu mnuCompositionAddFile
         Caption         =   "ファイル追加(&A)..."
      End
      Begin Menu mnuCompositionAddFolder
         Caption         =   "フォルダ追加(&F)..."
      End
      Begin Menu mnuCompositionRename
         Caption         =   "名前変更(&R)"
      End
      Begin Menu mnuCompositionRemove
         Caption         =   "削除(&D)"
      End
      Begin Menu mnuCompositionClear
         Caption         =   "クリア(&C)"
      End
   End
   Begin Menu mnuTrack
      Caption         =   "トラック(&T)"
      Begin Menu mnuTrackAdd
         Caption         =   "追加(&A)..."
      End
      Begin Menu mnuTrackRemove
         Caption         =   "削除(&D)"
      End
      Begin Menu mnuTrackPropertiesMenu
         Caption         =   "設定(&P)..."
      End
      Begin Menu mnuTrackClear
         Caption         =   "クリア(&C)"
      End
   End
   Begin Menu mnuTools
      Caption         =   "ツール(&U)"
      Begin Menu mnuToolsCopy
         Caption         =   "ディスク複製(&C)..."
      End
      Begin Menu mnuToolsErase
         Caption         =   "ディスク消去(&E)..."
      End
      Begin Menu mnuToolsReadTracks
         Caption         =   "トラック取り込み(&I)..."
      End
      Begin Menu mnuToolsReadImage
         Caption         =   "イメージ読み込み(&R)..."
      End
      Begin Menu mnuToolsWriteImage
         Caption         =   "イメージ書き込み(&W)..."
      End
      Begin Menu mnuToolsAnalyze
         Caption         =   "ディスク分析(&A)..."
      End
   End
   Begin Menu mnuHelp
      Caption         =   "ヘルプ(&H)"
      Begin Menu mnuHelpAbout
         Caption         =   "バージョン情報(&A)..."
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mTocInfo As TOCInformation
Private mTrackEntries As Collection
Private mFileEntries As Collection
Private mCurrentTrack As TrackEntry

Private Sub cmdAddFile_Click()
    Dim fileName As String
    Dim entry As FileEntry

    fileName = Trim$(InputBox$("File name to add:", "Add File", "newfile.bin"))
    If fileName = "" Then Exit Sub

    Set entry = New FileEntry
    entry.Name = fileName
    entry.SizeText = "0 KB"
    entry.ModifiedText = Format$(Date, "yyyy/mm/dd")
    entry.IsDirectory = False
    mFileEntries.Add entry

    RefreshFileDisplay
    lblStatus.Caption = "Added file: " & fileName
End Sub

Private Sub cmdAddFolder_Click()
    Dim folderName As String
    Dim entry As FileEntry

    folderName = Trim$(InputBox$("Folder name to add:", "Add Folder", "NewFolder"))
    If folderName = "" Then Exit Sub

    Set entry = New FileEntry
    entry.Name = folderName & "\"
    entry.SizeText = "<DIR>"
    entry.ModifiedText = Format$(Date, "yyyy/mm/dd")
    entry.IsDirectory = True
    mFileEntries.Add entry

    RefreshFileDisplay
    lblStatus.Caption = "Added folder: " & folderName
End Sub

Private Sub cmdAddTrack_Click()
    Dim entry As TrackEntry
    Dim defaultTitle As String

    Set entry = New TrackEntry
    entry.TrackNo = mTrackEntries.Count + 1
    entry.EnglishText.LanguageEnabled = True
    entry.EnglishText.TitleEnabled = True
    entry.EnglishText.PerformerEnabled = True

    defaultTitle = Trim$(InputBox$("Track title:", "Add Track", "Track " & Format$(entry.TrackNo, "00")))
    If defaultTitle = "" Then defaultTitle = "Track " & Format$(entry.TrackNo, "00")

    entry.EnglishText.Title = defaultTitle
    entry.EnglishText.Performer = ""
    mTrackEntries.Add entry
    Set mCurrentTrack = entry

    RefreshTrackDisplay
    SelectTrackIndex mTrackEntries.Count - 1
    lblStatus.Caption = "Added track: " & defaultTitle
End Sub

Private Sub cmdAnalyzeDisc_Click()
    DiscAnalyzeForm.LoadPreview txtDiscLabel.Text, cboMediaType.Text, TrackCount(), ProjectHasCdText()
    DiscAnalyzeForm.Show vbModal, Me
End Sub

Private Sub cmdCopyDisc_Click()
    DiscCopyForm.LoadPreview cboMediaType.Text
    DiscCopyForm.Show vbModal, Me
End Sub

Private Sub cmdEraseDisc_Click()
    DiscEraseForm.LoadPreview cboMediaType.Text
    DiscEraseForm.Show vbModal, Me
End Sub

Private Sub cmdImageRead_Click()
    DiscReadForm.LoadPreview cboMediaType.Text
    DiscReadForm.Show vbModal, Me
End Sub

Private Sub cmdImageWrite_Click()
    DiscWriteForm.LoadPreview txtDiscLabel.Text, cboMediaType.Text, TrackCount(), ProjectHasCdText()
    DiscWriteForm.Show vbModal, Me
End Sub

Private Sub cmdMoveTrackDown_Click()
    MoveSelectedTrack 1
End Sub

Private Sub cmdMoveTrackUp_Click()
    MoveSelectedTrack -1
End Sub

Private Sub cmdProperties_Click()
    Dim entry As FileEntry
    Dim selectedIndex As Long

    selectedIndex = lstFiles.ListIndex
    If selectedIndex < 0 Then
        MsgBox "Select a file or folder first.", vbExclamation, "Kureha VB6 Rebuild"
        Exit Sub
    End If

    Set entry = mFileEntries.Item(selectedIndex + 1)
    MsgBox "Name: " & entry.Name & vbCrLf & _
        "Size: " & entry.SizeText & vbCrLf & _
        "Modified: " & entry.ModifiedText, vbInformation, "Entry Properties"
End Sub

Private Sub cmdReadTracks_Click()
    TrackRippingForm.LoadPreview TrackCount()
    TrackRippingForm.Show vbModal, Me
End Sub

Private Sub cmdRemoveEntry_Click()
    Dim selectedIndex As Long

    selectedIndex = lstFiles.ListIndex
    If selectedIndex < 0 Then
        MsgBox "Select a file or folder to remove.", vbExclamation, "Kureha VB6 Rebuild"
        Exit Sub
    End If

    If MsgBox("Remove the selected entry?", vbYesNo + vbQuestion, "Kureha VB6 Rebuild") <> vbYes Then Exit Sub

    mFileEntries.Remove selectedIndex + 1
    RefreshFileDisplay
    lblStatus.Caption = "Entry removed."
End Sub

Private Sub cmdRemoveTrack_Click()
    Dim selectedIndex As Long

    selectedIndex = lstTracks.ListIndex
    If selectedIndex < 0 Then
        MsgBox "Select a track to remove.", vbExclamation, "Kureha VB6 Rebuild"
        Exit Sub
    End If

    If MsgBox("Remove the selected track?", vbYesNo + vbQuestion, "Kureha VB6 Rebuild") <> vbYes Then Exit Sub

    mTrackEntries.Remove selectedIndex + 1
    RenumberTracks
    If mTrackEntries.Count > 0 Then
        Set mCurrentTrack = mTrackEntries.Item(1)
    Else
        Set mCurrentTrack = Nothing
    End If
    RefreshTrackDisplay
    lblStatus.Caption = "Track removed."
End Sub

Private Sub cmdRenameEntry_Click()
    Dim selectedIndex As Long
    Dim entry As FileEntry
    Dim newName As String

    selectedIndex = lstFiles.ListIndex
    If selectedIndex < 0 Then
        MsgBox "Select a file or folder to rename.", vbExclamation, "Kureha VB6 Rebuild"
        Exit Sub
    End If

    Set entry = mFileEntries.Item(selectedIndex + 1)
    newName = Trim$(InputBox$("New name:", "Rename Entry", entry.Name))
    If newName = "" Then Exit Sub

    If entry.IsDirectory And Right$(newName, 1) <> "\" Then
        newName = newName & "\"
    End If

    entry.Name = newName
    RefreshFileDisplay
    SelectFileIndex selectedIndex
    lblStatus.Caption = "Entry renamed."
End Sub

Private Sub cmdSaveProject_Click()
    Dim savePath As String

    savePath = Trim$(InputBox$("Save project path:", "Save Project", App.Path & "\kureha-rebuild-project.krp"))
    If savePath = "" Then Exit Sub

    SaveProjectToPath savePath
End Sub

Private Sub cmdTrackProperties_Click()
    If mCurrentTrack Is Nothing Then
        MsgBox "Add or select a track first.", vbExclamation, "Kureha VB6 Rebuild"
        Exit Sub
    End If

    PropertyTrackForm.LoadFromTexts mCurrentTrack.EnglishText, mCurrentTrack.JapaneseText, True
    PropertyTrackForm.Show vbModal, Me
    RefreshTrackDisplay
End Sub

Private Sub cmdWriteDisc_Click()
    DiscWriteForm.LoadPreview txtDiscLabel.Text, cboMediaType.Text, TrackCount(), ProjectHasCdText()
    DiscWriteForm.Show vbModal, Me
End Sub

Private Sub Form_Load()
    Me.Font.Name = "MS UI Gothic"
    Me.Font.Size = 9

    Set mTocInfo = New TOCInformation
    Set mTrackEntries = New Collection
    Set mFileEntries = New Collection

    cboFileSystem.AddItem "JIS X 0606 (Joliet+ Extension)"
    cboFileSystem.AddItem "ISO 9660 Level 1"
    cboFileSystem.AddItem "ISO 9660 Level 2"
    cboFileSystem.ListIndex = 0

    cboMediaType.AddItem "CD-R 74min"
    cboMediaType.AddItem "CD-R 80min"
    cboMediaType.AddItem "CD-RW 74min"
    cboMediaType.ListIndex = 0

    txtDiscLabel.Text = "VOL_202605111641"
    txtAlbumName.Text = "No Album Name"

    lstDirectories.Font.Name = "MS UI Gothic"
    lstFiles.Font.Name = "MS Gothic"
    lstFiles.Font.Size = 10
    lstTracks.Font.Name = "MS Gothic"
    lstTracks.Font.Size = 10

    SeedDefaultDirectories

    SeedSampleFiles
    SeedSampleTracks

    RefreshTrackDisplay
    RefreshFileDisplay
    lblStatus.Caption = "UI scaffold loaded. Track CD-TEXT editor is wired."
End Sub

Private Sub lstTracks_Click()
    Dim selectedIndex As Long

    selectedIndex = lstTracks.ListIndex
    If selectedIndex < 0 Then Exit Sub

    Set mCurrentTrack = mTrackEntries.Item(selectedIndex + 1)
    txtAlbumName.Text = mCurrentTrack.DisplayTitle
End Sub

Private Sub lstTracks_DblClick()
    If lstTracks.ListIndex >= 0 Then
        cmdTrackProperties_Click
    End If
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileLoadProject_Click()
    Dim loadPath As String

    loadPath = Trim$(InputBox$("Load project path:", "Load Project", App.Path & "\kureha-rebuild-project.krp"))
    If loadPath = "" Then Exit Sub

    LoadProjectFromPath loadPath
End Sub

Private Sub mnuFileNew_Click()
    If MsgBox("Clear files and tracks for a new project?", vbYesNo + vbQuestion, "Kureha VB6 Rebuild") <> vbYes Then Exit Sub

    Set mFileEntries = New Collection
    Set mTrackEntries = New Collection
    Set mCurrentTrack = Nothing
    txtDiscLabel.Text = "VOL_" & Format$(Now, "yyyymmddhhnnss")
    txtAlbumName.Text = "No Album Name"
    RefreshFileDisplay
    RefreshTrackDisplay
    lblStatus.Caption = "New project created."
End Sub

Private Sub mnuFileOpenImage_Click()
    mnuFileLoadProject_Click
End Sub

Private Sub mnuFileSaveImage_Click()
    cmdSaveProject_Click
End Sub

Private Sub mnuFileSaveProject_Click()
    cmdSaveProject_Click
End Sub

Private Sub mnuFileWriteDisc_Click()
    cmdWriteDisc_Click
End Sub

Private Sub mnuHelpAbout_Click()
    AboutInformationForm.LoadPreview "Version 0.1"
    AboutInformationForm.Show vbModal, Me
End Sub

Private Sub mnuCompositionAddFile_Click()
    cmdAddFile_Click
End Sub

Private Sub mnuCompositionAddFolder_Click()
    cmdAddFolder_Click
End Sub

Private Sub mnuCompositionClear_Click()
    If MsgBox("Clear all files and folders?", vbYesNo + vbQuestion, "Kureha VB6 Rebuild") <> vbYes Then Exit Sub

    Set mFileEntries = New Collection
    RefreshFileDisplay
    lblStatus.Caption = "All file entries cleared."
End Sub

Private Sub mnuCompositionRemove_Click()
    cmdRemoveEntry_Click
End Sub

Private Sub mnuCompositionRename_Click()
    cmdRenameEntry_Click
End Sub

Private Sub mnuToolsAnalyze_Click()
    cmdAnalyzeDisc_Click
End Sub

Private Sub mnuToolsCopy_Click()
    cmdCopyDisc_Click
End Sub

Private Sub mnuToolsErase_Click()
    cmdEraseDisc_Click
End Sub

Private Sub mnuToolsReadImage_Click()
    cmdImageRead_Click
End Sub

Private Sub mnuToolsReadTracks_Click()
    cmdReadTracks_Click
End Sub

Private Sub mnuToolsWriteImage_Click()
    cmdImageWrite_Click
End Sub

Private Sub mnuTrackAdd_Click()
    cmdAddTrack_Click
End Sub

Private Sub mnuTrackClear_Click()
    If MsgBox("Clear all tracks?", vbYesNo + vbQuestion, "Kureha VB6 Rebuild") <> vbYes Then Exit Sub

    Set mTrackEntries = New Collection
    Set mCurrentTrack = Nothing
    RefreshTrackDisplay
    lblStatus.Caption = "All tracks cleared."
End Sub

Private Sub mnuTrackPropertiesMenu_Click()
    cmdTrackProperties_Click
End Sub

Private Sub mnuTrackRemove_Click()
    cmdRemoveTrack_Click
End Sub

Private Sub mnuViewAlwaysOnTop_Click()
    ShowPlaceholder "Always-on-top behavior is not wired yet."
End Sub

Private Sub mnuViewExplorer_Click()
    On Error GoTo ExplorerError

    Shell "explorer.exe", vbNormalFocus
    lblStatus.Caption = "Windows Explorer opened."
    Exit Sub

ExplorerError:
    MsgBox "Could not open Windows Explorer.", vbExclamation, "Kureha VB6 Rebuild"
End Sub

Private Sub RefreshFileDisplay()
    Dim i As Long
    Dim entry As FileEntry

    lstFiles.Clear

    For i = 1 To mFileEntries.Count
        Set entry = mFileEntries.Item(i)
        lstFiles.AddItem PadRight(entry.Name, 42) & PadLeft(entry.SizeText, 9) & "   " & entry.ModifiedText
    Next i
End Sub

Private Sub RefreshTrackDisplay()
    Dim i As Long
    Dim entry As TrackEntry

    lstTracks.Clear

    If mTrackEntries.Count = 0 Then
        txtAlbumName.Text = "No Album Name"
        lblStatus.Caption = "No tracks loaded."
        Exit Sub
    End If

    For i = 1 To mTrackEntries.Count
        Set entry = mTrackEntries.Item(i)
        lstTracks.AddItem PadRight(Format$(entry.TrackNo, "00"), 9) & _
            PadRight(entry.DisplayTitle, 28) & _
            PadRight(entry.DisplayPerformer, 28) & _
            PadRight(entry.Source, 14) & _
            PadRight(entry.Pregap, 9) & _
            PadRight(entry.Duration, 8) & _
            PadRight(entry.Postgap, 9) & _
            entry.Flags
    Next i

    If mCurrentTrack Is Nothing Then Set mCurrentTrack = mTrackEntries.Item(1)
    txtAlbumName.Text = mCurrentTrack.DisplayTitle
    lblStatus.Caption = "Track view refreshed."
End Sub

Private Sub RenumberTracks()
    Dim i As Long

    For i = 1 To mTrackEntries.Count
        mTrackEntries.Item(i).TrackNo = i
    Next i
End Sub

Private Sub MoveSelectedTrack(ByVal direction As Long)
    Dim selectedIndex As Long
    Dim targetIndex As Long
    Dim reordered As New Collection
    Dim i As Long
    Dim movingEntry As TrackEntry

    selectedIndex = lstTracks.ListIndex
    If selectedIndex < 0 Then
        MsgBox "Select a track first.", vbExclamation, "Kureha VB6 Rebuild"
        Exit Sub
    End If

    targetIndex = selectedIndex + direction
    If targetIndex < 0 Or targetIndex >= mTrackEntries.Count Then Exit Sub

    Set movingEntry = mTrackEntries.Item(selectedIndex + 1)

    For i = 1 To mTrackEntries.Count
        If i - 1 = targetIndex Then
            reordered.Add movingEntry
        End If

        If i - 1 <> selectedIndex Then
            reordered.Add mTrackEntries.Item(i)
        End If
    Next i

    If targetIndex = mTrackEntries.Count - 1 Then
        If reordered.Count < mTrackEntries.Count Then reordered.Add movingEntry
    End If

    Set mTrackEntries = reordered
    RenumberTracks
    Set mCurrentTrack = mTrackEntries.Item(targetIndex + 1)
    RefreshTrackDisplay
    SelectTrackIndex targetIndex
    lblStatus.Caption = "Track order updated."
End Sub

Private Sub SeedSampleFiles()
    Dim entry As FileEntry

    Set entry = New FileEntry
    entry.Name = "README.TXT"
    entry.SizeText = "12 KB"
    entry.ModifiedText = "2026/05/11"
    mFileEntries.Add entry

    Set entry = New FileEntry
    entry.Name = "cover.jpg"
    entry.SizeText = "421 KB"
    entry.ModifiedText = "2026/05/11"
    mFileEntries.Add entry

    Set entry = New FileEntry
    entry.Name = "bonus\"
    entry.SizeText = "<DIR>"
    entry.ModifiedText = "2026/05/11"
    entry.IsDirectory = True
    mFileEntries.Add entry
End Sub

Private Sub SeedSampleTracks()
    Dim entry As TrackEntry

    mTocInfo.MCN = "1234567890123"
    mTocInfo.TOCAdd 1, 1, 1, 0, 0, 0, 18000
    mTocInfo.RawTOCAdd 1, 65, 1, 1, 0, 2, 0, 0, 0, 4, 0

    Set entry = New TrackEntry
    entry.TrackNo = 1
    entry.EnglishText.LanguageEnabled = True
    entry.EnglishText.TitleEnabled = True
    entry.EnglishText.Title = "Sample Track"
    entry.EnglishText.PerformerEnabled = True
    entry.EnglishText.Performer = "Sample Performer"
    entry.JapaneseText.LanguageEnabled = True
    entry.JapaneseText.TitleEnabled = True
    entry.JapaneseText.Title = "Sample JP Title"

    mTrackEntries.Add entry
    Set mCurrentTrack = entry
End Sub

Private Sub SeedDefaultDirectories()
    lstDirectories.Clear
    lstDirectories.AddItem "\"
    lstDirectories.AddItem "Audio"
    lstDirectories.AddItem "Data"
    lstDirectories.AddItem "Extras"
    lstDirectories.ListIndex = 0
End Sub

Private Sub SaveProjectToPath(ByVal savePath As String)
    Dim fileNo As Integer
    Dim i As Long
    Dim fileEntry As FileEntry
    Dim trackEntry As TrackEntry

    On Error GoTo SaveError

    fileNo = FreeFile
    Open savePath For Output As #fileNo

    Print #fileNo, "KUREHA_REBUILD_PROJECT|1"
    Print #fileNo, "DISC_LABEL|" & EscapeValue(txtDiscLabel.Text)
    Print #fileNo, "FILE_SYSTEM|" & CStr(cboFileSystem.ListIndex) & "|" & EscapeValue(cboFileSystem.Text)
    Print #fileNo, "MEDIA_TYPE|" & CStr(cboMediaType.ListIndex) & "|" & EscapeValue(cboMediaType.Text)
    Print #fileNo, "ALBUM|" & EscapeValue(txtAlbumName.Text)

    For i = 0 To lstDirectories.ListCount - 1
        Print #fileNo, "DIRECTORY|" & EscapeValue(lstDirectories.List(i))
    Next i

    For i = 1 To mFileEntries.Count
        Set fileEntry = mFileEntries.Item(i)
        Print #fileNo, "FILE|" & EscapeValue(fileEntry.Name) & "|" & EscapeValue(fileEntry.SizeText) & "|" & EscapeValue(fileEntry.ModifiedText) & "|" & BoolText(fileEntry.IsDirectory)
    Next i

    For i = 1 To mTrackEntries.Count
        Set trackEntry = mTrackEntries.Item(i)
        Print #fileNo, "TRACK|" & CStr(trackEntry.TrackNo) & "|" & EscapeValue(trackEntry.Source) & "|" & EscapeValue(trackEntry.Pregap) & "|" & EscapeValue(trackEntry.Duration) & "|" & EscapeValue(trackEntry.Postgap) & "|" & EscapeValue(trackEntry.Flags)
        WriteCdTextBlock fileNo, "EN", trackEntry.EnglishText
        WriteCdTextBlock fileNo, "JP", trackEntry.JapaneseText
    Next i

    Close #fileNo
    lblStatus.Caption = "Project saved: " & savePath
    Exit Sub

SaveError:
    If fileNo <> 0 Then Close #fileNo
    MsgBox "Could not save project." & vbCrLf & Err.Description, vbExclamation, "Kureha VB6 Rebuild"
End Sub

Private Sub LoadProjectFromPath(ByVal loadPath As String)
    Dim fileNo As Integer
    Dim lineText As String
    Dim parts() As String
    Dim currentTrack As TrackEntry
    Dim currentFile As FileEntry

    On Error GoTo LoadError

    fileNo = FreeFile
    Open loadPath For Input As #fileNo

    Set mFileEntries = New Collection
    Set mTrackEntries = New Collection
    Set mCurrentTrack = Nothing
    lstDirectories.Clear

    Do While Not EOF(fileNo)
        Line Input #fileNo, lineText
        parts = Split(lineText, "|")

        Select Case parts(0)
            Case "KUREHA_REBUILD_PROJECT"
            Case "DISC_LABEL"
                txtDiscLabel.Text = UnescapeValue(SafeField(parts, 1))
            Case "FILE_SYSTEM"
                RestoreComboSelection cboFileSystem, CLng(Val(SafeField(parts, 1))), SafeField(parts, 2)
            Case "MEDIA_TYPE"
                RestoreComboSelection cboMediaType, CLng(Val(SafeField(parts, 1))), SafeField(parts, 2)
            Case "ALBUM"
                txtAlbumName.Text = UnescapeValue(SafeField(parts, 1))
            Case "DIRECTORY"
                lstDirectories.AddItem UnescapeValue(SafeField(parts, 1))
            Case "FILE"
                Set currentFile = New FileEntry
                currentFile.Name = UnescapeValue(SafeField(parts, 1))
                currentFile.SizeText = UnescapeValue(SafeField(parts, 2))
                currentFile.ModifiedText = UnescapeValue(SafeField(parts, 3))
                currentFile.IsDirectory = TextBool(SafeField(parts, 4))
                mFileEntries.Add currentFile
            Case "TRACK"
                Set currentTrack = New TrackEntry
                currentTrack.TrackNo = CLng(Val(SafeField(parts, 1)))
                currentTrack.Source = UnescapeValue(SafeField(parts, 2))
                currentTrack.Pregap = UnescapeValue(SafeField(parts, 3))
                currentTrack.Duration = UnescapeValue(SafeField(parts, 4))
                currentTrack.Postgap = UnescapeValue(SafeField(parts, 5))
                currentTrack.Flags = UnescapeValue(SafeField(parts, 6))
                mTrackEntries.Add currentTrack
            Case "TEXT"
                If Not currentTrack Is Nothing Then
                    If UCase$(SafeField(parts, 1)) = "EN" Then
                        ReadCdTextBlock currentTrack.EnglishText, parts
                    ElseIf UCase$(SafeField(parts, 1)) = "JP" Then
                        ReadCdTextBlock currentTrack.JapaneseText, parts
                    End If
                End If
        End Select
    Loop

    Close #fileNo

    If lstDirectories.ListCount = 0 Then SeedDefaultDirectories
    If mTrackEntries.Count > 0 Then
        Set mCurrentTrack = mTrackEntries.Item(1)
        txtAlbumName.Text = mCurrentTrack.DisplayTitle
    End If
    RefreshFileDisplay
    RefreshTrackDisplay
    lblStatus.Caption = "Project loaded: " & loadPath
    Exit Sub

LoadError:
    If fileNo <> 0 Then Close #fileNo
    MsgBox "Could not load project." & vbCrLf & Err.Description, vbExclamation, "Kureha VB6 Rebuild"
End Sub

Private Sub WriteCdTextBlock(ByVal fileNo As Integer, ByVal languageCode As String, ByVal textItem As TOCCDText)
    Print #fileNo, "TEXT|" & languageCode & "|" & _
        BoolText(textItem.LanguageEnabled) & "|" & _
        BoolText(textItem.TitleEnabled) & "|" & _
        BoolText(textItem.PerformerEnabled) & "|" & _
        BoolText(textItem.SongwriterEnabled) & "|" & _
        BoolText(textItem.ComposerEnabled) & "|" & _
        BoolText(textItem.ArrangerEnabled) & "|" & _
        BoolText(textItem.MessageEnabled) & "|" & _
        EscapeValue(textItem.Title) & "|" & _
        EscapeValue(textItem.Performer) & "|" & _
        EscapeValue(textItem.Songwriter) & "|" & _
        EscapeValue(textItem.Composer) & "|" & _
        EscapeValue(textItem.Arranger) & "|" & _
        EscapeValue(textItem.Message)
End Sub

Private Sub ReadCdTextBlock(ByVal textItem As TOCCDText, ByRef parts() As String)
    textItem.LanguageEnabled = TextBool(SafeField(parts, 2))
    textItem.TitleEnabled = TextBool(SafeField(parts, 3))
    textItem.PerformerEnabled = TextBool(SafeField(parts, 4))
    textItem.SongwriterEnabled = TextBool(SafeField(parts, 5))
    textItem.ComposerEnabled = TextBool(SafeField(parts, 6))
    textItem.ArrangerEnabled = TextBool(SafeField(parts, 7))
    textItem.MessageEnabled = TextBool(SafeField(parts, 8))
    textItem.Title = UnescapeValue(SafeField(parts, 9))
    textItem.Performer = UnescapeValue(SafeField(parts, 10))
    textItem.Songwriter = UnescapeValue(SafeField(parts, 11))
    textItem.Composer = UnescapeValue(SafeField(parts, 12))
    textItem.Arranger = UnescapeValue(SafeField(parts, 13))
    textItem.Message = UnescapeValue(SafeField(parts, 14))
End Sub

Private Sub RestoreComboSelection(ByRef combo As ComboBox, ByVal indexValue As Long, ByVal textValue As String)
    If indexValue >= 0 And indexValue < combo.ListCount Then
        combo.ListIndex = indexValue
    ElseIf Trim$(textValue) <> "" Then
        combo.AddItem UnescapeValue(textValue)
        combo.ListIndex = combo.ListCount - 1
    End If
End Sub

Private Function SafeField(ByRef parts() As String, ByVal indexValue As Long) As String
    If indexValue <= UBound(parts) Then
        SafeField = parts(indexValue)
    Else
        SafeField = ""
    End If
End Function

Private Function BoolText(ByVal value As Boolean) As String
    If value Then
        BoolText = "1"
    Else
        BoolText = "0"
    End If
End Function

Private Function TextBool(ByVal value As String) As Boolean
    TextBool = (Trim$(value) = "1")
End Function

Private Function EscapeValue(ByVal value As String) As String
    value = Replace(value, "%", "%25")
    value = Replace(value, "|", "%7C")
    value = Replace(value, vbCrLf, "%0D%0A")
    value = Replace(value, vbCr, "%0D")
    value = Replace(value, vbLf, "%0A")
    EscapeValue = value
End Function

Private Function UnescapeValue(ByVal value As String) As String
    value = Replace(value, "%0D%0A", vbCrLf)
    value = Replace(value, "%0D", vbCr)
    value = Replace(value, "%0A", vbLf)
    value = Replace(value, "%7C", "|")
    value = Replace(value, "%25", "%")
    UnescapeValue = value
End Function

Private Function TrackCount() As Long
    TrackCount = mTrackEntries.Count
End Function

Private Function ProjectHasCdText() As Boolean
    Dim i As Long
    Dim entry As TrackEntry

    For i = 1 To mTrackEntries.Count
        Set entry = mTrackEntries.Item(i)

        If entry.EnglishText.LanguageEnabled Then
            If entry.EnglishText.TitleEnabled And Trim$(entry.EnglishText.Title) <> "" Then
                ProjectHasCdText = True
                Exit Function
            End If
            If entry.EnglishText.PerformerEnabled And Trim$(entry.EnglishText.Performer) <> "" Then
                ProjectHasCdText = True
                Exit Function
            End If
        End If

        If entry.JapaneseText.LanguageEnabled Then
            If entry.JapaneseText.TitleEnabled And Trim$(entry.JapaneseText.Title) <> "" Then
                ProjectHasCdText = True
                Exit Function
            End If
            If entry.JapaneseText.PerformerEnabled And Trim$(entry.JapaneseText.Performer) <> "" Then
                ProjectHasCdText = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Sub SelectFileIndex(ByVal index As Long)
    If index >= 0 And index < lstFiles.ListCount Then
        lstFiles.ListIndex = index
    End If
End Sub

Private Sub SelectTrackIndex(ByVal index As Long)
    If index >= 0 And index < lstTracks.ListCount Then
        lstTracks.ListIndex = index
        Set mCurrentTrack = mTrackEntries.Item(index + 1)
        txtAlbumName.Text = mCurrentTrack.DisplayTitle
    End If
End Sub

Private Sub ShowPlaceholder(ByVal messageText As String)
    lblStatus.Caption = messageText
    MsgBox messageText, vbInformation, "Kureha VB6 Rebuild"
End Sub

Private Function PadLeft(ByVal value As String, ByVal totalWidth As Long) As String
    If Len(value) >= totalWidth Then
        PadLeft = value
    Else
        PadLeft = Space$(totalWidth - Len(value)) & value
    End If
End Function

Private Function PadRight(ByVal value As String, ByVal totalWidth As Long) As String
    If Len(value) >= totalWidth Then
        PadRight = Left$(value, totalWidth)
    Else
        PadRight = value & Space$(totalWidth - Len(value))
    End If
End Function
