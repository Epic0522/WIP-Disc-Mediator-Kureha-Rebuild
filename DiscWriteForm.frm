VERSION 5.00
Begin VB.Form DiscWriteForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Disc Write"
   ClientHeight    =   3180
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5925
   LinkTopic       =   "DiscWriteForm"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdStartWrite
      Caption         =   "Start"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdParameters
      Caption         =   "Parameters"
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose
      Caption         =   "Close"
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CheckBox chkWriteCDText
      Caption         =   "Write CD-TEXT when available"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   2775
   End
   Begin VB.ComboBox cboWriteMode
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1680
      Width           =   1995
   End
   Begin VB.Label lblWriteMode
      Caption         =   "Write mode:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1740
      Width           =   1155
   End
   Begin VB.Label lblSummary
      Caption         =   "No project loaded."
      Height          =   1035
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   5415
   End
   Begin VB.Label lblHeader
      Caption         =   "Disc write preview"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   1815
   End
End
Attribute VB_Name = "DiscWriteForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SIMULATED_WRITE_SECTOR_LIMIT As Long = 32
Private Const SIMULATED_WRITE_FILE_NAME As String = "kureha_filebacked_write_test.bin"

Private mDiscLabel As String
Private mMediaType As String
Private mTrackCount As Long
Private mHasCdText As Boolean
Private mLastTestLine As String
Private mZenki As ZenkiEngine

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdParameters_Click()
    PropertyWriteParameterForm.Show vbModal, Me
End Sub

Private Sub cmdStartWrite_Click()
    Dim report As String

    On Error GoTo Failed

    If mTrackCount <= 0 Then
        MsgBox "Add at least one track before running the simulated write test.", vbExclamation, "Kureha VB6 Rebuild"
        Exit Sub
    End If

    If mZenki Is Nothing Then
        MsgBox "Zenki engine is not available. Reopen the main window and try again.", vbExclamation, "Kureha VB6 Rebuild"
        Exit Sub
    End If

    If mZenki.Handle = 0 Then
        MsgBox "Zenki engine is not initialized.", vbExclamation, "Kureha VB6 Rebuild"
        Exit Sub
    End If

    EnsureNativeDllPath
    report = RunSimulatedWriteTest()
    mLastTestLine = "Last test: simulated write passed."
    UpdatePreviewText
    MsgBox report, vbInformation, "Kureha VB6 Rebuild"
    Exit Sub

Failed:
    mLastTestLine = "Last test: simulated write failed."
    UpdatePreviewText
    MsgBox "Simulated write failed." & vbCrLf & vbCrLf & Err.Description, vbExclamation, "Kureha VB6 Rebuild"
End Sub

Private Sub Form_Load()
    cboWriteMode.AddItem "SAO"
    cboWriteMode.AddItem "SAO RAW"
    cboWriteMode.AddItem "DAO RAW+96"
    cboWriteMode.ListIndex = 0
    cmdStartWrite.Caption = "Test Write"
End Sub

Public Sub LoadPreview(ByVal discLabel As String, ByVal mediaType As String, ByVal trackCount As Long, ByVal hasCdText As Boolean, ByVal zenkiEngine As ZenkiEngine)
    mDiscLabel = discLabel
    mMediaType = mediaType
    mTrackCount = trackCount
    mHasCdText = hasCdText
    mLastTestLine = "Start runs a file-backed simulated write. No real drive access."
    Set mZenki = zenkiEngine
    chkWriteCDText.Value = Abs(hasCdText)
    UpdatePreviewText
End Sub

Private Sub UpdatePreviewText()
    lblSummary.Caption = "Label: " & mDiscLabel & vbCrLf & _
        "Media: " & mMediaType & vbCrLf & _
        "Tracks: " & CStr(mTrackCount) & vbCrLf & _
        "CD-TEXT: " & IIf(mHasCdText, "enabled", "not set") & vbCrLf & _
        mLastTestLine
End Sub

Private Function RunSimulatedWriteTest() As String
    Dim writer As MomijiEngine
    Dim sector(0 To 2447) As Byte
    Dim testPath As String
    Dim sectorIndex As Long
    Dim readSectors As Long
    Dim wroteSectors As Long
    Dim cacheBeforeFlush As Long
    Dim cacheAfterFlush As Long
    Dim lastLBA As Long
    Dim emptyAfterErase As Boolean
    Dim readStarted As Boolean
    Dim errText As String

    On Error GoTo Failed

    testPath = App.Path & "\" & SIMULATED_WRITE_FILE_NAME
    Set writer = New MomijiEngine

    If Not writer.OpenDevicePath(testPath) Then
        Err.Raise vbObjectError + 2100, "DiscWriteForm", "Momiji.OpenEx failed for: " & testPath
    End If

    writer.SetWriteCacheBytes 1048576

    If Not writer.EraseMedia(True) Then
        Err.Raise vbObjectError + 2101, "DiscWriteForm", "Momiji.Erase failed for file-backed target."
    End If

    If Not writer.WriteStart(2048, 0) Then
        Err.Raise vbObjectError + 2102, "DiscWriteForm", "Momiji.WriteStart failed for file-backed target."
    End If

    If Not mZenki.ReadStart() Then
        Err.Raise vbObjectError + 2103, "DiscWriteForm", "Zenki.ReadStart failed."
    End If
    readStarted = True

    For sectorIndex = 0 To SIMULATED_WRITE_SECTOR_LIMIT - 1
        If Not mZenki.ReadSector(sector, False) Then Exit For
        readSectors = readSectors + 1

        If Not writer.WriteSector2048(sectorIndex, sector) Then
            Err.Raise vbObjectError + 2104, "DiscWriteForm", "Momiji.WriteLBA failed at sector " & CStr(sectorIndex) & "."
        End If
        wroteSectors = wroteSectors + 1
    Next

    If readSectors = 0 Or wroteSectors = 0 Then
        Err.Raise vbObjectError + 2105, "DiscWriteForm", "No sectors were transferred from Zenki to Momiji."
    End If

    cacheBeforeFlush = writer.UsedWriteCacheBytes()
    lastLBA = writer.LastWroteLBA()

    If Not writer.WriteFlush(False) Then
        Err.Raise vbObjectError + 2106, "DiscWriteForm", "Momiji.WriteFlush failed."
    End If
    cacheAfterFlush = writer.UsedWriteCacheBytes()

    If Not writer.EraseMedia(True) Then
        Err.Raise vbObjectError + 2107, "DiscWriteForm", "Momiji.Erase after test failed."
    End If
    emptyAfterErase = writer.IsDiscEmpty()

    mZenki.ReadEnd
    readStarted = False
    writer.CloseDevice

    RunSimulatedWriteTest = "File-backed simulated write completed." & vbCrLf & _
        "Tracks synced: " & CStr(mTrackCount) & vbCrLf & _
        "Sectors read: " & CStr(readSectors) & vbCrLf & _
        "Sectors written: " & CStr(wroteSectors) & vbCrLf & _
        "Last wrote LBA: " & CStr(lastLBA) & vbCrLf & _
        "Cache before flush: " & CStr(cacheBeforeFlush) & " bytes" & vbCrLf & _
        "Cache after flush: " & CStr(cacheAfterFlush) & " bytes" & vbCrLf & _
        "Empty after erase: " & CStr(emptyAfterErase) & vbCrLf & _
        "Target: " & testPath
    Exit Function

Failed:
    errText = Err.Description
    On Error Resume Next
    If readStarted Then mZenki.ReadEnd
    If Not (writer Is Nothing) Then writer.CloseDevice
    On Error GoTo 0
    Err.Raise vbObjectError + 2199, "DiscWriteForm", errText
End Function

