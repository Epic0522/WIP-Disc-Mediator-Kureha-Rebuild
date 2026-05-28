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

Private Const SIMULATED_WRITE_SECTOR_LIMIT As Long = 360000
Private Const SIMULATED_WRITE_SECTOR_BYTES As Long = 2448
Private Const SIMULATED_LEADIN_SECTORS As Long = 4500
Private Const SIMULATED_LEADOUT_SECTORS As Long = 6750
Private Const SIMULATED_WRITE_FILE_NAME As String = "kureha_filebacked_raw96_disc_test.bin"
Private Const SIMULATED_ERASE_FILE_NAME As String = "kureha_filebacked_erase_test.bin"

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
    Dim eraseWriter As MomijiEngine
    Dim sector(0 To 2447) As Byte
    Dim testPath As String
    Dim erasePath As String
    Dim leadInSectors As Long
    Dim leadOutSectors As Long
    Dim sectorIndex As Long
    Dim readSectors As Long
    Dim wroteSectors As Long
    Dim cacheBeforeFlush As Long
    Dim cacheAfterFlush As Long
    Dim lastLBA As Long
    Dim writtenBytes As Long
    Dim reachedLimit As Boolean
    Dim emptyAfterErase As Boolean
    Dim readStarted As Boolean
    Dim errText As String

    On Error GoTo Failed

    testPath = App.Path & "\" & SIMULATED_WRITE_FILE_NAME
    erasePath = App.Path & "\" & SIMULATED_ERASE_FILE_NAME

    Set writer = New MomijiEngine

    If Not writer.OpenDevicePath(testPath) Then
        Err.Raise vbObjectError + 2100, "DiscWriteForm", "Momiji.OpenEx failed for: " & testPath
    End If

    writer.SetWriteCacheBytes 1048576

    If Not writer.EraseMedia(True) Then
        Err.Raise vbObjectError + 2101, "DiscWriteForm", "Momiji.Erase failed for file-backed target."
    End If

    If Not writer.WriteStart(SIMULATED_WRITE_SECTOR_BYTES, 0) Then
        Err.Raise vbObjectError + 2102, "DiscWriteForm", "Momiji.WriteStart failed for file-backed target."
    End If

    WriteRaw96LeadIn writer, leadInSectors

    If Not mZenki.ReadStart() Then
        Err.Raise vbObjectError + 2103, "DiscWriteForm", "Zenki.ReadStart failed."
    End If
    readStarted = True

    For sectorIndex = 0 To SIMULATED_WRITE_SECTOR_LIMIT - 1
        If Not mZenki.ReadSector(sector, True) Then Exit For
        FillProgramSubQ sector, sectorIndex
        readSectors = readSectors + 1

        If Not writer.WriteSectorByType(leadInSectors + sectorIndex, sector, SIMULATED_WRITE_SECTOR_BYTES) Then
            Err.Raise vbObjectError + 2104, "DiscWriteForm", "Momiji.WriteLBA failed at sector " & CStr(sectorIndex) & "."
        End If
        wroteSectors = wroteSectors + 1
    Next
    reachedLimit = (wroteSectors >= SIMULATED_WRITE_SECTOR_LIMIT)

    If readSectors = 0 Or wroteSectors = 0 Then
        Err.Raise vbObjectError + 2105, "DiscWriteForm", "No sectors were transferred from Zenki to Momiji."
    End If

    WriteRaw96LeadOut writer, leadInSectors + wroteSectors, wroteSectors, leadOutSectors

    cacheBeforeFlush = writer.UsedWriteCacheBytes()
    lastLBA = writer.LastWroteLBA()

    If Not writer.WriteFlush(False) Then
        Err.Raise vbObjectError + 2106, "DiscWriteForm", "Momiji.WriteFlush failed."
    End If
    If Not writer.WriteEnd(False) Then
        Err.Raise vbObjectError + 2109, "DiscWriteForm", "Momiji.WriteEnd failed."
    End If
    cacheAfterFlush = writer.UsedWriteCacheBytes()

    mZenki.ReadEnd
    readStarted = False
    writer.CloseDevice
    writtenBytes = FileLen(testPath)

    Set eraseWriter = New MomijiEngine
    If Not eraseWriter.OpenDevicePath(erasePath) Then
        Err.Raise vbObjectError + 2107, "DiscWriteForm", "Momiji.OpenEx failed for erase target."
    End If
    If Not eraseWriter.EraseMedia(True) Then
        Err.Raise vbObjectError + 2108, "DiscWriteForm", "Momiji.Erase failed for separate erase target."
    End If
    emptyAfterErase = eraseWriter.IsDiscEmpty()
    eraseWriter.CloseDevice

    RunSimulatedWriteTest = "File-backed DAO RAW+SUB96 simulation completed." & vbCrLf & _
        "Write mode: DAO RAW+SUB96" & vbCrLf & _
        "Tracks synced: " & CStr(mTrackCount) & vbCrLf & _
        "Lead-in sectors: " & CStr(leadInSectors) & vbCrLf & _
        "Program sectors: " & CStr(wroteSectors) & vbCrLf & _
        "Lead-out sectors: " & CStr(leadOutSectors) & vbCrLf & _
        "Logical write range: " & CStr(-leadInSectors) & " to " & CStr(wroteSectors + leadOutSectors - 1) & vbCrLf & _
        "Sectors read: " & CStr(readSectors) & vbCrLf & _
        "Sectors written: " & CStr(leadInSectors + wroteSectors + leadOutSectors) & vbCrLf & _
        "Sector bytes: " & CStr(SIMULATED_WRITE_SECTOR_BYTES) & vbCrLf & _
        "Last wrote LBA: " & CStr(lastLBA) & vbCrLf & _
        "Written file bytes: " & CStr(writtenBytes) & vbCrLf & _
        "Reached safety limit: " & CStr(reachedLimit) & vbCrLf & _
        "Cache before flush: " & CStr(cacheBeforeFlush) & " bytes" & vbCrLf & _
        "Cache after flush: " & CStr(cacheAfterFlush) & " bytes" & vbCrLf & _
        "Separate erase target empty: " & CStr(emptyAfterErase) & vbCrLf & _
        "RAW+96 disc image: " & testPath & vbCrLf & _
        "Erase target: " & erasePath
    Exit Function

Failed:
    errText = Err.Description
    On Error Resume Next
    If readStarted Then mZenki.ReadEnd
    If Not (writer Is Nothing) Then writer.CloseDevice
    If Not (eraseWriter Is Nothing) Then eraseWriter.CloseDevice
    On Error GoTo 0
    Err.Raise vbObjectError + 2199, "DiscWriteForm", errText
End Function

Private Sub WriteRaw96LeadIn(ByVal writer As MomijiEngine, ByRef sectorsWritten As Long)
    Dim sector(0 To 2447) As Byte
    Dim cdTextPacks As String
    Dim leadFrame As Long
    Dim pointIndex As Long
    Dim pointValue As Long

    cdTextPacks = BuildCdTextPackStream()
    For leadFrame = 0 To SIMULATED_LEADIN_SECTORS - 1
        ClearBytes sector
        pointIndex = leadFrame Mod (mTrackCount + 3)
        Select Case pointIndex
            Case 0
                pointValue = &HA0
            Case 1
                pointValue = &HA1
            Case 2
                pointValue = &HA2
            Case Else
                pointValue = pointIndex - 2
        End Select
        FillLeadInSubQ sector, pointValue, leadFrame
        FillCdTextRW sector, leadFrame, cdTextPacks

        If Not writer.WriteSectorByType(leadFrame, sector, SIMULATED_WRITE_SECTOR_BYTES) Then
            Err.Raise vbObjectError + 2094, "DiscWriteForm", "Momiji.WriteLBA failed while writing RAW+96 lead-in sector " & CStr(leadFrame) & "."
        End If
        sectorsWritten = sectorsWritten + 1
    Next leadFrame
End Sub

Private Sub WriteRaw96LeadOut(ByVal writer As MomijiEngine, ByVal startFileSector As Long, ByVal programSectors As Long, ByRef sectorsWritten As Long)
    Dim sector(0 To 2447) As Byte
    Dim leadOutFrame As Long

    For leadOutFrame = 0 To SIMULATED_LEADOUT_SECTORS - 1
        ClearBytes sector
        FillLeadOutSubQ sector, programSectors, leadOutFrame
        If Not writer.WriteSectorByType(startFileSector + leadOutFrame, sector, SIMULATED_WRITE_SECTOR_BYTES) Then
            Err.Raise vbObjectError + 2095, "DiscWriteForm", "Momiji.WriteLBA failed while writing RAW+96 lead-out sector " & CStr(leadOutFrame) & "."
        End If
        sectorsWritten = sectorsWritten + 1
    Next leadOutFrame
End Sub

Private Function BuildCdTextPackStream() As String
    Dim trackNo As Long
    Dim sequenceNo As Long
    Dim stream As String

    If chkWriteCDText.Value = 0 Then
        BuildCdTextPackStream = ""
        Exit Function
    End If

    sequenceNo = 0
    For trackNo = 1 To mTrackCount
        AppendCdTextField stream, sequenceNo, &H80, trackNo, 0, mZenki.GetText(trackNo, 0, 0)
        AppendCdTextField stream, sequenceNo, &H81, trackNo, 0, mZenki.GetText(trackNo, 0, 1)
        AppendCdTextField stream, sequenceNo, &H80, trackNo, 1, mZenki.GetText(trackNo, 1, 0)
        AppendCdTextField stream, sequenceNo, &H81, trackNo, 1, mZenki.GetText(trackNo, 1, 1)
    Next trackNo

    BuildCdTextPackStream = stream
End Function

Private Sub AppendCdTextField(ByRef stream As String, ByRef sequenceNo As Long, ByVal packType As Long, ByVal trackNo As Long, ByVal blockNo As Long, ByVal value As String)
    Dim rawText As String
    Dim pos As Long
    Dim partNo As Long
    Dim payload As String

    If Trim$(value) = "" Then Exit Sub
    If Not IsShiftJisRoundTripSafe(value) Then
        Err.Raise vbObjectError + 2097, "DiscWriteForm", "CD-TEXT contains characters that cannot round-trip through Japanese Shift-JIS/CP932 bytes: " & value
    End If

    rawText = StrConv(value & vbNullChar, vbFromUnicode)
    pos = 1
    partNo = 0
    Do While pos <= LenB(rawText)
        payload = MidB$(rawText, pos, 12)
        stream = stream & BuildCdTextPack(packType, trackNo, sequenceNo, blockNo, partNo, payload)
        sequenceNo = (sequenceNo + 1) And &HFF
        partNo = (partNo + 1) And &HF
        pos = pos + 12
    Loop
End Sub

Private Function IsShiftJisRoundTripSafe(ByVal value As String) As Boolean
    Dim rawText As String
    Dim roundTrip As String

    rawText = StrConv(value, vbFromUnicode)
    roundTrip = StrConv(rawText, vbUnicode)
    IsShiftJisRoundTripSafe = (roundTrip = value)
End Function

Private Function BuildCdTextPack(ByVal packType As Long, ByVal trackNo As Long, ByVal sequenceNo As Long, ByVal blockNo As Long, ByVal partNo As Long, ByVal payload As String) As String
    Dim pack(0 To 17) As Byte
    Dim i As Long
    Dim maxBytes As Long

    pack(0) = packType And &HFF
    pack(1) = trackNo And &HFF
    pack(2) = sequenceNo And &HFF
    pack(3) = ((blockNo And &H7) * 16) Or (partNo And &HF)

    maxBytes = LenB(payload)
    If maxBytes > 12 Then maxBytes = 12
    For i = 1 To maxBytes
        pack(3 + i) = AscB(MidB$(payload, i, 1))
    Next i

    UpdateCdTextPackCRC pack
    BuildCdTextPack = BytesToRawString(pack)
End Function

Private Sub FillCdTextRW(ByRef sector() As Byte, ByVal leadFrame As Long, ByVal cdTextPacks As String)
    Dim packCount As Long
    Dim packIndex As Long
    Dim slotNo As Long
    Dim rawPack As String

    packCount = LenB(cdTextPacks) \ 18
    If packCount <= 0 Then Exit Sub

    For slotNo = 0 To 3
        packIndex = ((leadFrame * 4) + slotNo) Mod packCount
        rawPack = MidB$(cdTextPacks, packIndex * 18 + 1, 18)
        CopyRawStringToBytes rawPack, sector, 2352 + 24 + slotNo * 18
    Next slotNo
End Sub

Private Sub FillLeadInSubQ(ByRef sector() As Byte, ByVal pointValue As Long, ByVal leadFrame As Long)
    Dim q(0 To 11) As Byte
    Dim pFrame As Long

    q(0) = &H1
    q(1) = 0
    If pointValue >= &HA0 Then
        q(2) = pointValue
    Else
        q(2) = ByteToBcd(pointValue)
    End If
    PutMsf q, 3, leadFrame
    q(6) = 0

    Select Case pointValue
        Case &HA0
            q(7) = ByteToBcd(1)
            q(8) = &H0
            q(9) = &H0
        Case &HA1
            q(7) = ByteToBcd(mTrackCount)
            q(8) = &H0
            q(9) = &H0
        Case &HA2
            PutMsf q, 7, TotalProgramFrames() + 150
        Case Else
            pFrame = TrackStartFrame(pointValue) + 150
            PutMsf q, 7, pFrame
    End Select

    PutSubQ sector, q
End Sub

Private Sub FillProgramSubQ(ByRef sector() As Byte, ByVal programFrame As Long)
    Dim q(0 To 11) As Byte
    Dim trackNo As Long
    Dim relativeFrame As Long

    trackNo = TrackForProgramFrame(programFrame)
    relativeFrame = programFrame - TrackStartFrame(trackNo)
    If relativeFrame < 0 Then relativeFrame = 0

    q(0) = &H1
    q(1) = ByteToBcd(trackNo)
    q(2) = &H1
    PutMsf q, 3, relativeFrame
    q(6) = 0
    PutMsf q, 7, programFrame + 150
    PutSubQ sector, q
End Sub

Private Sub FillLeadOutSubQ(ByRef sector() As Byte, ByVal programSectors As Long, ByVal leadOutFrame As Long)
    Dim q(0 To 11) As Byte

    q(0) = &H1
    q(1) = &HAA
    q(2) = &H1
    PutMsf q, 3, leadOutFrame
    q(6) = 0
    PutMsf q, 7, programSectors + leadOutFrame + 150
    PutSubQ sector, q
End Sub

Private Sub PutSubQ(ByRef sector() As Byte, ByRef q() As Byte)
    Dim i As Long

    UpdateSubQCRC q
    For i = 0 To 11
        sector(2352 + 12 + i) = q(i)
    Next i
End Sub

Private Sub PutMsf(ByRef q() As Byte, ByVal offset As Long, ByVal frameNo As Long)
    Dim minuteValue As Long
    Dim secondValue As Long
    Dim frameValue As Long

    If frameNo < 0 Then frameNo = 0
    minuteValue = frameNo \ 4500
    secondValue = (frameNo Mod 4500) \ 75
    frameValue = frameNo Mod 75
    q(offset) = ByteToBcd(minuteValue)
    q(offset + 1) = ByteToBcd(secondValue)
    q(offset + 2) = ByteToBcd(frameValue)
End Sub

Private Function ByteToBcd(ByVal value As Long) As Byte
    If value < 0 Then value = 0
    If value > 99 Then value = 99
    ByteToBcd = ((value \ 10) * 16) Or (value Mod 10)
End Function

Private Sub UpdateSubQCRC(ByRef q() As Byte)
    Dim crc As Long
    Dim i As Long
    Dim bitNo As Long

    crc = 0
    For i = 0 To 9
        crc = crc Xor (CLng(q(i)) * 256)
        For bitNo = 0 To 7
            If (crc And &H8000&) <> 0 Then
                crc = ((crc * 2) Xor &H1021&) And &HFFFF&
            Else
                crc = (crc * 2) And &HFFFF&
            End If
        Next bitNo
    Next i

    crc = crc Xor &HFFFF&
    q(10) = (crc \ 256) And &HFF
    q(11) = crc And &HFF
End Sub

Private Sub UpdateCdTextPackCRC(ByRef pack() As Byte)
    Dim crc As Long
    Dim i As Long
    Dim bitNo As Long

    crc = 0
    For i = 0 To 15
        crc = crc Xor (CLng(pack(i)) * 256)
        For bitNo = 0 To 7
            If (crc And &H8000&) <> 0 Then
                crc = ((crc * 2) Xor &H1021&) And &HFFFF&
            Else
                crc = (crc * 2) And &HFFFF&
            End If
        Next bitNo
    Next i

    crc = crc Xor &HFFFF&
    pack(16) = (crc \ 256) And &HFF
    pack(17) = crc And &HFF
End Sub

Private Function TotalProgramFrames() As Long
    Dim trackNo As Long

    For trackNo = 1 To mTrackCount
        TotalProgramFrames = TotalProgramFrames + TrackLengthFrames(trackNo)
    Next trackNo
End Function

Private Function TrackStartFrame(ByVal trackNo As Long) As Long
    Dim i As Long

    If trackNo < 1 Then trackNo = 1
    For i = 1 To trackNo - 1
        TrackStartFrame = TrackStartFrame + TrackLengthFrames(i)
    Next i
End Function

Private Function TrackForProgramFrame(ByVal programFrame As Long) As Long
    Dim trackNo As Long
    Dim nextStart As Long

    TrackForProgramFrame = 1
    For trackNo = 1 To mTrackCount
        nextStart = TrackStartFrame(trackNo) + TrackLengthFrames(trackNo)
        If programFrame < nextStart Then
            TrackForProgramFrame = trackNo
            Exit Function
        End If
    Next trackNo
    If mTrackCount > 0 Then TrackForProgramFrame = mTrackCount
End Function

Private Function TrackLengthFrames(ByVal trackNo As Long) As Long
    TrackLengthFrames = mZenki.GetTrackPregapFrames(trackNo) + mZenki.GetTrackSectorCount(trackNo) + mZenki.GetTrackPostgapFrames(trackNo)
    If TrackLengthFrames < 1 Then TrackLengthFrames = 1
End Function

Private Sub ClearBytes(ByRef data() As Byte)
    Dim i As Long

    For i = LBound(data) To UBound(data)
        data(i) = 0
    Next i
End Sub

Private Sub CopyTextToBytes(ByVal textValue As String, ByRef data() As Byte)
    Dim rawText As String

    rawText = StrConv(textValue, vbFromUnicode)
    CopyRawBytes rawText, data
End Sub

Private Sub CopyRawBytes(ByVal rawText As String, ByRef data() As Byte)
    Dim i As Long
    Dim maxBytes As Long

    maxBytes = LenB(rawText)
    If maxBytes > UBound(data) - LBound(data) + 1 Then maxBytes = UBound(data) - LBound(data) + 1

    For i = 1 To maxBytes
        data(LBound(data) + i - 1) = AscB(MidB$(rawText, i, 1))
    Next i
End Sub

Private Sub CopyRawStringToBytes(ByVal rawText As String, ByRef data() As Byte, ByVal offset As Long)
    Dim i As Long
    Dim maxBytes As Long

    maxBytes = LenB(rawText)
    If offset + maxBytes > UBound(data) + 1 Then maxBytes = UBound(data) - offset + 1
    If maxBytes <= 0 Then Exit Sub

    For i = 1 To maxBytes
        data(offset + i - 1) = AscB(MidB$(rawText, i, 1))
    Next i
End Sub

Private Function BytesToRawString(ByRef data() As Byte) As String
    Dim output As String
    Dim i As Long

    For i = LBound(data) To UBound(data)
        output = output & ChrB$(data(i))
    Next i
    BytesToRawString = output
End Function

