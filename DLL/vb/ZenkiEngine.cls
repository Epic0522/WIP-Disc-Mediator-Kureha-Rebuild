VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "ZenkiEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mHandle As Long

Private Sub Class_Initialize()
    mHandle = Zenki_Initialize()
End Sub

Private Sub Class_Terminate()
    If mHandle <> 0 Then
        Zenki_Terminate mHandle
        mHandle = 0
    End If
End Sub

Public Property Get Handle() As Long
    Handle = mHandle
End Property

Public Function InitIsoFileSystem(ByVal isoFlags As Long) As Boolean
    InitIsoFileSystem = (Zenki_InitISOFS(mHandle, isoFlags) <> 0)
End Function

Public Sub ClearIso()
    If mHandle <> 0 Then Zenki_ClearISO mHandle
End Sub

Public Function IsIsoEmpty() As Boolean
    IsIsoEmpty = (Zenki_IsISOEmpty(mHandle) <> 0)
End Function

Public Function AddIsoFile(ByVal isoName As String, ByVal originalFileName As String) As Boolean
    AddIsoFile = (Zenki_AddISOFile(mHandle, isoName, originalFileName) <> 0)
End Function

Public Function AddIsoDummyFile(ByVal isoName As String, ByVal byteSize As Long, Optional ByVal fileAttribute As Long = 0, Optional ByVal timestampValue As Long = 0) As Boolean
    AddIsoDummyFile = (Zenki_AddISODummyFile(mHandle, isoName, byteSize, fileAttribute, timestampValue) <> 0)
End Function

Public Function MakeIsoDirectory(ByVal isoName As String, Optional ByVal fileAttribute As Long = 0, Optional ByVal timestampValue As Long = 0) As Boolean
    MakeIsoDirectory = (Zenki_MakeISODirectory(mHandle, isoName, fileAttribute, timestampValue) <> 0)
End Function

Public Function RemoveIsoFile(ByVal isoName As String) As Boolean
    RemoveIsoFile = (Zenki_RemoveISOFile(mHandle, isoName) <> 0)
End Function

Public Function RenameIsoFile(ByVal oldIsoName As String, ByVal newIsoName As String) As Boolean
    RenameIsoFile = (Zenki_RenameISOFile(mHandle, oldIsoName, newIsoName) <> 0)
End Function

Public Function ChangeIsoProperties(ByVal isoName As String, ByVal fileAttribute As Long, ByVal timestampValue As Long) As Boolean
    ChangeIsoProperties = (Zenki_ChangeISOProperties(mHandle, isoName, fileAttribute, timestampValue) <> 0)
End Function

Public Function IsoCurrentDirectory() As String
    Dim buf As String
    buf = String$(260, vbNullChar)
    Zenki_GetISOCurrentDirectory mHandle, buf, Len(buf)
    IsoCurrentDirectory = CStringFromBuffer(buf)
End Function

Public Function AddTrackFile(ByVal fileName As String, Optional ByVal pregap As Long = 0, Optional ByVal postgap As Long = 0, Optional ByVal flags As Long = 0) As Boolean
    AddTrackFile = (Zenki_AddTrack(mHandle, fileName, pregap, postgap, flags) <> 0)
End Function

Public Sub ClearTracks()
    If mHandle <> 0 Then Zenki_ClearTrack mHandle
End Sub

Public Function TrackCount() As Long
    TrackCount = Zenki_GetTrackCount(mHandle)
End Function

Public Sub EnableTrackText(ByVal enabled As Boolean)
    Zenki_EnabledTrackText mHandle, IIf(enabled, 1, 0)
End Sub

Public Sub SetText(ByVal trackNo As Long, ByVal languageNo As Long, ByVal informationNo As Long, ByVal value As String)
    Zenki_SetTrackText mHandle, trackNo, languageNo, informationNo, value
End Sub

Public Function GetText(ByVal trackNo As Long, ByVal languageNo As Long, ByVal informationNo As Long) As String
    Dim buf As String
    buf = String$(1024, vbNullChar)
    Zenki_GetTrackText mHandle, trackNo, languageNo, informationNo, buf, Len(buf)
    GetText = CStringFromBuffer(buf)
End Function

Public Function GetTrackPath(ByVal trackNo As Long) As String
    Dim infoBytes() As Byte

    infoBytes = ReadTrackInfoBytes(trackNo)
    GetTrackPath = CStringFromBytes(infoBytes, &H400)
End Function

Public Function GetTrackPathHex(ByVal trackNo As Long, Optional ByVal maxBytes As Long = 64) As String
    Dim infoBytes() As Byte

    infoBytes = ReadTrackInfoBytes(trackNo)
    GetTrackPathHex = HexFromBytes(infoBytes, maxBytes)
End Function

Public Function GetTrackPregapFrames(ByVal trackNo As Long) As Long
    Dim infoBytes() As Byte

    infoBytes = ReadTrackInfoBytes(trackNo)
    GetTrackPregapFrames = ReadLongFromBytes(infoBytes, &H400)
End Function

Public Function GetTrackPostgapFrames(ByVal trackNo As Long) As Long
    Dim infoBytes() As Byte

    infoBytes = ReadTrackInfoBytes(trackNo)
    GetTrackPostgapFrames = ReadLongFromBytes(infoBytes, &H404)
End Function

Public Function GetTrackControlFlag(ByVal trackNo As Long) As Long
    Dim infoBytes() As Byte

    infoBytes = ReadTrackInfoBytes(trackNo)
    GetTrackControlFlag = ReadLongFromBytes(infoBytes, &H408)
End Function

Public Function GetTrackSectorCount(ByVal trackNo As Long) As Long
    Dim infoBytes() As Byte

    infoBytes = ReadTrackInfoBytes(trackNo)
    GetTrackSectorCount = ReadLongFromBytes(infoBytes, &H40C)
End Function

Public Function ReadStart() As Boolean
    ReadStart = (Zenki_ReadStart(mHandle, vbNullString) = 0)
End Function

Public Function ReadSector(ByRef data() As Byte, ByVal rawMode As Boolean) As Boolean
    If UBound(data) - LBound(data) + 1 < 2448 Then Err.Raise 5, , "buffer must be at least 2448 bytes"
    ReadSector = (Zenki_Read(mHandle, VarPtr(data(LBound(data))), IIf(rawMode, 1, 0)) <> 0)
End Function

Public Sub ReadEnd()
    Zenki_ReadEnd mHandle
End Sub

Private Function ReadTrackInfoBytes(ByVal trackNo As Long) As Byte()
    Dim buffer(0 To &H40F) As Byte

    Zenki_GetTrackInformation mHandle, trackNo, VarPtr(buffer(0))
    ReadTrackInfoBytes = buffer
End Function

Private Function ReadLongFromBytes(ByRef value() As Byte, ByVal offset As Long) As Long
    Dim result As Long

    If offset < LBound(value) Or offset + 3 > UBound(value) Then Exit Function

    CopyMemory VarPtr(result), VarPtr(value(offset)), 4
    ReadLongFromBytes = result
End Function

Private Function CStringFromBytes(ByRef value() As Byte, ByVal maxBytes As Long) As String
    Dim i As Long
    Dim limitValue As Long
    Dim byteCount As Long
    Dim rawBytes() As Byte

    If maxBytes <= 0 Then Exit Function

    limitValue = maxBytes - 1
    If limitValue > UBound(value) Then limitValue = UBound(value)

    For i = 0 To limitValue
        If value(i) = 0 Then Exit For
        byteCount = byteCount + 1
    Next i

    If byteCount = 0 Then
        CStringFromBytes = ""
        Exit Function
    End If

    ReDim rawBytes(0 To byteCount - 1) As Byte
    CopyMemory VarPtr(rawBytes(0)), VarPtr(value(0)), byteCount
    CStringFromBytes = StrConv(rawBytes, vbUnicode)
End Function

Private Function HexFromBytes(ByRef value() As Byte, ByVal maxBytes As Long) As String
    Dim i As Long
    Dim lastIndex As Long
    Dim pieces() As String

    If maxBytes <= 0 Then Exit Function

    lastIndex = maxBytes - 1
    If lastIndex > UBound(value) Then lastIndex = UBound(value)
    If lastIndex < 0 Then Exit Function

    ReDim pieces(0 To lastIndex) As String
    For i = 0 To lastIndex
        pieces(i) = Right$("0" & Hex$(value(i)), 2)
    Next i

    HexFromBytes = Join(pieces, " ")
End Function
