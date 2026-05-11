VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TOCInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mTracks As Collection
Private mRawTracks As Collection
Private mCDTexts As Collection
Private mISRCs As Collection
Private mMCN As String

Private Sub Class_Initialize()
    ResetCollections
End Sub

Public Sub Clear()
    mMCN = ""
    ResetCollections
End Sub

Public Property Get MCN() As String
    MCN = mMCN
End Property

Public Property Let MCN(ByVal value As String)
    mMCN = Left$(value & String$(13, " "), 13)
End Property

Public Property Get TOCCount() As Long
    TOCCount = mTracks.Count
End Property

Public Property Get TOC(ByVal index As Long) As TOCTrack
    Set TOC = mTracks.Item(index)
End Property

Public Sub TOCAdd(ByVal sessionNo As Byte, ByVal trackNo As Byte, ByVal indexNo As Byte, ByVal flags As Long, ByVal sectorMode As Long, ByVal startLBA As Long, ByVal endLBA As Long)
    Dim item As TOCTrack

    Set item = New TOCTrack
    item.SessionNo = sessionNo
    item.TrackNo = trackNo
    item.IndexNo = indexNo
    item.Flags = flags
    item.SectorMode = sectorMode
    item.StartLBA = startLBA
    item.EndLBA = endLBA
    mTracks.Add item
End Sub

Public Sub TOCDelete(ByVal index As Long)
    mTracks.Remove index
End Sub

Public Property Get RawTOCCount() As Long
    RawTOCCount = mRawTracks.Count
End Property

Public Property Get RawTOC(ByVal index As Long) As TOCRawTrack
    Set RawTOC = mRawTracks.Item(index)
End Property

Public Sub RawTOCAdd(ByVal sessionNo As Byte, ByVal adrControl As Byte, ByVal track As Byte, ByVal point As Byte, ByVal minuteValue As Byte, ByVal secondValue As Byte, ByVal frameValue As Byte, ByVal zeroByte As Byte, ByVal pMinuteValue As Byte, ByVal pSecondValue As Byte, ByVal pFrameValue As Byte)
    Dim item As TOCRawTrack

    Set item = New TOCRawTrack
    item.SessionNo = sessionNo
    item.AdrControl = adrControl
    item.Track = track
    item.Point = point
    item.Minute = minuteValue
    item.Second = secondValue
    item.Frame = frameValue
    item.ZeroByte = zeroByte
    item.PMinute = pMinuteValue
    item.PSecond = pSecondValue
    item.PFrame = pFrameValue
    mRawTracks.Add item
End Sub

Public Sub RawTOCDelete(ByVal index As Long)
    mRawTracks.Remove index
End Sub

Public Property Get ISRCCount() As Long
    ISRCCount = mISRCs.Count
End Property

Public Property Get ISRC(ByVal index As Long) As String
    ISRC = mISRCs.Item(index)
End Property

Public Sub ISRCAdd(ByVal value As String)
    mISRCs.Add Left$(value & String$(12, " "), 12)
End Sub

Public Sub ISRCDelete(ByVal index As Long)
    mISRCs.Remove index
End Sub

Public Property Get CDTextCount() As Long
    CDTextCount = mCDTexts.Count
End Property

Public Property Get CDText(ByVal index As Long) As TOCCDText
    Set CDText = mCDTexts.Item(index)
End Property

Public Sub CDTextAdd(ByVal item As TOCCDText)
    mCDTexts.Add item
End Sub

Public Sub CDTextDelete(ByVal index As Long)
    mCDTexts.Remove index
End Sub

Public Function Summary() As String
    Dim output As String
    Dim i As Long

    output = "MCN: " & Trim$(mMCN) & vbCrLf
    output = output & "Tracks: " & CStr(mTracks.Count) & vbCrLf
    For i = 1 To mTracks.Count
        output = output & "  - " & mTracks.Item(i).Summary & vbCrLf
    Next i

    output = output & "Raw TOC entries: " & CStr(mRawTracks.Count) & vbCrLf
    For i = 1 To mRawTracks.Count
        output = output & "  - " & mRawTracks.Item(i).Summary & vbCrLf
    Next i

    output = output & "CD-TEXT blocks: " & CStr(mCDTexts.Count) & vbCrLf
    For i = 1 To mCDTexts.Count
        output = output & "  - " & mCDTexts.Item(i).Summary & vbCrLf
    Next i

    Summary = output
End Function

Private Sub ResetCollections()
    Set mTracks = New Collection
    Set mRawTracks = New Collection
    Set mCDTexts = New Collection
    Set mISRCs = New Collection
End Sub

