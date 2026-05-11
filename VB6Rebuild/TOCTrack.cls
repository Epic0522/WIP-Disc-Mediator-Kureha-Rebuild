VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TOCTrack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mSessionNo As Byte
Private mTrackNo As Byte
Private mIndexNo As Byte
Private mFlags As Long
Private mSectorMode As Long
Private mStartLBA As Long
Private mEndLBA As Long

Public Property Get SessionNo() As Byte
    SessionNo = mSessionNo
End Property

Public Property Let SessionNo(ByVal value As Byte)
    mSessionNo = value
End Property

Public Property Get TrackNo() As Byte
    TrackNo = mTrackNo
End Property

Public Property Let TrackNo(ByVal value As Byte)
    mTrackNo = value
End Property

Public Property Get IndexNo() As Byte
    IndexNo = mIndexNo
End Property

Public Property Let IndexNo(ByVal value As Byte)
    mIndexNo = value
End Property

Public Property Get Flags() As Long
    Flags = mFlags
End Property

Public Property Let Flags(ByVal value As Long)
    mFlags = value
End Property

Public Property Get SectorMode() As Long
    SectorMode = mSectorMode
End Property

Public Property Let SectorMode(ByVal value As Long)
    mSectorMode = value
End Property

Public Property Get StartLBA() As Long
    StartLBA = mStartLBA
End Property

Public Property Let StartLBA(ByVal value As Long)
    mStartLBA = value
End Property

Public Property Get EndLBA() As Long
    EndLBA = mEndLBA
End Property

Public Property Let EndLBA(ByVal value As Long)
    mEndLBA = value
End Property

Public Property Get StartMSF() As String
    StartMSF = LbaToMSF(mStartLBA, False)
End Property

Public Property Get AbsoluteStartMSF() As String
    AbsoluteStartMSF = LbaToMSF(mStartLBA, True)
End Property

Public Property Get EndMSF() As String
    EndMSF = LbaToMSF(mEndLBA, False)
End Property

Public Function Summary() As String
    Summary = "S" & CStr(mSessionNo) & " T" & Format$(mTrackNo, "00") & _
        " I" & Format$(mIndexNo, "00") & _
        " Start=" & StartMSF & _
        " End=" & EndMSF & _
        " Flags=" & CStr(mFlags) & _
        " Mode=" & CStr(mSectorMode)
End Function

Private Function LbaToMSF(ByVal lba As Long, ByVal includePregap As Boolean) As String
    Dim adjusted As Long
    Dim minuteValue As Long
    Dim secondValue As Long
    Dim frameValue As Long

    adjusted = lba
    If includePregap Then
        adjusted = adjusted + 150
    End If

    minuteValue = adjusted \ (75 * 60)
    secondValue = (adjusted Mod (75 * 60)) \ 75
    frameValue = adjusted Mod 75

    LbaToMSF = Format$(minuteValue, "00") & ":" & _
        Format$(secondValue, "00") & ":" & _
        Format$(frameValue, "00")
End Function

