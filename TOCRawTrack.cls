VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TOCRawTrack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mSessionNo As Byte
Private mAdrControl As Byte
Private mTrack As Byte
Private mPoint As Byte
Private mMinute As Byte
Private mSecond As Byte
Private mFrame As Byte
Private mZero As Byte
Private mPMinute As Byte
Private mPSecond As Byte
Private mPFrame As Byte

Public Property Get SessionNo() As Byte
    SessionNo = mSessionNo
End Property

Public Property Let SessionNo(ByVal value As Byte)
    mSessionNo = value
End Property

Public Property Get AdrControl() As Byte
    AdrControl = mAdrControl
End Property

Public Property Let AdrControl(ByVal value As Byte)
    mAdrControl = value
End Property

Public Property Get Track() As Byte
    Track = mTrack
End Property

Public Property Let Track(ByVal value As Byte)
    mTrack = value
End Property

Public Property Get Point() As Byte
    Point = mPoint
End Property

Public Property Let Point(ByVal value As Byte)
    mPoint = value
End Property

Public Property Get Minute() As Byte
    Minute = mMinute
End Property

Public Property Let Minute(ByVal value As Byte)
    mMinute = value
End Property

Public Property Get Second() As Byte
    Second = mSecond
End Property

Public Property Let Second(ByVal value As Byte)
    mSecond = value
End Property

Public Property Get Frame() As Byte
    Frame = mFrame
End Property

Public Property Let Frame(ByVal value As Byte)
    mFrame = value
End Property

Public Property Get ZeroByte() As Byte
    ZeroByte = mZero
End Property

Public Property Let ZeroByte(ByVal value As Byte)
    mZero = value
End Property

Public Property Get PMinute() As Byte
    PMinute = mPMinute
End Property

Public Property Let PMinute(ByVal value As Byte)
    mPMinute = value
End Property

Public Property Get PSecond() As Byte
    PSecond = mPSecond
End Property

Public Property Let PSecond(ByVal value As Byte)
    mPSecond = value
End Property

Public Property Get PFrame() As Byte
    PFrame = mPFrame
End Property

Public Property Let PFrame(ByVal value As Byte)
    mPFrame = value
End Property

Public Property Get LBA() As Long
    LBA = MsfToLba(mMinute, mSecond, mFrame)
End Property

Public Property Get PLBA() As Long
    PLBA = MsfToLba(mPMinute, mPSecond, mPFrame)
End Property

Public Function Summary() As String
    Summary = "Session=" & CStr(mSessionNo) & _
        " Track=" & Format$(mTrack, "00") & _
        " Point=" & Format$(mPoint, "00") & _
        " MSF=" & Format$(mMinute, "00") & ":" & Format$(mSecond, "00") & ":" & Format$(mFrame, "00") & _
        " LBA=" & CStr(LBA)
End Function

Private Function MsfToLba(ByVal minuteValue As Byte, ByVal secondValue As Byte, ByVal frameValue As Byte) As Long
    Dim totalFrames As Long

    totalFrames = (CLng(minuteValue) * 60& * 75&) + (CLng(secondValue) * 75&) + CLng(frameValue)
    If totalFrames >= 150 Then
        MsfToLba = totalFrames - 150
    Else
        MsfToLba = totalFrames
    End If
End Function

