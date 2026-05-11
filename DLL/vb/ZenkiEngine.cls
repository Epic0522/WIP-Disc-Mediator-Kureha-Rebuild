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

Public Function AddTrackFile(ByVal fileName As String, Optional ByVal pregap As Long = 0, Optional ByVal postgap As Long = 0, Optional ByVal flags As Long = 0) As Boolean
    AddTrackFile = (Zenki_AddTrack(mHandle, fileName, pregap, postgap, flags) <> 0)
End Function

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
