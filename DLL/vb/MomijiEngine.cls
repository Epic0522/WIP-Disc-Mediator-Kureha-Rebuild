VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "MomijiEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mHandle As Long

Private Sub Class_Initialize()
    mHandle = Momiji_Initialize()
End Sub

Private Sub Class_Terminate()
    If mHandle <> 0 Then
        Momiji_Terminate mHandle
        mHandle = 0
    End If
End Sub

Public Property Get Handle() As Long
    Handle = mHandle
End Property

Public Function OpenDriveLetter(ByVal driveLetter As String) As Boolean
    Dim idx As Long
    If Len(driveLetter) = 0 Then Exit Function
    idx = Asc(UCase$(Left$(driveLetter, 1))) - Asc("A")
    OpenDriveLetter = (Momiji_Open(mHandle, idx) <> 0)
End Function

Public Function OpenDevicePath(ByVal devicePath As String) As Boolean
    OpenDevicePath = (Momiji_OpenEx(mHandle, devicePath) <> 0)
End Function

Public Sub CloseDevice()
    If mHandle <> 0 Then Momiji_Close mHandle
End Sub

Public Function IsReady() As Boolean
    IsReady = (Momiji_IsReady(mHandle) <> 0)
End Function

Public Function ReadSector2048(ByVal lba As Long, ByRef data() As Byte) As Boolean
    If UBound(data) - LBound(data) + 1 < 2048 Then Err.Raise 5, , "buffer must be at least 2048 bytes"
    ReadSector2048 = (Momiji_ReadLBA(mHandle, VarPtr(data(LBound(data))), lba) <> 0)
End Function

Public Function WriteSector2048(ByVal lba As Long, ByRef data() As Byte) As Boolean
    If UBound(data) - LBound(data) + 1 < 2048 Then Err.Raise 5, , "buffer must be at least 2048 bytes"
    WriteSector2048 = (Momiji_WriteLBA(mHandle, 2048, VarPtr(data(LBound(data))), lba) <> 0)
End Function

Public Function WriteSectorByType(ByVal lba As Long, ByRef data() As Byte, ByVal dataType As Long) As Boolean
    Dim requiredBytes As Long

    requiredBytes = dataType
    If requiredBytes <= 0 Then requiredBytes = 2048
    If UBound(data) - LBound(data) + 1 < requiredBytes Then Err.Raise 5, , "buffer is smaller than requested sector type"
    WriteSectorByType = (Momiji_WriteLBA(mHandle, dataType, VarPtr(data(LBound(data))), lba) <> 0)
End Function

Public Function WriteStart(ByVal writeCommand As Long, ByVal writeFlag As Long) As Boolean
    WriteStart = (Momiji_WriteStart(mHandle, writeCommand, writeFlag) <> 0)
End Function

Public Function WriteFlush(ByVal abortFlag As Boolean) As Boolean
    WriteFlush = (Momiji_WriteFlush(mHandle, IIf(abortFlag, 1, 0)) <> 0)
End Function

Public Function WriteEnd(ByVal abortFlag As Boolean) As Boolean
    WriteEnd = (Momiji_WriteEnd(mHandle, IIf(abortFlag, 1, 0)) <> 0)
End Function

Public Function EraseMedia(ByVal quickErase As Boolean) As Boolean
    EraseMedia = (Momiji_Erase(mHandle, IIf(quickErase, 1, 0)) <> 0)
End Function

Public Function IsDiscEmpty() As Boolean
    IsDiscEmpty = (Momiji_IsDiscEmpty(mHandle) <> 0)
End Function

Public Function SetWriteCacheBytes(ByVal bytes As Long) As Long
    SetWriteCacheBytes = Momiji_SetWriteCacheBufferSize(mHandle, bytes)
End Function

Public Function UsedWriteCacheBytes() As Long
    UsedWriteCacheBytes = Momiji_GetUsedWriteCacheBufferSize(mHandle)
End Function

Public Function LastWroteLBA() As Long
    LastWroteLBA = Momiji_GetLastWroteLBA(mHandle)
End Function

Public Function LastErrorStatus() As Long
    LastErrorStatus = Momiji_GetSCSIErrorStatus(mHandle)
End Function
