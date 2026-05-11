Attribute VB_Name = "modMomijiZenkiDeclares"
Option Explicit

' Clean, prefixed declares reconstructed from the decompiled VB6 wrapper classes
' and from the native DLL export tables. These names avoid conflicts between
' Momiji.dll and Zenki.dll functions such as Initialize, Terminate, Read, ReadEnd.
'
' All functions are 32-bit stdcall native DLL calls. Use this in VB6 x86 only.

' ===== MOMIJI.DLL =====
Public Declare Function Momiji_GetEngineVersion Lib "Momiji.dll" Alias "GetEngineVersion" () As Long
Public Declare Function Momiji_Initialize Lib "Momiji.dll" Alias "Initialize" () As Long
Public Declare Sub Momiji_Terminate Lib "Momiji.dll" Alias "Terminate" (ByVal hEngine As Long)

Public Declare Function Momiji_GetRemoteHosts Lib "Momiji.dll" Alias "GetRemoteHosts" ( _
    ByVal hEngine As Long, _
    ByVal pBroadcastIP As Long, _
    ByVal portNo As Long, _
    ByVal bufferSize As Long, _
    ByVal timeoutMs As Long, _
    ByVal pOutBuffer As Long) As Long

Public Declare Function Momiji_Open Lib "Momiji.dll" Alias "Open" (ByVal hEngine As Long, ByVal driveIndex As Long) As Long
Public Declare Function Momiji_OpenEx Lib "Momiji.dll" Alias "OpenEx" (ByVal hEngine As Long, ByVal pDeviceName As Long) As Long
Public Declare Function Momiji_Close Lib "Momiji.dll" Alias "Close" (ByVal hEngine As Long) As Long

Public Declare Function Momiji_GetDeviceName Lib "Momiji.dll" Alias "GetDeviceName" ( _
    ByVal hEngine As Long, _
    ByVal nameType As Long, _
    ByVal pOutBuffer As Long, _
    ByVal bufferSize As Long) As Long

Public Declare Sub Momiji_ClearTOCStructure Lib "Momiji.dll" Alias "ClearTOCStructure" (ByVal hEngine As Long)
Public Declare Function Momiji_GetTOCStructure Lib "Momiji.dll" Alias "GetTOCStructure" ( _
    ByVal hEngine As Long, _
    ByVal structureType As Long, _
    ByVal pBuffer As Long, _
    ByVal bufferSize As Long) As Long
Public Declare Function Momiji_SetTOCStructure Lib "Momiji.dll" Alias "SetTOCStructure" ( _
    ByVal hEngine As Long, _
    ByVal structureType As Long, _
    ByVal pBuffer As Long, _
    ByVal countOrSize As Long) As Long

Public Declare Function Momiji_GetFirstLBA Lib "Momiji.dll" Alias "GetFirstLBA" (ByVal hEngine As Long) As Long
Public Declare Function Momiji_GetLastLBA Lib "Momiji.dll" Alias "GetLastLBA" (ByVal hEngine As Long) As Long
Public Declare Sub Momiji_SetBufferingBlockLength Lib "Momiji.dll" Alias "SetBufferingBlockLength" (ByVal hEngine As Long, ByVal blockLength As Long)

Public Declare Function Momiji_ReadStart Lib "Momiji.dll" Alias "ReadStart" (ByVal hEngine As Long, ByVal readCommand As Long, ByVal readFlag As Long) As Long
Public Declare Function Momiji_ReadLBA Lib "Momiji.dll" Alias "ReadLBA" (ByVal hEngine As Long, ByVal pData As Long, ByVal lba As Long) As Long
Public Declare Function Momiji_ReadLBADummy Lib "Momiji.dll" Alias "ReadLBADummy" (ByVal hEngine As Long, ByVal pData As Long, ByVal lba As Long) As Long
Public Declare Function Momiji_ReadEnd Lib "Momiji.dll" Alias "ReadEnd" (ByVal hEngine As Long) As Long

Public Declare Function Momiji_WriteStart Lib "Momiji.dll" Alias "WriteStart" (ByVal hEngine As Long, ByVal writeCommand As Long, ByVal writeFlag As Long) As Long
Public Declare Function Momiji_WriteLBA Lib "Momiji.dll" Alias "WriteLBA" (ByVal hEngine As Long, ByVal dataType As Long, ByVal pData As Long, ByVal lba As Long) As Long
Public Declare Function Momiji_WriteFlush Lib "Momiji.dll" Alias "WriteFlush" (ByVal hEngine As Long, ByVal abortWrite As Long) As Long
Public Declare Function Momiji_WriteEnd Lib "Momiji.dll" Alias "WriteEnd" (ByVal hEngine As Long, ByVal abortWrite As Long) As Long
Public Declare Function Momiji_GetLastWroteLBA Lib "Momiji.dll" Alias "GetLastWroteLBA" (ByVal hEngine As Long) As Long

Public Declare Function Momiji_LoadTray Lib "Momiji.dll" Alias "LoadTray" (ByVal hEngine As Long, ByVal loadTray As Long) As Long
Public Declare Function Momiji_IsReady Lib "Momiji.dll" Alias "IsReady" (ByVal hEngine As Long) As Long
Public Declare Function Momiji_GetMediaType Lib "Momiji.dll" Alias "GetMediaType" (ByVal hEngine As Long) As Long
Public Declare Function Momiji_GetMediaFamilyType Lib "Momiji.dll" Alias "GetMediaFamilyType" (ByVal hEngine As Long) As Long
Public Declare Function Momiji_LockUnlock Lib "Momiji.dll" Alias "LockUnlock" (ByVal hEngine As Long, ByVal doLock As Long) As Long
Public Declare Function Momiji_Erase Lib "Momiji.dll" Alias "Erase" (ByVal hEngine As Long, ByVal quickly As Long) As Long
Public Declare Function Momiji_IsSupportMedia Lib "Momiji.dll" Alias "IsSupportMedia" (ByVal hEngine As Long, ByVal mediaType As Long) As Long
Public Declare Function Momiji_IsReadSupport Lib "Momiji.dll" Alias "IsReadSupport" (ByVal hEngine As Long, ByVal mediaFamily As Long) As Long
Public Declare Function Momiji_IsWriteSupport Lib "Momiji.dll" Alias "IsWriteSupport" (ByVal hEngine As Long, ByVal mediaFamily As Long) As Long
Public Declare Function Momiji_IsDiscEmpty Lib "Momiji.dll" Alias "IsDiscEmpty" (ByVal hEngine As Long) As Long
Public Declare Sub Momiji_ResetSense Lib "Momiji.dll" Alias "ResetSense" (ByVal hEngine As Long)
Public Declare Function Momiji_SetECCMode Lib "Momiji.dll" Alias "SetECCMode" (ByVal hEngine As Long, ByVal enabled As Long, ByVal retryCount As Long) As Long
Public Declare Function Momiji_SetReadSpeed Lib "Momiji.dll" Alias "SetReadSpeed" (ByVal hEngine As Long, ByVal speed As Long, ByVal cav As Long) As Long
Public Declare Function Momiji_SetWriteSpeed Lib "Momiji.dll" Alias "SetWriteSpeed" (ByVal hEngine As Long, ByVal speed As Long, ByVal cav As Long) As Long

' Despite the name, the original wrapper parses a speed descriptor list from this export.
Public Declare Function Momiji_GetWriteSpeed Lib "Momiji.dll" Alias "GetWriteSpeed" (ByVal hEngine As Long, ByVal pOutBuffer As Long, ByVal maxCount As Long) As Long

Public Declare Function Momiji_CheckWriteMode Lib "Momiji.dll" Alias "CheckWriteMode" (ByVal hEngine As Long, ByVal writeCommand As Long, ByVal testMode As Long, ByVal mediaFamilyType As Long) As Long
Public Declare Function Momiji_GetSCSIErrorStatus Lib "Momiji.dll" Alias "GetSCSIErrorStatus" (ByVal hEngine As Long) As Long
Public Declare Sub Momiji_SetWriteCacheBufferSize Lib "Momiji.dll" Alias "SetWriteCacheBufferSize" (ByVal hEngine As Long, ByVal sizeBytes As Long)
Public Declare Function Momiji_GetUsedWriteCacheBufferSize Lib "Momiji.dll" Alias "GetUsedWriteCacheBufferSize" (ByVal hEngine As Long) As Long

' ===== ZENKI.DLL =====
Public Declare Function Zenki_GetEngineVersion Lib "Zenki.dll" Alias "GetEngineVersion" () As Long
Public Declare Function Zenki_Initialize Lib "Zenki.dll" Alias "Initialize" () As Long
Public Declare Sub Zenki_Terminate Lib "Zenki.dll" Alias "Terminate" (ByVal hEngine As Long)

Public Declare Function Zenki_GetTOCStructure Lib "Zenki.dll" Alias "GetTOCStructure" (ByVal hEngine As Long, ByVal structureType As Long, ByVal pBuffer As Long, ByVal bufferSize As Long) As Long
Public Declare Function Zenki_InitISOFS Lib "Zenki.dll" Alias "InitISOFS" (ByVal hEngine As Long, ByVal fileSystemConstants As Long) As Long
Public Declare Sub Zenki_ClearISO Lib "Zenki.dll" Alias "ClearISO" (ByVal hEngine As Long)
Public Declare Function Zenki_IsISOEmpty Lib "Zenki.dll" Alias "IsISOEmpty" (ByVal hEngine As Long) As Long
Public Declare Function Zenki_AddISOFile Lib "Zenki.dll" Alias "AddISOFile" (ByVal hEngine As Long, ByVal pName As Long, ByVal pOriginalFileName As Long) As Long
Public Declare Function Zenki_AddISODummyFile Lib "Zenki.dll" Alias "AddISODummyFile" (ByVal hEngine As Long, ByVal pName As Long, ByVal attr As Long, ByVal unixTime As Long, ByVal reserved As Long) As Long
Public Declare Function Zenki_MakeISODirectory Lib "Zenki.dll" Alias "MakeISODirectory" (ByVal hEngine As Long, ByVal pName As Long, ByVal attr As Long, ByVal unixTime As Long) As Long
Public Declare Function Zenki_GetISONewFileDirectoryName Lib "Zenki.dll" Alias "GetISONewFileDirectoryName" (ByVal hEngine As Long, ByVal pOutBuffer As Long, ByVal bufferSize As Long) As Long
Public Declare Function Zenki_ChangeISODirectory Lib "Zenki.dll" Alias "ChangeISODirectory" (ByVal hEngine As Long, ByVal pName As Long) As Long
Public Declare Function Zenki_RemoveISOFile Lib "Zenki.dll" Alias "RemoveISOFile" (ByVal hEngine As Long, ByVal pName As Long) As Long
Public Declare Function Zenki_RenameISOFile Lib "Zenki.dll" Alias "RenameISOFile" (ByVal hEngine As Long, ByVal pName As Long, ByVal pNewName As Long) As Long
Public Declare Function Zenki_ChangeISOProperties Lib "Zenki.dll" Alias "ChangeISOProperties" (ByVal hEngine As Long, ByVal pName As Long, ByVal attr As Long, ByVal unixTime As Long) As Long
Public Declare Function Zenki_GetISOCurrentDirectory Lib "Zenki.dll" Alias "GetISOCurrentDirectory" (ByVal hEngine As Long, ByVal pOutBuffer As Long, ByVal bufferSize As Long) As Long
Public Declare Function Zenki_FindISOFirstFile Lib "Zenki.dll" Alias "FindISOFirstFile" (ByVal hEngine As Long, ByVal pOutStructOrBuffer As Long) As Long
Public Declare Function Zenki_FindISONextFile Lib "Zenki.dll" Alias "FindISONextFile" (ByVal hEngine As Long, ByVal pOutStructOrBuffer As Long) As Long

Public Declare Sub Zenki_ClearTrack Lib "Zenki.dll" Alias "ClearTrack" (ByVal hEngine As Long)
Public Declare Function Zenki_AddTrack Lib "Zenki.dll" Alias "AddTrack" (ByVal hEngine As Long, ByVal pWaveFileName As Long, ByVal pregap As Long, ByVal postgap As Long, ByVal flags As Long) As Long
Public Declare Function Zenki_ResetTrack Lib "Zenki.dll" Alias "ResetTrack" (ByVal hEngine As Long, ByVal trackIndexZeroBased As Long, ByVal pWaveFileName As Long, ByVal pregap As Long, ByVal postgap As Long, ByVal flags As Long) As Long
Public Declare Function Zenki_RemoveTrack Lib "Zenki.dll" Alias "RemoveTrack" (ByVal hEngine As Long, ByVal trackIndexZeroBased As Long) As Long
Public Declare Function Zenki_GetTrackCount Lib "Zenki.dll" Alias "GetTrackCount" (ByVal hEngine As Long) As Long
Public Declare Function Zenki_GetTrackInformation Lib "Zenki.dll" Alias "GetTrackInformation" (ByVal hEngine As Long, ByVal trackIndexZeroBased As Long, ByVal pOutStructOrBuffer As Long) As Long

Public Declare Function Zenki_ReadStart Lib "Zenki.dll" Alias "ReadStart" (ByVal hEngine As Long, ByVal reservedOrMode As Long) As Long
Public Declare Function Zenki_Read Lib "Zenki.dll" Alias "Read" (ByVal hEngine As Long, ByVal pData As Long, ByVal rawMode As Long) As Long
Public Declare Sub Zenki_ReadEnd Lib "Zenki.dll" Alias "ReadEnd" (ByVal hEngine As Long)

Public Declare Function Zenki_EnabledTrackText Lib "Zenki.dll" Alias "EnabledTrackText" (ByVal hEngine As Long, ByVal enabled As Long) As Long
Public Declare Function Zenki_SetTrackText Lib "Zenki.dll" Alias "SetTrackText" (ByVal hEngine As Long, ByVal trackNo As Long, ByVal language As Long, ByVal information As Long, ByVal pText As Long) As Long
Public Declare Function Zenki_GetTrackText Lib "Zenki.dll" Alias "GetTrackText" (ByVal hEngine As Long, ByVal trackNo As Long, ByVal language As Long, ByVal information As Long, ByVal pOutBuffer As Long, ByVal bufferSize As Long) As Long

' Helper functions for ANSI buffers.
Public Function AnsiZ(ByVal s As String) As Byte()
    AnsiZ = StrConv(s & vbNullChar, vbFromUnicode)
End Function

Public Function PtrOfFirstByte(ByRef b() As Byte) As Long
    If (Not Not b) = 0 Then
        PtrOfFirstByte = 0
    Else
        PtrOfFirstByte = VarPtr(b(LBound(b)))
    End If
End Function

Public Function StringFromAnsiZ(ByRef b() As Byte) As String
    Dim i As Long
    Dim n As Long
    On Error GoTo EmptyArray
    n = UBound(b) - LBound(b) + 1
    For i = LBound(b) To UBound(b)
        If b(i) = 0 Then
            n = i - LBound(b)
            Exit For
        End If
    Next
    If n <= 0 Then
        StringFromAnsiZ = vbNullString
    Else
        Dim tmp() As Byte
        ReDim tmp(0 To n - 1) As Byte
        For i = 0 To n - 1
            tmp(i) = b(LBound(b) + i)
        Next
        StringFromAnsiZ = StrConv(tmp, vbUnicode)
    End If
    Exit Function
EmptyArray:
    StringFromAnsiZ = vbNullString
End Function
