Attribute VB_Name = "modMomijiZenkiDeclares"
Option Explicit

' Clean VB6 declares for the rebuilt Momiji.dll / Zenki.dll.
' All DLL Boolean-like results are declared As Long to avoid VB6 Boolean width issues.

' ===== Momiji.dll =====
Public Declare Function Momiji_GetEngineVersion Lib "Momiji.dll" Alias "GetEngineVersion" () As Long
Public Declare Function Momiji_Initialize Lib "Momiji.dll" Alias "Initialize" () As Long
Public Declare Sub Momiji_Terminate Lib "Momiji.dll" Alias "Terminate" (ByVal hEngine As Long)

Public Declare Function Momiji_GetRemoteHosts Lib "Momiji.dll" Alias "GetRemoteHosts" ( _
    ByVal hEngine As Long, _
    ByVal broadcastIP As String, _
    ByVal portNo As Long, _
    ByVal outBytes As Long, _
    ByVal timeoutMs As Long, _
    ByVal pOutBuffer As Long) As Long

Public Declare Function Momiji_Open Lib "Momiji.dll" Alias "Open" ( _
    ByVal hEngine As Long, _
    ByVal driveIndexZeroBased As Long) As Long

Public Declare Function Momiji_OpenEx Lib "Momiji.dll" Alias "OpenEx" ( _
    ByVal hEngine As Long, _
    ByVal deviceName As String) As Long

Public Declare Function Momiji_Close Lib "Momiji.dll" Alias "Close" ( _
    ByVal hEngine As Long) As Long

Public Declare Function Momiji_GetDeviceName Lib "Momiji.dll" Alias "GetDeviceName" ( _
    ByVal hEngine As Long, _
    ByVal nameType As Long, _
    ByVal outName As String, _
    ByVal outBytes As Long) As Long

Public Declare Sub Momiji_ClearTOCStructure Lib "Momiji.dll" Alias "ClearTOCStructure" ( _
    ByVal hEngine As Long)

Public Declare Function Momiji_GetTOCStructure Lib "Momiji.dll" Alias "GetTOCStructure" ( _
    ByVal hEngine As Long, _
    ByVal tocType As Long, _
    ByVal pOutBuffer As Long, _
    ByVal outBytes As Long) As Long

Public Declare Function Momiji_SetTOCStructure Lib "Momiji.dll" Alias "SetTOCStructure" ( _
    ByVal hEngine As Long, _
    ByVal tocType As Long, _
    ByVal pInBuffer As Long, _
    ByVal inBytes As Long) As Long

Public Declare Function Momiji_GetFirstLBA Lib "Momiji.dll" Alias "GetFirstLBA" ( _
    ByVal hEngine As Long) As Long

Public Declare Function Momiji_GetLastLBA Lib "Momiji.dll" Alias "GetLastLBA" ( _
    ByVal hEngine As Long) As Long

Public Declare Function Momiji_SetBufferingBlockLength Lib "Momiji.dll" Alias "SetBufferingBlockLength" ( _
    ByVal hEngine As Long, _
    ByVal blockCount As Long) As Long

Public Declare Function Momiji_ReadStart Lib "Momiji.dll" Alias "ReadStart" ( _
    ByVal hEngine As Long, _
    ByVal readCommand As Long, _
    ByVal readFlag As Long) As Long

Public Declare Function Momiji_ReadLBA Lib "Momiji.dll" Alias "ReadLBA" ( _
    ByVal hEngine As Long, _
    ByVal pData As Long, _
    ByVal lba As Long) As Long

Public Declare Function Momiji_ReadLBADummy Lib "Momiji.dll" Alias "ReadLBADummy" ( _
    ByVal hEngine As Long, _
    ByVal pData As Long, _
    ByVal lba As Long) As Long

Public Declare Function Momiji_ReadEnd Lib "Momiji.dll" Alias "ReadEnd" ( _
    ByVal hEngine As Long) As Long

Public Declare Function Momiji_WriteStart Lib "Momiji.dll" Alias "WriteStart" ( _
    ByVal hEngine As Long, _
    ByVal writeCommand As Long, _
    ByVal writeFlag As Long) As Long

Public Declare Function Momiji_WriteLBA Lib "Momiji.dll" Alias "WriteLBA" ( _
    ByVal hEngine As Long, _
    ByVal dataType As Long, _
    ByVal pData As Long, _
    ByVal lba As Long) As Long

Public Declare Function Momiji_WriteFlush Lib "Momiji.dll" Alias "WriteFlush" ( _
    ByVal hEngine As Long, _
    ByVal abortFlag As Long) As Long

Public Declare Function Momiji_WriteEnd Lib "Momiji.dll" Alias "WriteEnd" ( _
    ByVal hEngine As Long, _
    ByVal abortFlag As Long) As Long

Public Declare Function Momiji_GetLastWroteLBA Lib "Momiji.dll" Alias "GetLastWroteLBA" ( _
    ByVal hEngine As Long) As Long

Public Declare Function Momiji_LoadTray Lib "Momiji.dll" Alias "LoadTray" ( _
    ByVal hEngine As Long, _
    ByVal loadTray As Long) As Long

Public Declare Function Momiji_IsReady Lib "Momiji.dll" Alias "IsReady" ( _
    ByVal hEngine As Long) As Long

Public Declare Function Momiji_GetMediaType Lib "Momiji.dll" Alias "GetMediaType" ( _
    ByVal hEngine As Long) As Long

Public Declare Function Momiji_GetMediaFamilyType Lib "Momiji.dll" Alias "GetMediaFamilyType" ( _
    ByVal hEngine As Long) As Long

Public Declare Function Momiji_LockUnlock Lib "Momiji.dll" Alias "LockUnlock" ( _
    ByVal hEngine As Long, _
    ByVal doLock As Long) As Long

Public Declare Function Momiji_Erase Lib "Momiji.dll" Alias "Erase" ( _
    ByVal hEngine As Long, _
    ByVal quickErase As Long) As Long

Public Declare Function Momiji_IsSupportMedia Lib "Momiji.dll" Alias "IsSupportMedia" ( _
    ByVal hEngine As Long, _
    ByVal mediaType As Long) As Long

Public Declare Function Momiji_IsReadSupport Lib "Momiji.dll" Alias "IsReadSupport" ( _
    ByVal hEngine As Long, _
    ByVal mediaFamily As Long) As Long

Public Declare Function Momiji_IsWriteSupport Lib "Momiji.dll" Alias "IsWriteSupport" ( _
    ByVal hEngine As Long, _
    ByVal mediaFamily As Long) As Long

Public Declare Function Momiji_IsDiscEmpty Lib "Momiji.dll" Alias "IsDiscEmpty" ( _
    ByVal hEngine As Long) As Long

Public Declare Function Momiji_ResetSense Lib "Momiji.dll" Alias "ResetSense" ( _
    ByVal hEngine As Long) As Long

Public Declare Function Momiji_SetECCMode Lib "Momiji.dll" Alias "SetECCMode" ( _
    ByVal hEngine As Long, _
    ByVal enabled As Long, _
    ByVal retryCount As Long) As Long

Public Declare Function Momiji_SetReadSpeed Lib "Momiji.dll" Alias "SetReadSpeed" ( _
    ByVal hEngine As Long, _
    ByVal speedKBs As Long, _
    ByVal cav As Long) As Long

Public Declare Function Momiji_SetWriteSpeed Lib "Momiji.dll" Alias "SetWriteSpeed" ( _
    ByVal hEngine As Long, _
    ByVal speedKBs As Long, _
    ByVal cav As Long) As Long

Public Declare Function Momiji_GetWriteSpeed Lib "Momiji.dll" Alias "GetWriteSpeed" ( _
    ByVal hEngine As Long, _
    ByVal pSpeeds As Long, _
    ByVal maxCount As Long) As Long

Public Declare Function Momiji_CheckWriteMode Lib "Momiji.dll" Alias "CheckWriteMode" ( _
    ByVal hEngine As Long, _
    ByVal writeCommand As Long, _
    ByVal testMode As Long, _
    ByVal mediaFamilyType As Long) As Long

Public Declare Function Momiji_GetSCSIErrorStatus Lib "Momiji.dll" Alias "GetSCSIErrorStatus" ( _
    ByVal hEngine As Long) As Long

Public Declare Function Momiji_SetWriteCacheBufferSize Lib "Momiji.dll" Alias "SetWriteCacheBufferSize" ( _
    ByVal hEngine As Long, _
    ByVal bytes As Long) As Long

Public Declare Function Momiji_GetUsedWriteCacheBufferSize Lib "Momiji.dll" Alias "GetUsedWriteCacheBufferSize" ( _
    ByVal hEngine As Long) As Long

' ===== Zenki.dll =====
Public Declare Function Zenki_GetEngineVersion Lib "Zenki.dll" Alias "GetEngineVersion" () As Long
Public Declare Function Zenki_Initialize Lib "Zenki.dll" Alias "Initialize" () As Long
Public Declare Sub Zenki_Terminate Lib "Zenki.dll" Alias "Terminate" (ByVal hEngine As Long)

Public Declare Function Zenki_GetTOCStructure Lib "Zenki.dll" Alias "GetTOCStructure" ( _
    ByVal hEngine As Long, _
    ByVal tocType As Long, _
    ByVal pOutBuffer As Long, _
    ByVal outBytes As Long) As Long

Public Declare Function Zenki_InitISOFS Lib "Zenki.dll" Alias "InitISOFS" ( _
    ByVal hEngine As Long, _
    ByVal isoFileSystemFlags As Long) As Long

Public Declare Sub Zenki_ClearISO Lib "Zenki.dll" Alias "ClearISO" ( _
    ByVal hEngine As Long)

Public Declare Function Zenki_IsISOEmpty Lib "Zenki.dll" Alias "IsISOEmpty" ( _
    ByVal hEngine As Long) As Long

Public Declare Function Zenki_AddISOFile Lib "Zenki.dll" Alias "AddISOFile" ( _
    ByVal hEngine As Long, _
    ByVal isoName As String, _
    ByVal originalFileName As String) As Long

Public Declare Function Zenki_AddISODummyFile Lib "Zenki.dll" Alias "AddISODummyFile" ( _
    ByVal hEngine As Long, _
    ByVal isoName As String, _
    ByVal byteSizeOrAttribute As Long, _
    ByVal attribute As Long, _
    ByVal timestampValue As Long) As Long

Public Declare Function Zenki_MakeISODirectory Lib "Zenki.dll" Alias "MakeISODirectory" ( _
    ByVal hEngine As Long, _
    ByVal isoName As String, _
    ByVal attribute As Long, _
    ByVal timestampValue As Long) As Long

Public Declare Sub Zenki_GetISONewFileDirectoryName Lib "Zenki.dll" Alias "GetISONewFileDirectoryName" ( _
    ByVal hEngine As Long, _
    ByVal outName As String, _
    ByVal outBytes As Long)

Public Declare Function Zenki_ChangeISODirectory Lib "Zenki.dll" Alias "ChangeISODirectory" ( _
    ByVal hEngine As Long, _
    ByVal isoName As String) As Long

Public Declare Function Zenki_RemoveISOFile Lib "Zenki.dll" Alias "RemoveISOFile" ( _
    ByVal hEngine As Long, _
    ByVal isoName As String) As Long

Public Declare Function Zenki_RenameISOFile Lib "Zenki.dll" Alias "RenameISOFile" ( _
    ByVal hEngine As Long, _
    ByVal oldIsoName As String, _
    ByVal newIsoName As String) As Long

Public Declare Function Zenki_ChangeISOProperties Lib "Zenki.dll" Alias "ChangeISOProperties" ( _
    ByVal hEngine As Long, _
    ByVal isoName As String, _
    ByVal attribute As Long, _
    ByVal timestampValue As Long) As Long

Public Declare Sub Zenki_GetISOCurrentDirectory Lib "Zenki.dll" Alias "GetISOCurrentDirectory" ( _
    ByVal hEngine As Long, _
    ByVal outDir As String, _
    ByVal outBytes As Long)

Public Declare Function Zenki_FindISOFirstFile Lib "Zenki.dll" Alias "FindISOFirstFile" ( _
    ByVal hEngine As Long, _
    ByVal pFindData As Long) As Long

Public Declare Function Zenki_FindISONextFile Lib "Zenki.dll" Alias "FindISONextFile" ( _
    ByVal hEngine As Long, _
    ByVal pFindData As Long) As Long

Public Declare Sub Zenki_ClearTrack Lib "Zenki.dll" Alias "ClearTrack" ( _
    ByVal hEngine As Long)

Public Declare Function Zenki_AddTrack Lib "Zenki.dll" Alias "AddTrack" ( _
    ByVal hEngine As Long, _
    ByVal waveOrImageFileName As String, _
    ByVal pregapLength As Long, _
    ByVal postgapLength As Long, _
    ByVal trackControlFlag As Long) As Long

Public Declare Function Zenki_ResetTrack Lib "Zenki.dll" Alias "ResetTrack" ( _
    ByVal hEngine As Long, _
    ByVal trackNo As Long, _
    ByVal waveOrImageFileName As String, _
    ByVal pregapLength As Long, _
    ByVal postgapLength As Long, _
    ByVal trackControlFlag As Long) As Long

Public Declare Function Zenki_RemoveTrack Lib "Zenki.dll" Alias "RemoveTrack" ( _
    ByVal hEngine As Long, _
    ByVal trackNo As Long) As Long

Public Declare Function Zenki_GetTrackCount Lib "Zenki.dll" Alias "GetTrackCount" ( _
    ByVal hEngine As Long) As Long

Public Declare Sub Zenki_GetTrackInformation Lib "Zenki.dll" Alias "GetTrackInformation" ( _
    ByVal hEngine As Long, _
    ByVal trackNo As Long, _
    ByVal pTrackInfo As Long)

Public Declare Function Zenki_ReadStart Lib "Zenki.dll" Alias "ReadStart" ( _
    ByVal hEngine As Long, _
    ByVal optionalPath As String) As Long

Public Declare Function Zenki_Read Lib "Zenki.dll" Alias "Read" ( _
    ByVal hEngine As Long, _
    ByVal pSectorBuffer As Long, _
    ByVal rawMode As Long) As Long

Public Declare Sub Zenki_ReadEnd Lib "Zenki.dll" Alias "ReadEnd" ( _
    ByVal hEngine As Long)

Public Declare Function Zenki_EnabledTrackText Lib "Zenki.dll" Alias "EnabledTrackText" ( _
    ByVal hEngine As Long, _
    ByVal enabled As Long) As Long

Public Declare Sub Zenki_SetTrackText Lib "Zenki.dll" Alias "SetTrackText" ( _
    ByVal hEngine As Long, _
    ByVal trackNo As Long, _
    ByVal languageNo As Long, _
    ByVal informationNo As Long, _
    ByVal textValue As String)

Public Declare Sub Zenki_GetTrackText Lib "Zenki.dll" Alias "GetTrackText" ( _
    ByVal hEngine As Long, _
    ByVal trackNo As Long, _
    ByVal languageNo As Long, _
    ByVal informationNo As Long, _
    ByVal outText As String, _
    ByVal outBytes As Long)

Public Function CStringFromBuffer(ByVal value As String) As String
    Dim pos As Long
    pos = InStr(1, value, vbNullChar)
    If pos > 0 Then
        CStringFromBuffer = Left$(value, pos - 1)
    Else
        CStringFromBuffer = value
    End If
End Function
