Attribute VB_Name = "DllDeclares"
Option Explicit

' Reconstructed from export tables and stdcall stack cleanup.
' 32-bit VB6 only. Handles are opaque Long values returned by Initialize().
' For buffer arguments use fixed-length String buffers or VarPtr(byteArray(0)) as noted.

' --- Momiji.dll ---
Public Declare Function Momiji_GetEngineVersion Lib "Momiji.dll" Alias "GetEngineVersion" () As Long
Public Declare Function Momiji_Initialize Lib "Momiji.dll" Alias "Initialize" () As Long
Public Declare Sub Momiji_Terminate Lib "Momiji.dll" Alias "Terminate" (ByVal h As Long)
Public Declare Function Momiji_GetRemoteHosts Lib "Momiji.dll" Alias "GetRemoteHosts" (ByVal h As Long, ByVal host As String, ByVal port As Long, ByVal outPtr As Long, ByVal outBytes As Long, ByVal flags As Long) As Long
Public Declare Function Momiji_Open Lib "Momiji.dll" Alias "Open" (ByVal h As Long, ByVal driveIndexZeroBased As Long) As Long
Public Declare Function Momiji_OpenEx Lib "Momiji.dll" Alias "OpenEx" (ByVal h As Long, ByVal deviceName As String) As Long
Public Declare Function Momiji_Close Lib "Momiji.dll" Alias "Close" (ByVal h As Long) As Long
Public Declare Function Momiji_GetDeviceName Lib "Momiji.dll" Alias "GetDeviceName" (ByVal h As Long, ByVal index As Long, ByVal outName As String, ByVal outBytes As Long) As Long
Public Declare Sub Momiji_ClearTOCStructure Lib "Momiji.dll" Alias "ClearTOCStructure" (ByVal h As Long)
Public Declare Function Momiji_GetTOCStructure Lib "Momiji.dll" Alias "GetTOCStructure" (ByVal h As Long, ByVal tocPtr As Long, ByVal maxEntries As Long, ByVal flags As Long) As Long
Public Declare Function Momiji_SetTOCStructure Lib "Momiji.dll" Alias "SetTOCStructure" (ByVal h As Long, ByVal tocPtr As Long, ByVal entryCount As Long, ByVal flags As Long) As Long
Public Declare Function Momiji_GetFirstLBA Lib "Momiji.dll" Alias "GetFirstLBA" (ByVal h As Long) As Long
Public Declare Function Momiji_GetLastLBA Lib "Momiji.dll" Alias "GetLastLBA" (ByVal h As Long) As Long
Public Declare Function Momiji_SetBufferingBlockLength Lib "Momiji.dll" Alias "SetBufferingBlockLength" (ByVal h As Long, ByVal blocks As Long) As Long
Public Declare Function Momiji_ReadStart Lib "Momiji.dll" Alias "ReadStart" (ByVal h As Long, ByVal startLBA As Long, ByVal blocks As Long) As Long
Public Declare Function Momiji_ReadLBA Lib "Momiji.dll" Alias "ReadLBA" (ByVal h As Long, ByVal lba As Long, ByVal bufferPtr As Long) As Long
Public Declare Function Momiji_ReadLBADummy Lib "Momiji.dll" Alias "ReadLBADummy" (ByVal h As Long, ByVal lba As Long, ByVal bufferPtr As Long) As Long
Public Declare Function Momiji_ReadEnd Lib "Momiji.dll" Alias "ReadEnd" (ByVal h As Long) As Long
Public Declare Function Momiji_WriteStart Lib "Momiji.dll" Alias "WriteStart" (ByVal h As Long, ByVal startLBA As Long, ByVal blocks As Long) As Long
Public Declare Function Momiji_WriteLBA Lib "Momiji.dll" Alias "WriteLBA" (ByVal h As Long, ByVal lba As Long, ByVal bufferPtr As Long, ByVal bytes As Long) As Long
Public Declare Function Momiji_WriteFlush Lib "Momiji.dll" Alias "WriteFlush" (ByVal h As Long, ByVal sync As Long) As Long
Public Declare Function Momiji_WriteEnd Lib "Momiji.dll" Alias "WriteEnd" (ByVal h As Long, ByVal finalizeDisc As Long) As Long
Public Declare Function Momiji_GetLastWroteLBA Lib "Momiji.dll" Alias "GetLastWroteLBA" (ByVal h As Long) As Long
Public Declare Function Momiji_LoadTray Lib "Momiji.dll" Alias "LoadTray" (ByVal h As Long, ByVal load As Long) As Long
Public Declare Function Momiji_IsReady Lib "Momiji.dll" Alias "IsReady" (ByVal h As Long) As Long
Public Declare Function Momiji_GetMediaType Lib "Momiji.dll" Alias "GetMediaType" (ByVal h As Long) As Long
Public Declare Function Momiji_GetMediaFamilyType Lib "Momiji.dll" Alias "GetMediaFamilyType" (ByVal h As Long) As Long
Public Declare Function Momiji_LockUnlock Lib "Momiji.dll" Alias "LockUnlock" (ByVal h As Long, ByVal lockDevice As Long) As Long
Public Declare Function Momiji_Erase Lib "Momiji.dll" Alias "Erase" (ByVal h As Long, ByVal fullErase As Long) As Long
Public Declare Function Momiji_IsSupportMedia Lib "Momiji.dll" Alias "IsSupportMedia" (ByVal h As Long, ByVal mediaType As Long) As Long
Public Declare Function Momiji_IsReadSupport Lib "Momiji.dll" Alias "IsReadSupport" (ByVal h As Long, ByVal mediaType As Long) As Long
Public Declare Function Momiji_IsWriteSupport Lib "Momiji.dll" Alias "IsWriteSupport" (ByVal h As Long, ByVal mediaType As Long) As Long
Public Declare Function Momiji_IsDiscEmpty Lib "Momiji.dll" Alias "IsDiscEmpty" (ByVal h As Long) As Long
Public Declare Function Momiji_ResetSense Lib "Momiji.dll" Alias "ResetSense" (ByVal h As Long) As Long
Public Declare Function Momiji_SetECCMode Lib "Momiji.dll" Alias "SetECCMode" (ByVal h As Long, ByVal enable As Long, ByVal mode As Long) As Long
Public Declare Function Momiji_SetReadSpeed Lib "Momiji.dll" Alias "SetReadSpeed" (ByVal h As Long, ByVal speed As Long, ByVal restoreDefault As Long) As Long
Public Declare Function Momiji_SetWriteSpeed Lib "Momiji.dll" Alias "SetWriteSpeed" (ByVal h As Long, ByVal speed As Long, ByVal restoreDefault As Long) As Long
Public Declare Function Momiji_GetWriteSpeed Lib "Momiji.dll" Alias "GetWriteSpeed" (ByVal h As Long, ByVal currentPtr As Long, ByVal maximumPtr As Long) As Long
Public Declare Function Momiji_CheckWriteMode Lib "Momiji.dll" Alias "CheckWriteMode" (ByVal h As Long, ByVal mode As Long, ByVal testWrite As Long, ByVal flags As Long) As Long
Public Declare Function Momiji_GetSCSIErrorStatus Lib "Momiji.dll" Alias "GetSCSIErrorStatus" (ByVal h As Long) As Long
Public Declare Function Momiji_SetWriteCacheBufferSize Lib "Momiji.dll" Alias "SetWriteCacheBufferSize" (ByVal h As Long, ByVal bytes As Long) As Long
Public Declare Function Momiji_GetUsedWriteCacheBufferSize Lib "Momiji.dll" Alias "GetUsedWriteCacheBufferSize" (ByVal h As Long) As Long

' --- Zenki.dll ---
Public Declare Function Zenki_GetEngineVersion Lib "Zenki.dll" Alias "GetEngineVersion" () As Long
Public Declare Function Zenki_Initialize Lib "Zenki.dll" Alias "Initialize" () As Long
Public Declare Sub Zenki_Terminate Lib "Zenki.dll" Alias "Terminate" (ByVal h As Long)
Public Declare Function Zenki_GetTOCStructure Lib "Zenki.dll" Alias "GetTOCStructure" (ByVal h As Long, ByVal tocPtr As Long, ByVal maxEntries As Long, ByVal flags As Long) As Long
Public Declare Function Zenki_InitISOFS Lib "Zenki.dll" Alias "InitISOFS" (ByVal h As Long, ByVal flags As Long) As Long
Public Declare Sub Zenki_ClearISO Lib "Zenki.dll" Alias "ClearISO" (ByVal h As Long)
Public Declare Function Zenki_IsISOEmpty Lib "Zenki.dll" Alias "IsISOEmpty" (ByVal h As Long) As Long
Public Declare Function Zenki_AddISOFile Lib "Zenki.dll" Alias "AddISOFile" (ByVal h As Long, ByVal sourcePath As String, ByVal isoPath As String) As Long
Public Declare Function Zenki_AddISODummyFile Lib "Zenki.dll" Alias "AddISODummyFile" (ByVal h As Long, ByVal isoPath As String, ByVal byteSize As Long, ByVal flags As Integer, ByVal timestampOrLba As Long) As Long
Public Declare Function Zenki_MakeISODirectory Lib "Zenki.dll" Alias "MakeISODirectory" (ByVal h As Long, ByVal isoPath As String, ByVal flags As Integer, ByVal timestampOrLba As Long) As Long
Public Declare Sub Zenki_GetISONewFileDirectoryName Lib "Zenki.dll" Alias "GetISONewFileDirectoryName" (ByVal h As Long, ByVal outName As String, ByVal outBytes As Long)
Public Declare Function Zenki_ChangeISODirectory Lib "Zenki.dll" Alias "ChangeISODirectory" (ByVal h As Long, ByVal isoPath As String) As Long
Public Declare Function Zenki_RemoveISOFile Lib "Zenki.dll" Alias "RemoveISOFile" (ByVal h As Long, ByVal isoPath As String) As Long
Public Declare Function Zenki_RenameISOFile Lib "Zenki.dll" Alias "RenameISOFile" (ByVal h As Long, ByVal oldIsoPath As String, ByVal newIsoPath As String) As Long
Public Declare Function Zenki_ChangeISOProperties Lib "Zenki.dll" Alias "ChangeISOProperties" (ByVal h As Long, ByVal isoPath As String, ByVal flags As Integer, ByVal timestampOrLba As Long) As Long
Public Declare Sub Zenki_GetISOCurrentDirectory Lib "Zenki.dll" Alias "GetISOCurrentDirectory" (ByVal h As Long, ByVal outDir As String, ByVal outBytes As Long)
Public Declare Function Zenki_FindISOFirstFile Lib "Zenki.dll" Alias "FindISOFirstFile" (ByVal h As Long, ByVal findDataPtr As Long) As Long
Public Declare Function Zenki_FindISONextFile Lib "Zenki.dll" Alias "FindISONextFile" (ByVal h As Long, ByVal findDataPtr As Long) As Long
Public Declare Sub Zenki_ClearTrack Lib "Zenki.dll" Alias "ClearTrack" (ByVal h As Long)
Public Declare Function Zenki_AddTrack Lib "Zenki.dll" Alias "AddTrack" (ByVal h As Long, ByVal filePath As String, ByVal trackType As Long, ByVal sectorSize As Long, ByVal sectorCount As Long) As Long
Public Declare Function Zenki_ResetTrack Lib "Zenki.dll" Alias "ResetTrack" (ByVal h As Long, ByVal trackIndex As Long, ByVal filePath As String, ByVal trackType As Long, ByVal sectorSize As Long, ByVal sectorCount As Long) As Long
Public Declare Function Zenki_RemoveTrack Lib "Zenki.dll" Alias "RemoveTrack" (ByVal h As Long, ByVal trackIndex As Long) As Long
Public Declare Function Zenki_GetTrackCount Lib "Zenki.dll" Alias "GetTrackCount" (ByVal h As Long) As Long
Public Declare Sub Zenki_GetTrackInformation Lib "Zenki.dll" Alias "GetTrackInformation" (ByVal h As Long, ByVal trackIndex As Long, ByVal outInfoPtr As Long)
Public Declare Function Zenki_ReadStart Lib "Zenki.dll" Alias "ReadStart" (ByVal h As Long, ByVal imagePath As String) As Long
Public Declare Function Zenki_Read Lib "Zenki.dll" Alias "Read" (ByVal h As Long, ByVal sectorBufferPtr As Long, ByVal rawMode As Long) As Long
Public Declare Sub Zenki_ReadEnd Lib "Zenki.dll" Alias "ReadEnd" (ByVal h As Long)
Public Declare Function Zenki_EnabledTrackText Lib "Zenki.dll" Alias "EnabledTrackText" (ByVal h As Long, ByVal enable As Long) As Long
Public Declare Sub Zenki_SetTrackText Lib "Zenki.dll" Alias "SetTrackText" (ByVal h As Long, ByVal trackIndex As Long, ByVal packType As Long, ByVal itemIndex As Long, ByVal text As String)
Public Declare Sub Zenki_GetTrackText Lib "Zenki.dll" Alias "GetTrackText" (ByVal h As Long, ByVal trackIndex As Long, ByVal packType As Long, ByVal itemIndex As Long, ByVal outText As String, ByVal outBytes As Long)
