#ifndef MOMIJI_H
#define MOMIJI_H

#ifdef __cplusplus
extern "C" {
#endif

#if defined(_MSC_VER) || defined(__clang__)
#define MOMIJI_STDCALL __stdcall
#else
#define MOMIJI_STDCALL __attribute__((stdcall))
#endif

typedef void* MOMIJI_HANDLE;
typedef unsigned int   MOMIJI_U32;
typedef int            MOMIJI_I32;
typedef unsigned char  MOMIJI_BOOL;

/* Exported ABI reconstructed from MOMIJI.DLL.
   Version encoding appears to be: major*10000 + minor*100 + patch.
   Original GetEngineVersion() returns 0x2969 == 10601 == 1.06.01. */
MOMIJI_U32    MOMIJI_STDCALL GetEngineVersion(void);
MOMIJI_HANDLE MOMIJI_STDCALL Initialize(void);
void          MOMIJI_STDCALL Terminate(MOMIJI_HANDLE ctx);

/* Low-level device / remote-SCSI functions. Unknown parameters are kept as u32/void*. */
MOMIJI_BOOL MOMIJI_STDCALL GetRemoteHosts(MOMIJI_HANDLE ctx, const char* host, MOMIJI_I32 port, void* out, MOMIJI_U32 outBytes, MOMIJI_U32 flags);
MOMIJI_BOOL MOMIJI_STDCALL Open(MOMIJI_HANDLE ctx, MOMIJI_I32 driveIndexZeroBased);
MOMIJI_BOOL MOMIJI_STDCALL OpenEx(MOMIJI_HANDLE ctx, const char* deviceName);
MOMIJI_BOOL MOMIJI_STDCALL Close(MOMIJI_HANDLE ctx);
MOMIJI_BOOL MOMIJI_STDCALL GetDeviceName(MOMIJI_HANDLE ctx, MOMIJI_I32 index, char* outName, MOMIJI_U32 outBytes);

/* TOC / LBA / read-write operations. */
void       MOMIJI_STDCALL ClearTOCStructure(MOMIJI_HANDLE ctx);
MOMIJI_U32 MOMIJI_STDCALL GetTOCStructure(MOMIJI_HANDLE ctx, void* toc, MOMIJI_U32 maxEntries, MOMIJI_U32 flags);
MOMIJI_U32 MOMIJI_STDCALL SetTOCStructure(MOMIJI_HANDLE ctx, const void* toc, MOMIJI_U32 entryCount, MOMIJI_U32 flags);
MOMIJI_I32 MOMIJI_STDCALL GetFirstLBA(MOMIJI_HANDLE ctx);
MOMIJI_I32 MOMIJI_STDCALL GetLastLBA(MOMIJI_HANDLE ctx);
MOMIJI_U32 MOMIJI_STDCALL SetBufferingBlockLength(MOMIJI_HANDLE ctx, MOMIJI_U32 blocks);
MOMIJI_U32 MOMIJI_STDCALL ReadStart(MOMIJI_HANDLE ctx, MOMIJI_I32 startLBA, MOMIJI_U32 blocks);
MOMIJI_BOOL MOMIJI_STDCALL ReadLBA(MOMIJI_HANDLE ctx, MOMIJI_I32 lba, void* buffer);
MOMIJI_BOOL MOMIJI_STDCALL ReadLBADummy(MOMIJI_HANDLE ctx, MOMIJI_I32 lba, void* buffer);
MOMIJI_BOOL MOMIJI_STDCALL ReadEnd(MOMIJI_HANDLE ctx);
MOMIJI_U32 MOMIJI_STDCALL WriteStart(MOMIJI_HANDLE ctx, MOMIJI_I32 startLBA, MOMIJI_U32 blocks);
MOMIJI_BOOL MOMIJI_STDCALL WriteLBA(MOMIJI_HANDLE ctx, MOMIJI_I32 lba, const void* buffer, MOMIJI_U32 bytes);
MOMIJI_BOOL MOMIJI_STDCALL WriteFlush(MOMIJI_HANDLE ctx, MOMIJI_BOOL sync);
MOMIJI_BOOL MOMIJI_STDCALL WriteEnd(MOMIJI_HANDLE ctx, MOMIJI_BOOL finalizeDisc);
MOMIJI_I32 MOMIJI_STDCALL GetLastWroteLBA(MOMIJI_HANDLE ctx);

/* Media control and capability helpers. */
MOMIJI_BOOL MOMIJI_STDCALL LoadTray(MOMIJI_HANDLE ctx, MOMIJI_BOOL load);
MOMIJI_BOOL MOMIJI_STDCALL IsReady(MOMIJI_HANDLE ctx);
MOMIJI_U32  MOMIJI_STDCALL GetMediaType(MOMIJI_HANDLE ctx);
MOMIJI_U32  MOMIJI_STDCALL GetMediaFamilyType(MOMIJI_HANDLE ctx);
MOMIJI_BOOL MOMIJI_STDCALL LockUnlock(MOMIJI_HANDLE ctx, MOMIJI_BOOL lock);
MOMIJI_BOOL MOMIJI_STDCALL Erase(MOMIJI_HANDLE ctx, MOMIJI_BOOL fullErase);
MOMIJI_BOOL MOMIJI_STDCALL IsSupportMedia(MOMIJI_HANDLE ctx, MOMIJI_U32 mediaType);
MOMIJI_BOOL MOMIJI_STDCALL IsReadSupport(MOMIJI_HANDLE ctx, MOMIJI_U32 mediaType);
MOMIJI_BOOL MOMIJI_STDCALL IsWriteSupport(MOMIJI_HANDLE ctx, MOMIJI_U32 mediaType);
MOMIJI_BOOL MOMIJI_STDCALL IsDiscEmpty(MOMIJI_HANDLE ctx);
MOMIJI_U32  MOMIJI_STDCALL ResetSense(MOMIJI_HANDLE ctx);
MOMIJI_BOOL MOMIJI_STDCALL SetECCMode(MOMIJI_HANDLE ctx, MOMIJI_BOOL enable, MOMIJI_U32 mode);
MOMIJI_BOOL MOMIJI_STDCALL SetReadSpeed(MOMIJI_HANDLE ctx, MOMIJI_U32 speed, MOMIJI_BOOL restoreDefault);
MOMIJI_BOOL MOMIJI_STDCALL SetWriteSpeed(MOMIJI_HANDLE ctx, MOMIJI_U32 speed, MOMIJI_BOOL restoreDefault);
MOMIJI_U32  MOMIJI_STDCALL GetWriteSpeed(MOMIJI_HANDLE ctx, MOMIJI_U32* current, MOMIJI_U32* maximum);
MOMIJI_BOOL MOMIJI_STDCALL CheckWriteMode(MOMIJI_HANDLE ctx, MOMIJI_U32 mode, MOMIJI_BOOL testWrite, MOMIJI_U32 flags);
MOMIJI_U32  MOMIJI_STDCALL GetSCSIErrorStatus(MOMIJI_HANDLE ctx);
MOMIJI_U32  MOMIJI_STDCALL SetWriteCacheBufferSize(MOMIJI_HANDLE ctx, MOMIJI_U32 bytes);
MOMIJI_U32  MOMIJI_STDCALL GetUsedWriteCacheBufferSize(MOMIJI_HANDLE ctx);

#ifdef __cplusplus
}
#endif
#endif
