#ifndef ZENKI_H
#define ZENKI_H

#ifdef __cplusplus
extern "C" {
#endif

#if defined(_MSC_VER) || defined(__clang__)
#define ZENKI_STDCALL __stdcall
#else
#define ZENKI_STDCALL __attribute__((stdcall))
#endif

typedef void* ZENKI_HANDLE;
typedef unsigned int   ZENKI_U32;
typedef int            ZENKI_I32;
typedef unsigned short ZENKI_U16;
typedef unsigned char  ZENKI_BOOL;

/* Track-info buffer inferred from GetTrackInformation():
   the original copies a text field at offset 0 and writes 4 DWORDs at +0x400..+0x40c. */
typedef struct ZENKI_TRACK_INFO_TAG {
    char text[1024];
    ZENKI_U32 value0;
    ZENKI_U32 value1;
    ZENKI_U32 value2;
    ZENKI_U32 value3;
} ZENKI_TRACK_INFO;

/* Exported ABI reconstructed from ZENKI.DLL.
   GetEngineVersion() returns 0x28A7 == 10407 == 1.04.07. */
ZENKI_U32    ZENKI_STDCALL GetEngineVersion(void);
ZENKI_HANDLE ZENKI_STDCALL Initialize(void);
void         ZENKI_STDCALL Terminate(ZENKI_HANDLE ctx);

ZENKI_U32  ZENKI_STDCALL GetTOCStructure(ZENKI_HANDLE ctx, void* toc, ZENKI_U32 maxEntries, ZENKI_U32 flags);
ZENKI_BOOL ZENKI_STDCALL InitISOFS(ZENKI_HANDLE ctx, ZENKI_U32 flags);
void       ZENKI_STDCALL ClearISO(ZENKI_HANDLE ctx);
ZENKI_BOOL ZENKI_STDCALL IsISOEmpty(ZENKI_HANDLE ctx);
ZENKI_BOOL ZENKI_STDCALL AddISOFile(ZENKI_HANDLE ctx, const char* sourcePath, const char* isoPath);
ZENKI_BOOL ZENKI_STDCALL AddISODummyFile(ZENKI_HANDLE ctx, const char* isoPath, ZENKI_U32 byteSize, ZENKI_U16 flags, ZENKI_U32 timestampOrLba);
ZENKI_BOOL ZENKI_STDCALL MakeISODirectory(ZENKI_HANDLE ctx, const char* isoPath, ZENKI_U16 flags, ZENKI_U32 timestampOrLba);
void       ZENKI_STDCALL GetISONewFileDirectoryName(ZENKI_HANDLE ctx, char* outName, ZENKI_U32 outBytes);
ZENKI_BOOL ZENKI_STDCALL ChangeISODirectory(ZENKI_HANDLE ctx, const char* isoPath);
ZENKI_BOOL ZENKI_STDCALL RemoveISOFile(ZENKI_HANDLE ctx, const char* isoPath);
ZENKI_BOOL ZENKI_STDCALL RenameISOFile(ZENKI_HANDLE ctx, const char* oldIsoPath, const char* newIsoPath);
ZENKI_BOOL ZENKI_STDCALL ChangeISOProperties(ZENKI_HANDLE ctx, const char* isoPath, ZENKI_U16 flags, ZENKI_U32 timestampOrLba);
void       ZENKI_STDCALL GetISOCurrentDirectory(ZENKI_HANDLE ctx, char* outDir, ZENKI_U32 outBytes);
ZENKI_BOOL ZENKI_STDCALL FindISOFirstFile(ZENKI_HANDLE ctx, void* findData);
ZENKI_BOOL ZENKI_STDCALL FindISONextFile(ZENKI_HANDLE ctx, void* findData);

void       ZENKI_STDCALL ClearTrack(ZENKI_HANDLE ctx);
ZENKI_BOOL ZENKI_STDCALL AddTrack(ZENKI_HANDLE ctx, const char* filePath, ZENKI_U32 trackType, ZENKI_U32 sectorSize, ZENKI_U32 sectorCount);
ZENKI_BOOL ZENKI_STDCALL ResetTrack(ZENKI_HANDLE ctx, ZENKI_U32 trackIndex, const char* filePath, ZENKI_U32 trackType, ZENKI_U32 sectorSize, ZENKI_U32 sectorCount);
ZENKI_BOOL ZENKI_STDCALL RemoveTrack(ZENKI_HANDLE ctx, ZENKI_U32 trackIndex);
ZENKI_U32  ZENKI_STDCALL GetTrackCount(ZENKI_HANDLE ctx);
void       ZENKI_STDCALL GetTrackInformation(ZENKI_HANDLE ctx, ZENKI_U32 trackIndex, ZENKI_TRACK_INFO* outInfo);
ZENKI_I32  ZENKI_STDCALL ReadStart(ZENKI_HANDLE ctx, const char* imagePath);
ZENKI_BOOL ZENKI_STDCALL Read(ZENKI_HANDLE ctx, void* sectorBuffer, ZENKI_BOOL rawMode);
void       ZENKI_STDCALL ReadEnd(ZENKI_HANDLE ctx);
ZENKI_BOOL ZENKI_STDCALL EnabledTrackText(ZENKI_HANDLE ctx, ZENKI_BOOL enable);
void       ZENKI_STDCALL SetTrackText(ZENKI_HANDLE ctx, ZENKI_U32 trackIndex, ZENKI_U32 packType, ZENKI_U32 itemIndex, const char* text);
void       ZENKI_STDCALL GetTrackText(ZENKI_HANDLE ctx, ZENKI_U32 trackIndex, ZENKI_U32 packType, ZENKI_U32 itemIndex, char* outText, ZENKI_U32 outBytes);

#ifdef __cplusplus
}
#endif
#endif
