#include "zenki.h"

#define ZENKI_MAX_CONTEXTS 16
#define ZENKI_MAX_TRACKS   99
#define ZENKI_MAX_ISO_ITEMS 256
#define ZENKI_MAGIC 0x5A454E4Bu /* ZENK */

typedef struct ZENKI_TRACK_TAG {
    ZENKI_BOOL used;
    char path[260];
    ZENKI_U32 trackType;
    ZENKI_U32 sectorSize;
    ZENKI_U32 sectorCount;
    char text[128];
} ZENKI_TRACK;

typedef struct ZENKI_CONTEXT_TAG {
    ZENKI_U32 magic;
    ZENKI_BOOL used;
    ZENKI_BOOL trackTextEnabled;
    ZENKI_U32 isoItemCount;
    ZENKI_U32 trackCount;
    char currentDir[260];
    char lastIsoName[260];
    ZENKI_TRACK tracks[ZENKI_MAX_TRACKS];
} ZENKI_CONTEXT;

static ZENKI_CONTEXT g_ctx[ZENKI_MAX_CONTEXTS];

static ZENKI_CONTEXT* ctx_from_handle(ZENKI_HANDLE h) {
    ZENKI_CONTEXT* c = (ZENKI_CONTEXT*)h;
    if (!c || c->magic != ZENKI_MAGIC || !c->used) return 0;
    return c;
}

static void copy_str(char* dst, ZENKI_U32 cap, const char* src) {
    ZENKI_U32 i;
    if (!dst || cap == 0) return;
    if (!src) src = "";
    for (i = 0; i + 1 < cap && src[i]; ++i) dst[i] = src[i];
    dst[i] = 0;
}

static void zero_track(ZENKI_TRACK* t) {
    t->used = 0;
    t->path[0] = 0;
    t->trackType = 0;
    t->sectorSize = 0;
    t->sectorCount = 0;
    t->text[0] = 0;
}

static void reset_context(ZENKI_CONTEXT* c) {
    int i;
    c->trackTextEnabled = 0;
    c->isoItemCount = 0;
    c->trackCount = 0;
    copy_str(c->currentDir, sizeof(c->currentDir), "/");
    copy_str(c->lastIsoName, sizeof(c->lastIsoName), "");
    for (i = 0; i < ZENKI_MAX_TRACKS; ++i) zero_track(&c->tracks[i]);
}

ZENKI_U32 ZENKI_STDCALL GetEngineVersion(void) { return 0x28A7u; }

ZENKI_HANDLE ZENKI_STDCALL Initialize(void) {
    int i;
    for (i = 0; i < ZENKI_MAX_CONTEXTS; ++i) {
        if (!g_ctx[i].used) {
            ZENKI_CONTEXT* c = &g_ctx[i];
            c->magic = ZENKI_MAGIC;
            c->used = 1;
            reset_context(c);
            return (ZENKI_HANDLE)c;
        }
    }
    return 0;
}

void ZENKI_STDCALL Terminate(ZENKI_HANDLE h) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    if (!c) return;
    reset_context(c);
    c->used = 0;
    c->magic = 0;
}

ZENKI_U32 ZENKI_STDCALL GetTOCStructure(ZENKI_HANDLE h, void* toc, ZENKI_U32 maxEntries, ZENKI_U32 flags) {
    (void)toc; (void)maxEntries; (void)flags;
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    return c ? c->trackCount : 0;
}

ZENKI_BOOL ZENKI_STDCALL InitISOFS(ZENKI_HANDLE h, ZENKI_U32 flags) {
    (void)flags;
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    if (!c) return 0;
    c->isoItemCount = 0;
    copy_str(c->currentDir, sizeof(c->currentDir), "/");
    copy_str(c->lastIsoName, sizeof(c->lastIsoName), "");
    return 1;
}

void ZENKI_STDCALL ClearISO(ZENKI_HANDLE h) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    if (!c) return;
    c->isoItemCount = 0;
    copy_str(c->lastIsoName, sizeof(c->lastIsoName), "");
}

ZENKI_BOOL ZENKI_STDCALL IsISOEmpty(ZENKI_HANDLE h) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    return (c && c->isoItemCount == 0) ? 1 : 0;
}

ZENKI_BOOL ZENKI_STDCALL AddISOFile(ZENKI_HANDLE h, const char* sourcePath, const char* isoPath) {
    (void)sourcePath;
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    if (!c || !isoPath || c->isoItemCount >= ZENKI_MAX_ISO_ITEMS) return 0;
    ++c->isoItemCount;
    copy_str(c->lastIsoName, sizeof(c->lastIsoName), isoPath);
    return 1;
}

ZENKI_BOOL ZENKI_STDCALL AddISODummyFile(ZENKI_HANDLE h, const char* isoPath, ZENKI_U32 byteSize, ZENKI_U16 flags, ZENKI_U32 timestampOrLba) {
    (void)byteSize; (void)flags; (void)timestampOrLba;
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    if (!c || !isoPath || c->isoItemCount >= ZENKI_MAX_ISO_ITEMS) return 0;
    ++c->isoItemCount;
    copy_str(c->lastIsoName, sizeof(c->lastIsoName), isoPath);
    return 1;
}

ZENKI_BOOL ZENKI_STDCALL MakeISODirectory(ZENKI_HANDLE h, const char* isoPath, ZENKI_U16 flags, ZENKI_U32 timestampOrLba) {
    (void)flags; (void)timestampOrLba;
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    if (!c || !isoPath || c->isoItemCount >= ZENKI_MAX_ISO_ITEMS) return 0;
    ++c->isoItemCount;
    copy_str(c->lastIsoName, sizeof(c->lastIsoName), isoPath);
    return 1;
}

void ZENKI_STDCALL GetISONewFileDirectoryName(ZENKI_HANDLE h, char* outName, ZENKI_U32 outBytes) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    copy_str(outName, outBytes, c ? c->lastIsoName : "");
}

ZENKI_BOOL ZENKI_STDCALL ChangeISODirectory(ZENKI_HANDLE h, const char* isoPath) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    if (!c || !isoPath) return 0;
    copy_str(c->currentDir, sizeof(c->currentDir), isoPath);
    return 1;
}

ZENKI_BOOL ZENKI_STDCALL RemoveISOFile(ZENKI_HANDLE h, const char* isoPath) {
    (void)isoPath;
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    if (!c || c->isoItemCount == 0) return 0;
    --c->isoItemCount;
    return 1;
}

ZENKI_BOOL ZENKI_STDCALL RenameISOFile(ZENKI_HANDLE h, const char* oldIsoPath, const char* newIsoPath) {
    (void)oldIsoPath;
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    if (!c || !newIsoPath) return 0;
    copy_str(c->lastIsoName, sizeof(c->lastIsoName), newIsoPath);
    return 1;
}

ZENKI_BOOL ZENKI_STDCALL ChangeISOProperties(ZENKI_HANDLE h, const char* isoPath, ZENKI_U16 flags, ZENKI_U32 timestampOrLba) {
    (void)flags; (void)timestampOrLba;
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    if (!c || !isoPath) return 0;
    copy_str(c->lastIsoName, sizeof(c->lastIsoName), isoPath);
    return 1;
}

void ZENKI_STDCALL GetISOCurrentDirectory(ZENKI_HANDLE h, char* outDir, ZENKI_U32 outBytes) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    copy_str(outDir, outBytes, c ? c->currentDir : "");
}

ZENKI_BOOL ZENKI_STDCALL FindISOFirstFile(ZENKI_HANDLE h, void* findData) { (void)findData; return ctx_from_handle(h) ? 0 : 0; }
ZENKI_BOOL ZENKI_STDCALL FindISONextFile(ZENKI_HANDLE h, void* findData) { (void)findData; return ctx_from_handle(h) ? 0 : 0; }

void ZENKI_STDCALL ClearTrack(ZENKI_HANDLE h) {
    int i;
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    if (!c) return;
    c->trackCount = 0;
    for (i = 0; i < ZENKI_MAX_TRACKS; ++i) zero_track(&c->tracks[i]);
}

ZENKI_BOOL ZENKI_STDCALL AddTrack(ZENKI_HANDLE h, const char* filePath, ZENKI_U32 trackType, ZENKI_U32 sectorSize, ZENKI_U32 sectorCount) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    ZENKI_TRACK* t;
    if (!c || !filePath || c->trackCount >= ZENKI_MAX_TRACKS) return 0;
    t = &c->tracks[c->trackCount];
    t->used = 1;
    copy_str(t->path, sizeof(t->path), filePath);
    t->trackType = trackType;
    t->sectorSize = sectorSize;
    t->sectorCount = sectorCount;
    ++c->trackCount;
    return 1;
}

ZENKI_BOOL ZENKI_STDCALL ResetTrack(ZENKI_HANDLE h, ZENKI_U32 trackIndex, const char* filePath, ZENKI_U32 trackType, ZENKI_U32 sectorSize, ZENKI_U32 sectorCount) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    ZENKI_TRACK* t;
    if (!c || trackIndex >= ZENKI_MAX_TRACKS || !filePath) return 0;
    t = &c->tracks[trackIndex];
    t->used = 1;
    copy_str(t->path, sizeof(t->path), filePath);
    t->trackType = trackType;
    t->sectorSize = sectorSize;
    t->sectorCount = sectorCount;
    if (trackIndex >= c->trackCount) c->trackCount = trackIndex + 1;
    return 1;
}

ZENKI_BOOL ZENKI_STDCALL RemoveTrack(ZENKI_HANDLE h, ZENKI_U32 trackIndex) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    if (!c || trackIndex >= c->trackCount) return 0;
    zero_track(&c->tracks[trackIndex]);
    while (c->trackCount > 0 && !c->tracks[c->trackCount - 1].used) --c->trackCount;
    return 1;
}

ZENKI_U32 ZENKI_STDCALL GetTrackCount(ZENKI_HANDLE h) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    return c ? c->trackCount : 0;
}

void ZENKI_STDCALL GetTrackInformation(ZENKI_HANDLE h, ZENKI_U32 trackIndex, ZENKI_TRACK_INFO* outInfo) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    ZENKI_TRACK* t;
    if (!outInfo) return;
    copy_str(outInfo->text, sizeof(outInfo->text), "");
    outInfo->value0 = outInfo->value1 = outInfo->value2 = outInfo->value3 = 0;
    if (!c || trackIndex >= c->trackCount) return;
    t = &c->tracks[trackIndex];
    copy_str(outInfo->text, sizeof(outInfo->text), t->path);
    outInfo->value0 = t->trackType;
    outInfo->value1 = t->sectorSize;
    outInfo->value2 = t->sectorCount;
    outInfo->value3 = t->used;
}

ZENKI_I32 ZENKI_STDCALL ReadStart(ZENKI_HANDLE h, const char* imagePath) { (void)imagePath; return ctx_from_handle(h) ? 0 : -1; }
ZENKI_BOOL ZENKI_STDCALL Read(ZENKI_HANDLE h, void* sectorBuffer, ZENKI_BOOL rawMode) { (void)h; (void)sectorBuffer; (void)rawMode; return 0; }
void ZENKI_STDCALL ReadEnd(ZENKI_HANDLE h) { (void)h; }
ZENKI_BOOL ZENKI_STDCALL EnabledTrackText(ZENKI_HANDLE h, ZENKI_BOOL enable) { ZENKI_CONTEXT* c = ctx_from_handle(h); if (!c) return 0; c->trackTextEnabled = enable ? 1 : 0; return 1; }

void ZENKI_STDCALL SetTrackText(ZENKI_HANDLE h, ZENKI_U32 trackIndex, ZENKI_U32 packType, ZENKI_U32 itemIndex, const char* text) {
    (void)packType; (void)itemIndex;
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    if (!c || trackIndex >= c->trackCount) return;
    copy_str(c->tracks[trackIndex].text, sizeof(c->tracks[trackIndex].text), text);
}

void ZENKI_STDCALL GetTrackText(ZENKI_HANDLE h, ZENKI_U32 trackIndex, ZENKI_U32 packType, ZENKI_U32 itemIndex, char* outText, ZENKI_U32 outBytes) {
    (void)packType; (void)itemIndex;
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    if (!c || trackIndex >= c->trackCount) { copy_str(outText, outBytes, ""); return; }
    copy_str(outText, outBytes, c->tracks[trackIndex].text);
}
