#include "momiji.h"

#define MOMIJI_MAX_CONTEXTS 16
#define MOMIJI_MAGIC 0x4D4F4D49u /* MOMI */

typedef struct MOMIJI_CONTEXT_TAG {
    MOMIJI_U32 magic;
    MOMIJI_BOOL used;
    MOMIJI_BOOL opened;
    MOMIJI_BOOL locked;
    MOMIJI_I32 driveIndex;
    MOMIJI_U32 writeCacheBytes;
    MOMIJI_U32 blockLength;
    MOMIJI_U32 readSpeed;
    MOMIJI_U32 writeSpeed;
    MOMIJI_I32 firstLBA;
    MOMIJI_I32 lastLBA;
    MOMIJI_I32 lastWroteLBA;
    char deviceName[64];
} MOMIJI_CONTEXT;

static MOMIJI_CONTEXT g_ctx[MOMIJI_MAX_CONTEXTS];

static MOMIJI_CONTEXT* ctx_from_handle(MOMIJI_HANDLE h) {
    MOMIJI_CONTEXT* c = (MOMIJI_CONTEXT*)h;
    if (!c || c->magic != MOMIJI_MAGIC || !c->used) return 0;
    return c;
}

static void copy_str(char* dst, MOMIJI_U32 cap, const char* src) {
    MOMIJI_U32 i;
    if (!dst || cap == 0) return;
    if (!src) src = "";
    for (i = 0; i + 1 < cap && src[i]; ++i) dst[i] = src[i];
    dst[i] = 0;
}

static void set_device_from_drive(MOMIJI_CONTEXT* c, int index) {
    c->deviceName[0] = '\\';
    c->deviceName[1] = '\\';
    c->deviceName[2] = '.';
    c->deviceName[3] = '\\';
    c->deviceName[4] = (char)('A' + (index & 31));
    c->deviceName[5] = ':';
    c->deviceName[6] = 0;
}

MOMIJI_U32 MOMIJI_STDCALL GetEngineVersion(void) { return 0x2969u; }

MOMIJI_HANDLE MOMIJI_STDCALL Initialize(void) {
    int i;
    for (i = 0; i < MOMIJI_MAX_CONTEXTS; ++i) {
        if (!g_ctx[i].used) {
            MOMIJI_CONTEXT* c = &g_ctx[i];
            c->magic = MOMIJI_MAGIC;
            c->used = 1;
            c->opened = 0;
            c->locked = 0;
            c->driveIndex = -1;
            c->writeCacheBytes = 0;
            c->blockLength = 0;
            c->readSpeed = 0;
            c->writeSpeed = 0;
            c->firstLBA = 0;
            c->lastLBA = 0;
            c->lastWroteLBA = -1;
            copy_str(c->deviceName, sizeof(c->deviceName), "MOMIJI_STUB");
            return (MOMIJI_HANDLE)c;
        }
    }
    return 0;
}

void MOMIJI_STDCALL Terminate(MOMIJI_HANDLE h) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    if (!c) return;
    c->opened = 0;
    c->locked = 0;
    c->used = 0;
    c->magic = 0;
}

MOMIJI_BOOL MOMIJI_STDCALL GetRemoteHosts(MOMIJI_HANDLE h, const char* host, MOMIJI_I32 port, void* out, MOMIJI_U32 outBytes, MOMIJI_U32 flags) {
    (void)h; (void)host; (void)port; (void)flags;
    copy_str((char*)out, outBytes, "");
    return 0;
}

MOMIJI_BOOL MOMIJI_STDCALL Open(MOMIJI_HANDLE h, MOMIJI_I32 driveIndexZeroBased) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    if (!c || driveIndexZeroBased < 0 || driveIndexZeroBased > 25) return 0;
    c->driveIndex = driveIndexZeroBased;
    c->opened = 1;
    set_device_from_drive(c, driveIndexZeroBased);
    return 1;
}

MOMIJI_BOOL MOMIJI_STDCALL OpenEx(MOMIJI_HANDLE h, const char* deviceName) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    if (!c || !deviceName) return 0;
    copy_str(c->deviceName, sizeof(c->deviceName), deviceName);
    c->opened = 1;
    return 1;
}

MOMIJI_BOOL MOMIJI_STDCALL Close(MOMIJI_HANDLE h) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    if (!c) return 0;
    c->opened = 0;
    c->locked = 0;
    return 1;
}

MOMIJI_BOOL MOMIJI_STDCALL GetDeviceName(MOMIJI_HANDLE h, MOMIJI_I32 index, char* outName, MOMIJI_U32 outBytes) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    char tmp[16];
    if (!c || !outName || outBytes == 0) return 0;
    if (index >= 0 && index <= 25) {
        tmp[0] = (char)('A' + index);
        tmp[1] = ':';
        tmp[2] = 0;
        copy_str(outName, outBytes, tmp);
    } else {
        copy_str(outName, outBytes, c->deviceName);
    }
    return 1;
}

void MOMIJI_STDCALL ClearTOCStructure(MOMIJI_HANDLE h) { (void)h; }
MOMIJI_U32 MOMIJI_STDCALL GetTOCStructure(MOMIJI_HANDLE h, void* toc, MOMIJI_U32 maxEntries, MOMIJI_U32 flags) { (void)h; (void)toc; (void)maxEntries; (void)flags; return 0; }
MOMIJI_U32 MOMIJI_STDCALL SetTOCStructure(MOMIJI_HANDLE h, const void* toc, MOMIJI_U32 entryCount, MOMIJI_U32 flags) { (void)h; (void)toc; (void)entryCount; (void)flags; return 0; }
MOMIJI_I32 MOMIJI_STDCALL GetFirstLBA(MOMIJI_HANDLE h) { MOMIJI_CONTEXT* c = ctx_from_handle(h); return c ? c->firstLBA : 0; }
MOMIJI_I32 MOMIJI_STDCALL GetLastLBA(MOMIJI_HANDLE h) { MOMIJI_CONTEXT* c = ctx_from_handle(h); return c ? c->lastLBA : 0; }
MOMIJI_U32 MOMIJI_STDCALL SetBufferingBlockLength(MOMIJI_HANDLE h, MOMIJI_U32 blocks) { MOMIJI_CONTEXT* c = ctx_from_handle(h); if (!c) return 0; c->blockLength = blocks; return blocks; }
MOMIJI_U32 MOMIJI_STDCALL ReadStart(MOMIJI_HANDLE h, MOMIJI_I32 startLBA, MOMIJI_U32 blocks) { (void)startLBA; (void)blocks; return ctx_from_handle(h) ? 0 : 0; }
MOMIJI_BOOL MOMIJI_STDCALL ReadLBA(MOMIJI_HANDLE h, MOMIJI_I32 lba, void* buffer) { (void)h; (void)lba; (void)buffer; return 0; }
MOMIJI_BOOL MOMIJI_STDCALL ReadLBADummy(MOMIJI_HANDLE h, MOMIJI_I32 lba, void* buffer) { (void)h; (void)lba; (void)buffer; return 0; }
MOMIJI_BOOL MOMIJI_STDCALL ReadEnd(MOMIJI_HANDLE h) { return ctx_from_handle(h) ? 1 : 0; }
MOMIJI_U32 MOMIJI_STDCALL WriteStart(MOMIJI_HANDLE h, MOMIJI_I32 startLBA, MOMIJI_U32 blocks) { (void)startLBA; (void)blocks; return ctx_from_handle(h) ? 0 : 0; }
MOMIJI_BOOL MOMIJI_STDCALL WriteLBA(MOMIJI_HANDLE h, MOMIJI_I32 lba, const void* buffer, MOMIJI_U32 bytes) { (void)buffer; (void)bytes; MOMIJI_CONTEXT* c = ctx_from_handle(h); if (!c) return 0; c->lastWroteLBA = lba; return 0; }
MOMIJI_BOOL MOMIJI_STDCALL WriteFlush(MOMIJI_HANDLE h, MOMIJI_BOOL sync) { (void)sync; return ctx_from_handle(h) ? 1 : 0; }
MOMIJI_BOOL MOMIJI_STDCALL WriteEnd(MOMIJI_HANDLE h, MOMIJI_BOOL finalizeDisc) { (void)finalizeDisc; return ctx_from_handle(h) ? 1 : 0; }
MOMIJI_I32 MOMIJI_STDCALL GetLastWroteLBA(MOMIJI_HANDLE h) { MOMIJI_CONTEXT* c = ctx_from_handle(h); return c ? c->lastWroteLBA : -1; }
MOMIJI_BOOL MOMIJI_STDCALL LoadTray(MOMIJI_HANDLE h, MOMIJI_BOOL load) { (void)load; return ctx_from_handle(h) ? 1 : 0; }
MOMIJI_BOOL MOMIJI_STDCALL IsReady(MOMIJI_HANDLE h) { MOMIJI_CONTEXT* c = ctx_from_handle(h); return (c && c->opened) ? 1 : 0; }
MOMIJI_U32 MOMIJI_STDCALL GetMediaType(MOMIJI_HANDLE h) { (void)h; return 0; }
MOMIJI_U32 MOMIJI_STDCALL GetMediaFamilyType(MOMIJI_HANDLE h) { (void)h; return 0; }
MOMIJI_BOOL MOMIJI_STDCALL LockUnlock(MOMIJI_HANDLE h, MOMIJI_BOOL lock) { MOMIJI_CONTEXT* c = ctx_from_handle(h); if (!c) return 0; c->locked = lock ? 1 : 0; return 1; }
MOMIJI_BOOL MOMIJI_STDCALL Erase(MOMIJI_HANDLE h, MOMIJI_BOOL fullErase) { (void)h; (void)fullErase; return 0; }
MOMIJI_BOOL MOMIJI_STDCALL IsSupportMedia(MOMIJI_HANDLE h, MOMIJI_U32 mediaType) { (void)h; (void)mediaType; return 0; }
MOMIJI_BOOL MOMIJI_STDCALL IsReadSupport(MOMIJI_HANDLE h, MOMIJI_U32 mediaType) { (void)h; (void)mediaType; return 0; }
MOMIJI_BOOL MOMIJI_STDCALL IsWriteSupport(MOMIJI_HANDLE h, MOMIJI_U32 mediaType) { (void)h; (void)mediaType; return 0; }
MOMIJI_BOOL MOMIJI_STDCALL IsDiscEmpty(MOMIJI_HANDLE h) { (void)h; return 0; }
MOMIJI_U32 MOMIJI_STDCALL ResetSense(MOMIJI_HANDLE h) { (void)h; return 0; }
MOMIJI_BOOL MOMIJI_STDCALL SetECCMode(MOMIJI_HANDLE h, MOMIJI_BOOL enable, MOMIJI_U32 mode) { (void)enable; (void)mode; return ctx_from_handle(h) ? 1 : 0; }
MOMIJI_BOOL MOMIJI_STDCALL SetReadSpeed(MOMIJI_HANDLE h, MOMIJI_U32 speed, MOMIJI_BOOL restoreDefault) { MOMIJI_CONTEXT* c = ctx_from_handle(h); if (!c) return 0; c->readSpeed = restoreDefault ? 0 : speed; return 1; }
MOMIJI_BOOL MOMIJI_STDCALL SetWriteSpeed(MOMIJI_HANDLE h, MOMIJI_U32 speed, MOMIJI_BOOL restoreDefault) { MOMIJI_CONTEXT* c = ctx_from_handle(h); if (!c) return 0; c->writeSpeed = restoreDefault ? 0 : speed; return 1; }
MOMIJI_U32 MOMIJI_STDCALL GetWriteSpeed(MOMIJI_HANDLE h, MOMIJI_U32* current, MOMIJI_U32* maximum) { MOMIJI_CONTEXT* c = ctx_from_handle(h); if (!c) return 0; if (current) *current = c->writeSpeed; if (maximum) *maximum = c->writeSpeed; return c->writeSpeed; }
MOMIJI_BOOL MOMIJI_STDCALL CheckWriteMode(MOMIJI_HANDLE h, MOMIJI_U32 mode, MOMIJI_BOOL testWrite, MOMIJI_U32 flags) { (void)mode; (void)testWrite; (void)flags; return ctx_from_handle(h) ? 0 : 0; }
MOMIJI_U32 MOMIJI_STDCALL GetSCSIErrorStatus(MOMIJI_HANDLE h) { (void)h; return 0; }
MOMIJI_U32 MOMIJI_STDCALL SetWriteCacheBufferSize(MOMIJI_HANDLE h, MOMIJI_U32 bytes) { MOMIJI_CONTEXT* c = ctx_from_handle(h); if (!c) return 0; c->writeCacheBytes = bytes; return bytes; }
MOMIJI_U32 MOMIJI_STDCALL GetUsedWriteCacheBufferSize(MOMIJI_HANDLE h) { MOMIJI_CONTEXT* c = ctx_from_handle(h); return c ? c->writeCacheBytes : 0; }
