#include "common_winapi.h"

#define STDCALL __stdcall
#define MOMIJI_MAGIC 0x4d4f4d32u
#define MAX_CTX 16
#define IOCTL_SCSI_PASS_THROUGH_DIRECT 0x0004D014u
#define SCSI_IOCTL_DATA_OUT 0u
#define SCSI_IOCTL_DATA_IN 1u
#define SCSI_IOCTL_DATA_UNSPECIFIED 2u

#pragma pack(push, 1)
typedef struct SCSI_PASS_THROUGH_DIRECT_TAG {
    u16 Length;
    u8  ScsiStatus;
    u8  PathId;
    u8  TargetId;
    u8  Lun;
    u8  CdbLength;
    u8  SenseInfoLength;
    u8  DataIn;
    u32 DataTransferLength;
    u32 TimeOutValue;
    ptr DataBuffer;
    u32 SenseInfoOffset;
    u8  Cdb[16];
} SCSI_PASS_THROUGH_DIRECT;
#pragma pack(pop)

typedef struct SPTD_WITH_SENSE_TAG {
    SCSI_PASS_THROUGH_DIRECT sptd;
    u32 filler;
    u8 sense[32];
} SPTD_WITH_SENSE;

typedef struct MOMIJI_CONTEXT_TAG {
    u32 magic;
    u32 used;
    u32 opened;
    u32 fileBacked;
    HANDLE32 h;
    char deviceName[260];
    i32 driveIndex;
    u32 blockLength;
    u32 readCommand;
    u32 readFlag;
    u32 writeCommand;
    u32 writeFlag;
    u32 readSpeed;
    u32 writeSpeed;
    u32 writeCacheBytes;
    u32 blockBytes;
    u32 isTestWrite;
    u32 tocTrackCount;
    u32 tocSectorCount[100];
    u32 tocFlags[100];
    u32 lastSense;
    u32 lastScsiStatus;
    u32 mediaProfile;
    i32 firstLBA;
    i32 lastLBA;
    i32 lastWroteLBA;
    u32 eccEnable;
    u32 eccRetry;
    u32 writeSectorBytes;
    u32 reserveSectors;
    u32 tocLen;
    u8  tocBuf[4096];
} MOMIJI_CONTEXT;

static MOMIJI_CONTEXT g_ctx[MAX_CTX];

static MOMIJI_CONTEXT* ctx_from_handle(ptr h) {
    MOMIJI_CONTEXT* c = (MOMIJI_CONTEXT*)h;
    if (!c || c->magic != MOMIJI_MAGIC || !c->used) return (MOMIJI_CONTEXT*)0;
    return c;
}

static void reset_runtime(MOMIJI_CONTEXT* c) {
    c->opened = 0;
    c->fileBacked = 0;
    c->h = K_INVALID_HANDLE;
    c->driveIndex = -1;
    c->blockLength = 1;
    c->readCommand = 0;
    c->readFlag = 0;
    c->writeCommand = 0;
    c->writeFlag = 0;
    c->readSpeed = 0;
    c->writeSpeed = 0;
    c->writeCacheBytes = 0;
    c->blockBytes = 2048;
    c->isTestWrite = 0;
    c->tocTrackCount = 0;
    kureha_zero(c->tocSectorCount, sizeof(c->tocSectorCount));
    kureha_zero(c->tocFlags, sizeof(c->tocFlags));
    c->lastSense = 0;
    c->lastScsiStatus = 0;
    c->mediaProfile = 0;
    c->firstLBA = 0;
    c->lastLBA = 0;
    c->lastWroteLBA = -1;
    c->eccEnable = 0;
    c->eccRetry = 0;
    c->writeSectorBytes = 2048;
    c->reserveSectors = 0;
    c->tocLen = 0;
    kureha_zero(c->tocBuf, sizeof(c->tocBuf));
    kureha_strcopy(c->deviceName, sizeof(c->deviceName), "");
}

static void close_device(MOMIJI_CONTEXT* c) {
    if (c && c->h && c->h != K_INVALID_HANDLE) kureha_CloseHandle(c->h);
    if (c) {
        c->h = K_INVALID_HANDLE;
        c->opened = 0;
        c->fileBacked = 0;
    }
}

static void set_device_name_from_drive(MOMIJI_CONTEXT* c, int drive) {
    c->deviceName[0] = '\\';
    c->deviceName[1] = '\\';
    c->deviceName[2] = '.';
    c->deviceName[3] = '\\';
    c->deviceName[4] = (char)('A' + (drive & 31));
    c->deviceName[5] = ':';
    c->deviceName[6] = 0;
}

static int is_device_path(const char* s) {
    return kureha_starts_with_i(s, "\\\\.\\") || kureha_starts_with_i(s, "\\\\?\\");
}

static int open_path(MOMIJI_CONTEXT* c, const char* path) {
    HANDLE32 h;
    close_device(c);
    if (!path) return 0;
    kureha_strcopy(c->deviceName, sizeof(c->deviceName), path);
    h = kureha_CreateFileA(path, K_GENERIC_READ | K_GENERIC_WRITE,
                           K_FILE_SHARE_READ | K_FILE_SHARE_WRITE, K_OPEN_EXISTING);
    if (h == K_INVALID_HANDLE && !is_device_path(path)) {
        h = kureha_CreateFileA(path, K_GENERIC_READ | K_GENERIC_WRITE,
                               K_FILE_SHARE_READ | K_FILE_SHARE_WRITE, K_CREATE_ALWAYS);
    }
    if (h == K_INVALID_HANDLE) {
        h = kureha_CreateFileA(path, K_GENERIC_READ,
                               K_FILE_SHARE_READ | K_FILE_SHARE_WRITE, K_OPEN_EXISTING);
    }
    if (h == K_INVALID_HANDLE) {
        c->h = K_INVALID_HANDLE;
        c->opened = 0;
        return 0;
    }
    c->h = h;
    c->opened = 1;
    c->fileBacked = is_device_path(path) ? 0u : 1u;
    c->lastSense = 0;
    c->lastScsiStatus = 0;
    return 1;
}

static int send_cdb(MOMIJI_CONTEXT* c, const u8* cdb, u32 cdbLen, u32 dataIn, ptr buf, u32 bytes, u32 timeoutSeconds) {
    SPTD_WITH_SENSE swb;
    u32 returned = 0;
    int ok;
    if (!c || !c->opened || c->fileBacked || !c->h || c->h == K_INVALID_HANDLE) return 0;
    kureha_zero(&swb, sizeof(swb));
    swb.sptd.Length = (u16)sizeof(SCSI_PASS_THROUGH_DIRECT);
    swb.sptd.CdbLength = (u8)cdbLen;
    swb.sptd.SenseInfoLength = sizeof(swb.sense);
    swb.sptd.DataIn = (u8)dataIn;
    swb.sptd.DataTransferLength = bytes;
    swb.sptd.TimeOutValue = timeoutSeconds ? timeoutSeconds : 20;
    swb.sptd.DataBuffer = buf;
    swb.sptd.SenseInfoOffset = (u32)((u8*)swb.sense - (u8*)&swb);
    kureha_copy(swb.sptd.Cdb, cdb, cdbLen);
    ok = kureha_DeviceIoControl(c->h, IOCTL_SCSI_PASS_THROUGH_DIRECT, &swb, sizeof(swb), &swb, sizeof(swb), &returned) ? 1 : 0;
    c->lastScsiStatus = swb.sptd.ScsiStatus;
    c->lastSense = ((u32)(swb.sense[2] & 0x0f) << 16) | ((u32)swb.sense[12] << 8) | swb.sense[13];
    return ok && swb.sptd.ScsiStatus == 0;
}

static int test_unit_ready(MOMIJI_CONTEXT* c) {
    u8 cdb[6];
    kureha_zero(cdb, sizeof(cdb));
    cdb[0] = 0x00;
    return send_cdb(c, cdb, 6, SCSI_IOCTL_DATA_UNSPECIFIED, (ptr)0, 0, 10);
}

static int update_capacity(MOMIJI_CONTEXT* c) {
    u8 cdb[10];
    u8 data[8];
    if (!c || !c->opened) return 0;
    if (c->fileBacked) {
        u32 hi = 0;
        u32 sz = kureha_GetFileSize(c->h, &hi);
        c->firstLBA = 0;
        if (sz == 0xffffffffu || hi != 0 || sz == 0) c->lastLBA = 0;
        else c->lastLBA = (i32)((sz - 1u) / 2048u);
        return 1;
    }
    kureha_zero(cdb, sizeof(cdb));
    kureha_zero(data, sizeof(data));
    cdb[0] = 0x25;
    if (!send_cdb(c, cdb, 10, SCSI_IOCTL_DATA_IN, data, sizeof(data), 20)) return 0;
    c->firstLBA = 0;
    c->lastLBA = (i32)kureha_be32(data);
    return 1;
}

static u32 update_profile(MOMIJI_CONTEXT* c) {
    u8 cdb[10];
    u8 data[32];
    if (!c || !c->opened || c->fileBacked) return c ? c->mediaProfile : 0;
    kureha_zero(cdb, sizeof(cdb));
    kureha_zero(data, sizeof(data));
    cdb[0] = 0x46;        /* GET CONFIGURATION */
    cdb[1] = 0x02;        /* current only */
    cdb[8] = sizeof(data);
    if (send_cdb(c, cdb, 10, SCSI_IOCTL_DATA_IN, data, sizeof(data), 20)) {
        c->mediaProfile = kureha_be16(data + 6);
    }
    return c->mediaProfile;
}

static u32 profile_family(u32 p) {
    if ((p >= 0x0008 && p <= 0x000a) || p == 0x0002) return 1; /* CD */
    if (p >= 0x0010 && p <= 0x002b) return 2;                  /* DVD */
    if (p >= 0x0040 && p <= 0x0043) return 3;                  /* BD */
    if (p >= 0x0050 && p <= 0x0052) return 4;                  /* HD DVD */
    return 0;
}

static u32 read_sector_size(MOMIJI_CONTEXT* c) {
    u32 fam = profile_family(c ? c->mediaProfile : 0);
    if (fam == 1) {
        if (c->readCommand == 2) return 2448;       /* RAW + SUB96 */
        if (c->readCommand == 1) return 2368;       /* RAW + SUB16 */
        if (c->readCommand == 0) return 2352;       /* RAW */
    }
    return 2048;
}

static int read10(MOMIJI_CONTEXT* c, u32 lba, ptr buffer) {
    u8 cdb[10];
    if (!buffer) return 0;
    kureha_zero(cdb, sizeof(cdb));
    cdb[0] = 0x28;
    kureha_put_be32(cdb + 2, lba);
    kureha_put_be16(cdb + 7, 1);
    return send_cdb(c, cdb, 10, SCSI_IOCTL_DATA_IN, buffer, 2048, 30);
}

static int read_cd(MOMIJI_CONTEXT* c, u32 lba, ptr buffer, u32 bytes) {
    u8 cdb[12];
    if (!buffer) return 0;
    kureha_zero(cdb, sizeof(cdb));
    cdb[0] = 0xbe; /* READ CD */
    cdb[1] = 0x00; /* any sector type */
    kureha_put_be32(cdb + 2, lba);
    kureha_put_be24(cdb + 6, 1);
    if (bytes >= 2352) cdb[9] = 0xf8; /* sync + headers + user + EDC/ECC */
    else cdb[9] = 0x10;               /* user data only */
    if (bytes == 2368) cdb[10] = 0x02;
    else if (bytes >= 2448) cdb[10] = 0x01;
    return send_cdb(c, cdb, 12, SCSI_IOCTL_DATA_IN, buffer, bytes, 30);
}

static u32 write_sector_size_from_type(u32 dataType) {
    if (dataType == 0 || dataType == 8 || dataType == 2048) return 2048;
    if (dataType == 1 || dataType == 2352) return 2352;
    if (dataType == 3 || dataType == 2368) return 2368;
    if (dataType == 4 || dataType == 2448) return 2448;
    return 2048;
}

static u8 write_type_from_command(u32 writeCommand) {
    if (writeCommand == 5 || writeCommand == 8) return 0x02; /* SAO/DVD DAO style */
    if (writeCommand == 3 || writeCommand == 4) return 0x02; /* DAO raw */
    if (writeCommand == 1) return 0x02;                     /* SAO raw */
    return 0x02;
}

static u8 data_block_type_from_bytes(u32 bytes) {
    if (bytes >= 2448) return 0x03; /* raw PW 96 */
    if (bytes >= 2368) return 0x02; /* raw packed PW / 16-ish */
    if (bytes >= 2352) return 0x01; /* raw 2352 */
    return 0x08;                    /* Mode 1 2048 */
}

static u32 bytes_from_write_command(u32 writeCommand) {
    if (writeCommand == 1) return 2352;
    if (writeCommand == 3) return 2368;
    if (writeCommand == 4) return 2448;
    return 2048;
}

/* Older write-parameter helper variants removed; the active writer setup starts at guess_sector_bytes_from_write_command(). */

static int write10(MOMIJI_CONTEXT* c, u32 lba, const void* buffer, u32 bytes) {
    u8 cdb[10];
    if (!c || !buffer || bytes < 2048 || bytes > 2448) return 0;
    if (c->fileBacked) {
        i32 hi = 0;
        u32 wrote = 0;
        kureha_SetFilePointer(c->h, (i32)(lba * bytes), &hi, K_FILE_BEGIN);
        return kureha_WriteFile(c->h, buffer, bytes, &wrote) && wrote == bytes;
    }
    kureha_zero(cdb, sizeof(cdb));
    cdb[0] = 0x2a;
    kureha_put_be32(cdb + 2, lba);
    kureha_put_be16(cdb + 7, 1);
    return send_cdb(c, cdb, 10, SCSI_IOCTL_DATA_OUT, (ptr)buffer, bytes, 120);
}

static int write12_raw(MOMIJI_CONTEXT* c, u32 lba, const void* buffer, u32 bytes) {
    u8 cdb[12];
    if (!c || !buffer || bytes == 0) return 0;
    if (c->fileBacked) {
        i32 hi = 0;
        u32 wrote = 0;
        u32 posBytes = lba * bytes;
        kureha_SetFilePointer(c->h, (i32)posBytes, &hi, K_FILE_BEGIN);
        return kureha_WriteFile(c->h, buffer, bytes, &wrote) && wrote == bytes;
    }
    kureha_zero(cdb, sizeof(cdb));
    cdb[0] = 0xaa; /* WRITE(12), used here for non-2048 experimental block sizes. */
    kureha_put_be32(cdb + 2, lba);
    kureha_put_be32(cdb + 6, 1);
    return send_cdb(c, cdb, 12, SCSI_IOCTL_DATA_OUT, (ptr)buffer, bytes, 120);
}

static int read_toc_raw(MOMIJI_CONTEXT* c, u32 format, ptr outBuffer, u32 outBytes) {
    u8 cdb[10];
    if (!c || !outBuffer || outBytes == 0 || c->fileBacked) return 0;
    if (outBytes > 0xfffeu) outBytes = 0xfffeu;
    kureha_zero(cdb, sizeof(cdb));
    cdb[0] = 0x43;                 /* READ TOC/PMA/ATIP */
    cdb[2] = (u8)(format & 0x0f);  /* format: 0 formatted TOC, 1 multi-session, 2 raw TOC, ... */
    cdb[6] = 1;                    /* starting track/session */
    kureha_put_be16(cdb + 7, outBytes);
    return send_cdb(c, cdb, 10, SCSI_IOCTL_DATA_IN, outBuffer, outBytes, 30);
}

static u32 guess_sector_bytes_from_write_command(u32 writeCommand, u32 writeFlag) {
    (void)writeFlag;
    if (writeCommand == 2048 || writeCommand == 8 || writeCommand == 5 || writeCommand == 0) return 2048;
    if (writeCommand == 2352 || writeCommand == 1 || writeCommand == 2) return 2352;
    if (writeCommand == 2368 || writeCommand == 3) return 2368;
    if (writeCommand == 2448 || writeCommand == 4 || writeCommand == 96) return 2448;
    return 2048;
}

static int mode_select_write_params(MOMIJI_CONTEXT* c, u32 sectorBytes, u32 writeCommand, u32 writeFlag) {
    u8 cdb[10];
    u8 page[64];
    u32 dataBlockType = 8; /* MMC block type 8 is commonly Mode1/2048. */
    u32 writeType = 0;     /* packet/incremental default */
    if (!c || c->fileBacked) return 1;

    if (sectorBytes >= 2448) dataBlockType = 3;
    else if (sectorBytes >= 2368) dataBlockType = 2;
    else if (sectorBytes >= 2352) dataBlockType = 0;
    if (writeCommand == 1 || writeCommand == 2 || writeCommand == 2352) writeType = 2; /* SAO-like */
    if (writeCommand == 3 || writeCommand == 4 || writeCommand == 96 || sectorBytes >= 2368) writeType = 3; /* RAW-like */
    if (writeFlag & 1u) writeType |= 0x10u; /* Test-write bit in many MMC implementations. */

    kureha_zero(page, sizeof(page));
    /* MODE SELECT(10) parameter list: 8-byte header + Write Parameters mode page 05h. */
    page[8] = 0x05;
    page[9] = 0x32;
    page[10] = (u8)((writeType & 0x1f) | 0x40u); /* BUFE + write type/test bit */
    page[11] = 0x04;             /* track mode: data, copy permitted */
    page[12] = (u8)(dataBlockType & 0x0f);
    page[13] = 0;                /* link size */
    page[14] = 0; page[15] = 0; page[16] = 0; page[17] = 0; /* packet size = 0 for SAO/TAO */
    page[18] = 0; page[19] = 150; /* audio pause length = 150 frames */

    kureha_zero(cdb, sizeof(cdb));
    cdb[0] = 0x55;               /* MODE SELECT(10) */
    cdb[1] = 0x10;               /* PF bit */
    kureha_put_be16(cdb + 7, sizeof(page));
    return send_cdb(c, cdb, 10, SCSI_IOCTL_DATA_OUT, page, sizeof(page), 30);
}

static int send_opc_information(MOMIJI_CONTEXT* c, u32 doOpc) {
    u8 cdb[10];
    if (!c || c->fileBacked) return 1;
    kureha_zero(cdb, sizeof(cdb));
    cdb[0] = 0x54; /* SEND OPC INFORMATION */
    if (doOpc) cdb[1] = 0x01;
    return send_cdb(c, cdb, 10, SCSI_IOCTL_DATA_UNSPECIFIED, (ptr)0, 0, 120);
}

static int reserve_track(MOMIJI_CONTEXT* c, u32 sectors) {
    u8 cdb[10];
    if (!c || c->fileBacked || sectors == 0) return 1;
    kureha_zero(cdb, sizeof(cdb));
    cdb[0] = 0x53; /* RESERVE TRACK */
    kureha_put_be32(cdb + 5, sectors);
    return send_cdb(c, cdb, 10, SCSI_IOCTL_DATA_UNSPECIFIED, (ptr)0, 0, 60);
}

static int send_cue_sheet_if_present(MOMIJI_CONTEXT* c) {
    u8 cdb[10];
    if (!c || c->fileBacked || c->tocLen < 8 || c->tocLen > sizeof(c->tocBuf)) return 1;
    /* Only send if the buffer looks like an MMC cue sheet: 8-byte records and <= 99 tracks. */
    if ((c->tocLen & 7u) != 0) return 1;
    kureha_zero(cdb, sizeof(cdb));
    cdb[0] = 0x5d; /* SEND CUE SHEET */
    kureha_put_be24(cdb + 6, c->tocLen);
    return send_cdb(c, cdb, 10, SCSI_IOCTL_DATA_OUT, c->tocBuf, c->tocLen, 60);
}

static int close_track_session(MOMIJI_CONTEXT* c, u32 closeSession) {
    u8 cdb[12];
    if (!c || c->fileBacked) return 1;
    kureha_zero(cdb, sizeof(cdb));
    cdb[0] = 0x5b; /* CLOSE TRACK/SESSION */
    cdb[2] = closeSession ? 0x02 : 0x01;
    return send_cdb(c, cdb, 12, SCSI_IOCTL_DATA_UNSPECIFIED, (ptr)0, 0, 600);
}

u32 STDCALL GetEngineVersion(void) { return 10601u; }

ptr STDCALL Initialize(void) {
    int i;
    for (i = 0; i < MAX_CTX; ++i) {
        if (!g_ctx[i].used) {
            MOMIJI_CONTEXT* c = &g_ctx[i];
            kureha_zero(c, sizeof(*c));
            c->magic = MOMIJI_MAGIC;
            c->used = 1;
            reset_runtime(c);
            return (ptr)c;
        }
    }
    return (ptr)0;
}

void STDCALL Terminate(ptr h) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    if (!c) return;
    close_device(c);
    kureha_zero(c, sizeof(*c));
}

/* Best-effort stub for original REMOTESCSI discovery.  The original used Winsock; this remake does not. */
u32 STDCALL GetRemoteHosts(ptr h, const char* broadcastIP, i32 port, u32 outBytes, u32 timeoutMs, char* outBuffer) {
    (void)h; (void)broadcastIP; (void)port; (void)timeoutMs;
    if (outBuffer && outBytes) outBuffer[0] = 0;
    return 0;
}

u32 STDCALL Open(ptr h, i32 driveIndexZeroBased) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    if (!c || driveIndexZeroBased < 0 || driveIndexZeroBased > 25) return 0;
    c->driveIndex = driveIndexZeroBased;
    set_device_name_from_drive(c, driveIndexZeroBased);
    return open_path(c, c->deviceName);
}

u32 STDCALL OpenEx(ptr h, const char* deviceName) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    char tmp[16];
    if (!c || !deviceName) return 0;
    if (deviceName[0] && deviceName[1] == ':' && deviceName[2] == 0) {
        tmp[0]='\\'; tmp[1]='\\'; tmp[2]='.'; tmp[3]='\\'; tmp[4]=deviceName[0]; tmp[5]=':'; tmp[6]=0;
        return open_path(c, tmp);
    }
    return open_path(c, deviceName);
}

u32 STDCALL Close(ptr h) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    if (!c) return 0;
    close_device(c);
    return 1;
}

u32 STDCALL GetDeviceName(ptr h, i32 nameType, char* outName, u32 outBytes) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    char tmp[16];
    if (!c || !outName || outBytes == 0) return 0;
    if (nameType >= 0 && nameType <= 25) {
        tmp[0] = (char)('A' + nameType);
        tmp[1] = ':';
        tmp[2] = 0;
        kureha_strcopy(outName, outBytes, tmp);
    } else if (c->deviceName[0]) {
        kureha_strcopy(outName, outBytes, c->deviceName);
    } else {
        kureha_strcopy(outName, outBytes, "");
    }
    return 1;
}

void STDCALL ClearTOCStructure(ptr h) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    if (!c) return;
    c->firstLBA = 0;
    c->lastLBA = 0;
}

u32 STDCALL GetTOCStructure(ptr h, u32 tocType, void* outBuffer, u32 outBytes) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    if (!c) return 0;
    if (outBuffer && outBytes) kureha_zero(outBuffer, outBytes);
    if (c->opened && !c->fileBacked && outBuffer && outBytes >= 4) {
        if (read_toc_raw(c, tocType, outBuffer, outBytes)) return 1;
    }
    update_capacity(c);
    if (outBuffer && outBytes >= 8) {
        ((u32*)outBuffer)[0] = 1;                  /* fallback: one data track */
        ((u32*)outBuffer)[1] = (u32)c->lastLBA;
        return 1;
    }
    return c->lastLBA ? 1u : 0u;
}

u32 STDCALL SetTOCStructure(ptr h, u32 tocType, const void* inBuffer, u32 inBytes) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    u32 i, tracks, sectors;
    (void)tocType;
    if (!c) return 0;
    c->tocLen = 0;
    c->reserveSectors = 0;
    if (inBuffer && inBytes) {
        if (inBytes > sizeof(c->tocBuf)) inBytes = sizeof(c->tocBuf);
        kureha_copy(c->tocBuf, inBuffer, inBytes);
        c->tocLen = inBytes;
        if (inBytes >= 8) c->lastLBA = (i32)((const u32*)inBuffer)[1];
        /* Clean Zenki fallback layout: DWORD trackCount, reserved, then repeated track,pregap,sectors,flags. */
        tracks = ((const u32*)inBuffer)[0];
        if (tracks > 0 && tracks < 100 && inBytes >= (2 + tracks * 4) * 4u) {
            sectors = 0;
            for (i = 0; i < tracks; ++i) sectors += ((const u32*)inBuffer)[2 + i * 4 + 1] + ((const u32*)inBuffer)[2 + i * 4 + 2];
            c->reserveSectors = sectors;
        } else if ((inBytes & 7u) == 0 && inBytes >= 24) {
            /* MMC cue sheet: estimate lead-out from last A2 record if present. */
            const u8* b = (const u8*)inBuffer;
            for (i = 0; i + 7 < inBytes; i += 8) {
                if (b[i + 1] == 0xaa) {
                    u32 mm = b[i + 5], ss = b[i + 6], ff = b[i + 7];
                    u32 lba = mm * 60u * 75u + ss * 75u + ff;
                    c->reserveSectors = lba > 150 ? lba - 150 : lba;
                }
            }
        }
    }
    return 1;
}

i32 STDCALL GetFirstLBA(ptr h) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    return c ? c->firstLBA : 0;
}

i32 STDCALL GetLastLBA(ptr h) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    if (!c) return 0;
    update_capacity(c);
    return c->lastLBA;
}

u32 STDCALL SetBufferingBlockLength(ptr h, u32 blocks) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    if (!c) return 0;
    if (blocks < 1) blocks = 1;
    if (blocks > 26) blocks = 26;
    c->blockLength = blocks;
    return blocks;
}

u32 STDCALL ReadStart(ptr h, u32 readCommand, u32 readFlag) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    if (!c || !c->opened) return 0;
    c->readCommand = readCommand;
    c->readFlag = readFlag;
    update_profile(c);
    update_capacity(c);
    return 1;
}

u32 STDCALL ReadLBA(ptr h, void* buffer, i32 lba) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    u32 sz;
    if (!c || !c->opened || !buffer || lba < 0) return 0;
    if (c->fileBacked) {
        i32 hi = 0;
        u32 got = 0;
        kureha_zero(buffer, 2448);
        kureha_SetFilePointer(c->h, (i32)((u32)lba * 2048u), &hi, K_FILE_BEGIN);
        return kureha_ReadFile(c->h, buffer, 2048, &got) && got > 0;
    }
    sz = read_sector_size(c);
    kureha_zero(buffer, sz < 2448 ? 2448 : sz);
    if (sz > 2048 && read_cd(c, (u32)lba, buffer, sz)) return 1;
    if (read10(c, (u32)lba, buffer)) return 1;
    return 0;
}

u32 STDCALL ReadLBADummy(ptr h, void* buffer, i32 lba) {
    (void)lba;
    if (!ctx_from_handle(h) || !buffer) return 0;
    kureha_zero(buffer, 2448);
    return 1;
}

u32 STDCALL ReadEnd(ptr h) {
    return ctx_from_handle(h) ? 1u : 0u;
}

u32 STDCALL WriteStart(ptr h, u32 writeCommand, u32 writeFlag) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    if (!c || !c->opened) return 0;
    c->writeCommand = writeCommand;
    c->writeFlag = writeFlag;
    c->writeSectorBytes = guess_sector_bytes_from_write_command(writeCommand, writeFlag);
    c->lastWroteLBA = -1;
    if (c->fileBacked) return 1;
    update_profile(c);
    /* Experimental physical write path: prepare write parameters, reserve if size is known, then send cue sheet if the VB side provided one. */
    mode_select_write_params(c, c->writeSectorBytes, writeCommand, writeFlag);
    send_opc_information(c, 1);
    reserve_track(c, c->reserveSectors);
    send_cue_sheet_if_present(c);
    return 1;
}

u32 STDCALL WriteLBA(ptr h, u32 dataType, const void* buffer, i32 lba) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    u32 sz = write_sector_size_from_type(dataType);
    if (!c || !c->opened || !buffer || lba < 0) return 0;
    /* WRITE(10) is the normal MMC path for all current block types after MODE SELECT. */
    if (write10(c, (u32)lba, buffer, sz)) { c->lastWroteLBA = lba; return 1; }
    /* Some bridges expose raw-sized blocks through WRITE(12); keep it as a fallback. */
    if (sz != 2048 && write12_raw(c, (u32)lba, buffer, sz)) {
        c->lastWroteLBA = lba;
        return 1;
    }
    return 0;
}

u32 STDCALL WriteFlush(ptr h, u32 abortFlag) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    u8 cdb[10];
    (void)abortFlag;
    if (!c || !c->opened) return 0;
    if (c->fileBacked) return kureha_FlushFileBuffers(c->h) ? 1u : 0u;
    kureha_zero(cdb, sizeof(cdb));
    cdb[0] = 0x35; /* SYNCHRONIZE CACHE */
    return send_cdb(c, cdb, 10, SCSI_IOCTL_DATA_UNSPECIFIED, (ptr)0, 0, 120) ? 1u : 0u;
}

u32 STDCALL WriteEnd(ptr h, u32 abortFlag) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    if (!c || !c->opened) return 0;
    if (!abortFlag) {
        WriteFlush(h, 0);
        close_track_session(c, 1);
    }
    return 1;
}

i32 STDCALL GetLastWroteLBA(ptr h) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    return c ? c->lastWroteLBA : -1;
}

u32 STDCALL LoadTray(ptr h, u32 load) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    u8 cdb[6];
    if (!c || !c->opened) return 0;
    if (c->fileBacked) return 1;
    kureha_zero(cdb, sizeof(cdb));
    cdb[0] = 0x1b;
    cdb[4] = load ? 0x03 : 0x02;
    return send_cdb(c, cdb, 6, SCSI_IOCTL_DATA_UNSPECIFIED, (ptr)0, 0, 30) ? 1u : 0u;
}

u32 STDCALL IsReady(ptr h) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    if (!c || !c->opened) return 0;
    if (c->fileBacked) return 1;
    return test_unit_ready(c) ? 1u : 0u;
}

u32 STDCALL GetMediaType(ptr h) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    if (!c) return 0;
    return update_profile(c);
}

u32 STDCALL GetMediaFamilyType(ptr h) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    if (!c) return 0;
    return profile_family(update_profile(c));
}

u32 STDCALL LockUnlock(ptr h, u32 lock) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    u8 cdb[6];
    if (!c || !c->opened) return 0;
    if (c->fileBacked) return 1;
    kureha_zero(cdb, sizeof(cdb));
    cdb[0] = 0x1e;
    cdb[4] = lock ? 1 : 0;
    return send_cdb(c, cdb, 6, SCSI_IOCTL_DATA_UNSPECIFIED, (ptr)0, 0, 20) ? 1u : 0u;
}

u32 STDCALL Erase(ptr h, u32 quickly) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    u8 cdb[12];
    if (!c || !c->opened || c->fileBacked) return 0;
    kureha_zero(cdb, sizeof(cdb));
    cdb[0] = 0xa1;              /* BLANK, mainly CD-RW */
    cdb[1] = quickly ? 0x01 : 0x00;
    return send_cdb(c, cdb, 12, SCSI_IOCTL_DATA_UNSPECIFIED, (ptr)0, 0, quickly ? 180 : 3600) ? 1u : 0u;
}

u32 STDCALL IsSupportMedia(ptr h, u32 mediaType) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    (void)mediaType;
    return (c && c->opened) ? 1u : 0u;
}

u32 STDCALL IsReadSupport(ptr h, u32 mediaFamily) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    (void)mediaFamily;
    return (c && c->opened) ? 1u : 0u;
}

u32 STDCALL IsWriteSupport(ptr h, u32 mediaFamily) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    (void)mediaFamily;
    return (c && c->opened) ? 1u : 0u;
}

u32 STDCALL IsDiscEmpty(ptr h) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    u8 cdb[10];
    u8 data[32];
    if (!c || !c->opened || c->fileBacked) return 0;
    kureha_zero(cdb, sizeof(cdb));
    kureha_zero(data, sizeof(data));
    cdb[0] = 0x51; /* READ DISC INFORMATION */
    cdb[8] = sizeof(data);
    if (!send_cdb(c, cdb, 10, SCSI_IOCTL_DATA_IN, data, sizeof(data), 20)) return 0;
    return ((data[2] & 0x03) == 0) ? 1u : 0u;
}

u32 STDCALL ResetSense(ptr h) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    if (!c) return 0;
    c->lastSense = 0;
    c->lastScsiStatus = 0;
    return 1;
}

u32 STDCALL SetECCMode(ptr h, u32 enable, u32 retryCount) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    if (!c) return 0;
    c->eccEnable = enable ? 1u : 0u;
    c->eccRetry = retryCount;
    return 1;
}

u32 STDCALL SetReadSpeed(ptr h, u32 speed, u32 cav) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    u8 cdb[12];
    (void)cav;
    if (!c) return 0;
    c->readSpeed = speed;
    if (!c->opened || c->fileBacked) return 1;
    kureha_zero(cdb, sizeof(cdb));
    cdb[0] = 0xbb; /* SET CD SPEED */
    kureha_put_be16(cdb + 2, speed ? speed : 0xffffu);
    kureha_put_be16(cdb + 4, 0xffffu);
    send_cdb(c, cdb, 12, SCSI_IOCTL_DATA_UNSPECIFIED, (ptr)0, 0, 20);
    return 1;
}

u32 STDCALL SetWriteSpeed(ptr h, u32 speed, u32 cav) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    u8 cdb[12];
    (void)cav;
    if (!c) return 0;
    c->writeSpeed = speed;
    if (!c->opened || c->fileBacked) return 1;
    kureha_zero(cdb, sizeof(cdb));
    cdb[0] = 0xbb;
    kureha_put_be16(cdb + 2, 0xffffu);
    kureha_put_be16(cdb + 4, speed ? speed : 0xffffu);
    send_cdb(c, cdb, 12, SCSI_IOCTL_DATA_UNSPECIFIED, (ptr)0, 0, 20);
    return 1;
}

/* Original ABI: GetWriteSpeed(ctx, outBuffer, maxCount) returns number of DWORD entries. */
u32 STDCALL GetWriteSpeed(ptr h, u32* outSpeeds, u32 maxCount) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    u8 cdb[12];
    u8 data[256];
    u32 count = 0;
    if (!c || !outSpeeds || maxCount == 0) return 0;
    if (c->opened && !c->fileBacked) {
        kureha_zero(cdb, sizeof(cdb));
        kureha_zero(data, sizeof(data));
        cdb[0] = 0xac;              /* GET PERFORMANCE */
        cdb[1] = 0x00;
        cdb[8] = 0x10;              /* maximum number of descriptors requested */
        cdb[10] = 0x03;             /* write speed descriptors */
        if (send_cdb(c, cdb, 12, SCSI_IOCTL_DATA_IN, data, sizeof(data), 20)) {
            u32 len = kureha_be32(data);
            u32 pos = 8;
            if (len + 4 > sizeof(data)) len = sizeof(data) - 4;
            while (pos + 16 <= len + 4 && count < maxCount) {
                u32 sp = kureha_be32(data + pos + 12);
                if (sp) outSpeeds[count++] = sp;
                pos += 16;
            }
        }
    }
    if (count == 0) {
        outSpeeds[count++] = c->writeSpeed ? c->writeSpeed : 1764u;
    }
    return count;
}

u32 STDCALL CheckWriteMode(ptr h, u32 writeCommand, u32 testMode, u32 mediaFamilyType) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    (void)testMode; (void)mediaFamilyType;
    if (!c || !c->opened) return 0;
    /* v3 deliberately reports the experimental raw modes as usable so the clean VB6 UI can exercise them on test media. */
    if (writeCommand == 0 || writeCommand == 1 || writeCommand == 2 || writeCommand == 3 || writeCommand == 4 ||
        writeCommand == 5 || writeCommand == 8 || writeCommand == 96 || writeCommand == 2048 ||
        writeCommand == 2352 || writeCommand == 2368 || writeCommand == 2448) return 1;
    return c->fileBacked ? 1u : 0u;
}

u32 STDCALL GetSCSIErrorStatus(ptr h) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    return c ? ((c->lastScsiStatus << 24) | (c->lastSense & 0x00ffffffu)) : 0;
}

u32 STDCALL SetWriteCacheBufferSize(ptr h, u32 bytes) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    if (!c) return 0;
    c->writeCacheBytes = bytes;
    return bytes;
}

u32 STDCALL GetUsedWriteCacheBufferSize(ptr h) {
    MOMIJI_CONTEXT* c = ctx_from_handle(h);
    return c ? c->writeCacheBytes : 0;
}
