#include "common_winapi.h"

#define STDCALL __stdcall
#define ZENKI_MAGIC 0x5a454e33u
#define MAX_CTX 8
#define MAX_TRACKS 99
#define MAX_ISO_ITEMS 512
#define MAX_LANG 8
#define MAX_INFO 16
#define ISO_ROOT_PARENT 0xffffffffu

void* memcpy(void* dst, const void* src, u32 n) {
    u8* d = (u8*)dst;
    const u8* s = (const u8*)src;
    while (n--) *d++ = *s++;
    return dst;
}

void* memset(void* dst, int value, u32 n) {
    u8* d = (u8*)dst;
    while (n--) *d++ = (u8)value;
    return dst;
}

typedef struct ZENKI_TRACK_TAG {
    u32 used;
    char path[260];
    u32 pregap;
    u32 postgap;
    u32 flag;
    u32 sectorSize;
    u32 sectorCount;
    u32 fileBytes;
    u32 dataOffset;
    u32 dataBytes;
    u32 isWave;
} ZENKI_TRACK;

typedef struct ZENKI_ISO_ITEM_TAG {
    u32 used;
    u32 isDir;
    u32 attr;
    u32 stamp;
    u32 sizeBytes;
    u32 extent;
    u32 sectorCount;
    u32 parent;
    u32 dirBytes;
    u32 dirSectors;
    char name[260];
    char original[260];
    char isoName[96];
} ZENKI_ISO_ITEM;

typedef struct ZENKI_CONTEXT_TAG {
    u32 magic;
    u32 used;
    u32 isoFlags;
    u32 isoCount;
    u32 isoEnumIndex;
    char currentDir[260];
    char lastIsoName[260];
    ZENKI_ISO_ITEM isoItems[MAX_ISO_ITEMS];

    u32 isoLayoutValid;
    u32 isoPvdSector;
    u32 isoVdstSector;
    u32 isoLPathSector;
    u32 isoMPathSector;
    u32 isoPathTableBytes;
    u32 isoPathTableSectors;
    u32 isoRootSector;
    u32 isoRootBytes;
    u32 isoRootSectors;
    u32 isoDataStart;
    u32 isoVolumeSectors;
    u32 isoReadSector;

    u32 trackCount;
    ZENKI_TRACK tracks[MAX_TRACKS];

    u32 trackTextEnabled;
    u32 textEnabled[MAX_LANG][MAX_INFO];
    u32 textUsed[100][MAX_LANG][MAX_INFO];
    char text[100][MAX_LANG][MAX_INFO][128];

    HANDLE32 readHandle;
    u32 reading;
    u32 readMode; /* 1=track file stream, 2=generated ISO image */
    u32 readTrack;
    u32 readPhase;
    u32 readSectorInPhase;
    u32 readByteInData;
} ZENKI_CONTEXT;

static ZENKI_CONTEXT g_ctx[MAX_CTX];

static ZENKI_CONTEXT* ctx_from_handle(ptr h) {
    ZENKI_CONTEXT* c = (ZENKI_CONTEXT*)h;
    if (!c || c->magic != ZENKI_MAGIC || !c->used) return (ZENKI_CONTEXT*)0;
    return c;
}

static u32 min_u32(u32 a, u32 b) { return a < b ? a : b; }
static u32 align2048(u32 v) { return (v + 2047u) & ~2047u; }
static u32 divceil2048(u32 v) { return (v + 2047u) / 2048u; }

static void put_le16(u8* p, u32 v) { p[0] = (u8)v; p[1] = (u8)(v >> 8); }
static void put_le32(u8* p, u32 v) { p[0] = (u8)v; p[1] = (u8)(v >> 8); p[2] = (u8)(v >> 16); p[3] = (u8)(v >> 24); }
static void put_both16(u8* p, u32 v) { put_le16(p, v); kureha_put_be16(p + 2, v); }
static void put_both32(u8* p, u32 v) { put_le32(p, v); kureha_put_be32(p + 4, v); }

static void close_read(ZENKI_CONTEXT* c) {
    if (c && c->readHandle && c->readHandle != K_INVALID_HANDLE) kureha_CloseHandle(c->readHandle);
    if (c) {
        c->readHandle = K_INVALID_HANDLE;
        c->reading = 0;
        c->readMode = 0;
    }
}

static void clear_tracks(ZENKI_CONTEXT* c) {
    u32 i;
    if (!c) return;
    c->trackCount = 0;
    for (i = 0; i < MAX_TRACKS; ++i) kureha_zero(&c->tracks[i], sizeof(c->tracks[i]));
}

static void clear_iso(ZENKI_CONTEXT* c) {
    u32 i;
    if (!c) return;
    c->isoCount = 0;
    c->isoEnumIndex = 0;
    c->isoLayoutValid = 0;
    c->isoPvdSector = 16;
    c->isoVdstSector = 17;
    c->isoLPathSector = 18;
    c->isoMPathSector = 19;
    c->isoPathTableBytes = 10;
    c->isoPathTableSectors = 1;
    c->isoRootSector = 20;
    c->isoRootBytes = 0;
    c->isoRootSectors = 1;
    c->isoDataStart = 21;
    c->isoVolumeSectors = 21;
    c->isoReadSector = 0;
    kureha_strcopy(c->currentDir, sizeof(c->currentDir), "\\");
    kureha_strcopy(c->lastIsoName, sizeof(c->lastIsoName), "");
    for (i = 0; i < MAX_ISO_ITEMS; ++i) kureha_zero(&c->isoItems[i], sizeof(c->isoItems[i]));
}

static void clear_text(ZENKI_CONTEXT* c) {
    u32 t, l, i;
    if (!c) return;
    c->trackTextEnabled = 0;
    for (l = 0; l < MAX_LANG; ++l) {
        for (i = 0; i < MAX_INFO; ++i) c->textEnabled[l][i] = 0;
    }
    for (t = 0; t < 100; ++t) {
        for (l = 0; l < MAX_LANG; ++l) {
            for (i = 0; i < MAX_INFO; ++i) {
                c->textUsed[t][l][i] = 0;
                c->text[t][l][i][0] = 0;
            }
        }
    }
}

static void reset_context(ZENKI_CONTEXT* c) {
    close_read(c);
    clear_iso(c);
    clear_tracks(c);
    clear_text(c);
}

static u32 file_size_low(const char* path) {
    HANDLE32 h;
    u32 hi = 0, lo;
    if (!path || !path[0]) return 0;
    h = kureha_CreateFileA(path, K_GENERIC_READ, K_FILE_SHARE_READ | K_FILE_SHARE_WRITE, K_OPEN_EXISTING);
    if (h == K_INVALID_HANDLE) return 0;
    lo = kureha_GetFileSize(h, &hi);
    kureha_CloseHandle(h);
    if (hi != 0 || lo == 0xffffffffu) return 0;
    return lo;
}

static u32 infer_sector_size(u32 flag, u32 fileBytes) {
    if (flag == 2048 || flag == 8) return 2048;
    if (flag == 2352 || flag == 1) return 2352;
    if (flag == 2368 || flag == 3) return 2368;
    if (flag == 2448 || flag == 4) return 2448;
    if (fileBytes && (fileBytes % 2352u) == 0) return 2352;
    return 2048;
}

static int read_exact(HANDLE32 h, ptr buf, u32 bytes) {
    u32 got = 0;
    if (!kureha_ReadFile(h, buf, bytes, &got)) return 0;
    return got == bytes;
}

static int read_at(HANDLE32 h, u32 pos, ptr buf, u32 bytes) {
    i32 hi = 0;
    kureha_SetFilePointer(h, (i32)pos, &hi, K_FILE_BEGIN);
    return read_exact(h, buf, bytes);
}

static void analyze_track_file(ZENKI_TRACK* t) {
    HANDLE32 h;
    u8 hdr[12];
    u8 chunk[8];
    u32 pos, fileBytes;
    if (!t || !t->path[0]) return;
    fileBytes = file_size_low(t->path);
    t->fileBytes = fileBytes;
    t->dataOffset = 0;
    t->dataBytes = fileBytes;
    t->isWave = 0;
    t->sectorSize = infer_sector_size(t->flag, fileBytes);
    t->sectorCount = t->sectorSize ? (fileBytes / t->sectorSize) : 0;

    h = kureha_CreateFileA(t->path, K_GENERIC_READ, K_FILE_SHARE_READ | K_FILE_SHARE_WRITE, K_OPEN_EXISTING);
    if (h == K_INVALID_HANDLE) return;
    if (read_exact(h, hdr, 12)) {
        if (hdr[0]=='R' && hdr[1]=='I' && hdr[2]=='F' && hdr[3]=='F' && hdr[8]=='W' && hdr[9]=='A' && hdr[10]=='V' && hdr[11]=='E') {
            pos = 12;
            while (pos + 8 < fileBytes && pos < 0x100000u) {
                if (!read_at(h, pos, chunk, 8)) break;
                {
                    u32 sz = (u32)chunk[4] | ((u32)chunk[5]<<8) | ((u32)chunk[6]<<16) | ((u32)chunk[7]<<24);
                    if (chunk[0]=='d' && chunk[1]=='a' && chunk[2]=='t' && chunk[3]=='a') {
                        t->isWave = 1;
                        t->dataOffset = pos + 8;
                        t->dataBytes = sz;
                        t->sectorSize = 2352;
                        t->sectorCount = sz / 2352u;
                        break;
                    }
                    pos += 8 + sz + (sz & 1u);
                }
            }
        }
    }
    kureha_CloseHandle(h);
}

static char iso_upper_char(char ch) {
    if (ch >= 'a' && ch <= 'z') return (char)(ch - 32);
    if (ch >= 'A' && ch <= 'Z') return ch;
    if (ch >= '0' && ch <= '9') return ch;
    if (ch == '_') return ch;
    if (ch == '.') return ch;
    return '_';
}

static const char* basename_ptr(const char* path) {
    const char* b = path;
    u32 i;
    if (!path) return "";
    for (i = 0; path[i]; ++i) {
        if (path[i] == '\\' || path[i] == '/' || path[i] == ':') b = path + i + 1;
    }
    return b;
}

static int iso_is_slash(char ch) {
    return ch == '\\' || ch == '/';
}

static void normalize_iso_path(char* dst, u32 cap, const char* currentDir, const char* name, u32 isDir) {
    u32 k = 0, i = 0;
    int lastSlash = 0;
    if (!dst || cap == 0) return;
    dst[0] = 0;
    if (!name) name = "";

    if (!iso_is_slash(name[0]) && currentDir && currentDir[0] && !kureha_streq(currentDir, "\\")) {
        while (currentDir[i] && k + 1 < cap) {
            char ch = iso_is_slash(currentDir[i]) ? '\\' : currentDir[i];
            if (ch == '\\') {
                if (!lastSlash && k > 0) dst[k++] = ch;
                lastSlash = 1;
            } else {
                dst[k++] = ch;
                lastSlash = 0;
            }
            ++i;
        }
        if (k > 0 && !lastSlash && k + 1 < cap) {
            dst[k++] = '\\';
            lastSlash = 1;
        }
    }

    i = 0;
    while (name[i] && k + 1 < cap) {
        char ch = iso_is_slash(name[i]) ? '\\' : name[i];
        if (ch == '\\') {
            if (!lastSlash && k > 0) dst[k++] = ch;
            lastSlash = 1;
        } else if (ch != ':') {
            dst[k++] = ch;
            lastSlash = 0;
        }
        ++i;
    }

    while (k > 0 && dst[k - 1] == '\\') --k;
    dst[k] = 0;
    if (dst[0] == 0 && isDir) kureha_strcopy(dst, cap, "\\");
}

static void parent_path_of(char* dst, u32 cap, const char* path) {
    u32 i, last = 0xffffffffu;
    if (!dst || cap == 0) return;
    dst[0] = 0;
    if (!path) return;
    for (i = 0; path[i]; ++i) {
        if (iso_is_slash(path[i])) last = i;
    }
    if (last == 0xffffffffu) return;
    for (i = 0; i < last && i + 1 < cap; ++i) dst[i] = path[i];
    dst[i] = 0;
}

static int iso_path_equal(const char* a, const char* b) {
    u32 i = 0;
    if (!a || !b) return 0;
    while (a[i] && b[i]) {
        char ca = iso_is_slash(a[i]) ? '\\' : a[i];
        char cb = iso_is_slash(b[i]) ? '\\' : b[i];
        if (kureha_lower_char(ca) != kureha_lower_char(cb)) return 0;
        ++i;
    }
    return a[i] == 0 && b[i] == 0;
}

static int iso_is_descendant_path(const char* child, const char* parent) {
    u32 i = 0;
    if (!child || !parent || !parent[0]) return 0;
    while (parent[i]) {
        char cc = iso_is_slash(child[i]) ? '\\' : child[i];
        char pc = iso_is_slash(parent[i]) ? '\\' : parent[i];
        if (kureha_lower_char(cc) != kureha_lower_char(pc)) return 0;
        ++i;
    }
    return iso_is_slash(child[i]);
}

static u32 find_iso_item_index(ZENKI_CONTEXT* c, const char* path, u32 wantDir, u32 useType) {
    u32 i;
    if (!c || !path) return ISO_ROOT_PARENT;
    for (i = 0; i < c->isoCount; ++i) {
        ZENKI_ISO_ITEM* it = &c->isoItems[i];
        if (!it->used) continue;
        if (useType && it->isDir != wantDir) continue;
        if (iso_path_equal(it->name, path) || iso_path_equal(it->isoName, path)) return i;
    }
    return ISO_ROOT_PARENT;
}

static u32 find_parent_index(ZENKI_CONTEXT* c, const char* path) {
    char parent[260];
    parent_path_of(parent, sizeof(parent), path);
    if (!parent[0]) return ISO_ROOT_PARENT;
    return find_iso_item_index(c, parent, 1, 1);
}

static void make_iso_name(char* dst, u32 cap, const char* source, u32 isDir) {
    const char* b = basename_ptr(source);
    char base[32];
    char ext[16];
    u32 i = 0, j = 0, dot = 0xffffffffu;
    if (!dst || cap == 0) return;
    base[0] = 0; ext[0] = 0;
    if (!b || !b[0]) b = "FILE";
    for (i = 0; b[i]; ++i) if (b[i] == '.') dot = i;
    if (dot == 0xffffffffu) dot = i;
    for (i = 0; b[i] && i < dot && j < 8; ++i) base[j++] = iso_upper_char(b[i]);
    if (j == 0) base[j++] = 'F';
    base[j] = 0;
    j = 0;
    if (!isDir && b[dot] == '.') {
        for (i = dot + 1; b[i] && j < 3; ++i) ext[j++] = iso_upper_char(b[i]);
    }
    ext[j] = 0;
    if (isDir || ext[0] == 0) {
        kureha_strcopy(dst, cap, base);
    } else {
        u32 k = 0, m = 0;
        while (base[m] && k + 1 < cap) dst[k++] = base[m++];
        if (k + 1 < cap) dst[k++] = '.';
        m = 0; while (ext[m] && k + 1 < cap) dst[k++] = ext[m++];
        if (k + 2 < cap) { dst[k++] = ';'; dst[k++] = '1'; }
        dst[k] = 0;
    }
}

static u32 dir_record_length(u32 nameLen) {
    u32 n = 33u + nameLen;
    if (n & 1u) ++n;
    return n;
}

static void iso_record_date(u8* p) {
    p[0] = 125; p[1] = 1; p[2] = 1; p[3] = 0; p[4] = 0; p[5] = 0; p[6] = 0;
}

static u32 write_dir_record_raw(u8* p, u32 extent, u32 size, u32 flags, const u8* name, u32 nameLen) {
    u32 len = dir_record_length(nameLen);
    if (!p || !name || nameLen == 0 || len > 255) return 0;
    kureha_zero(p, len);
    p[0] = (u8)len;
    p[1] = 0;
    put_both32(p + 2, extent);
    put_both32(p + 10, size);
    iso_record_date(p + 18);
    p[25] = (u8)flags;
    p[26] = 0;
    p[27] = 0;
    put_both16(p + 28, 1);
    p[32] = (u8)nameLen;
    kureha_copy(p + 33, name, nameLen);
    return len;
}

static u32 write_dir_record_name(u8* p, u32 extent, u32 size, u32 flags, const char* name) {
    return write_dir_record_raw(p, extent, size, flags, (const u8*)name, kureha_strlen(name));
}

static ZENKI_ISO_ITEM* add_iso_item(ZENKI_CONTEXT* c, const char* name, u32 isDir) {
    ZENKI_ISO_ITEM* it;
    char fullPath[260];
    if (!c || !name || !name[0] || c->isoCount >= MAX_ISO_ITEMS) return (ZENKI_ISO_ITEM*)0;
    normalize_iso_path(fullPath, sizeof(fullPath), c->currentDir, name, isDir);
    if (!fullPath[0] || kureha_streq(fullPath, "\\")) return (ZENKI_ISO_ITEM*)0;
    if (find_iso_item_index(c, fullPath, isDir, 0) != ISO_ROOT_PARENT) return (ZENKI_ISO_ITEM*)0;
    it = &c->isoItems[c->isoCount++];
    kureha_zero(it, sizeof(*it));
    it->used = 1;
    it->isDir = isDir;
    it->parent = find_parent_index(c, fullPath);
    kureha_strcopy(it->name, sizeof(it->name), fullPath);
    make_iso_name(it->isoName, sizeof(it->isoName), fullPath, isDir);
    kureha_strcopy(c->lastIsoName, sizeof(c->lastIsoName), fullPath);
    c->isoLayoutValid = 0;
    return it;
}

static u32 compute_dir_bytes(ZENKI_CONTEXT* c, u32 parentIndex) {
    u32 i, total;
    if (!c) return 2048;
    total = dir_record_length(1) + dir_record_length(1);
    for (i = 0; i < c->isoCount; ++i) {
        ZENKI_ISO_ITEM* it = &c->isoItems[i];
        if (!it->used) continue;
        if (it->parent != parentIndex) continue;
        if (!it->isoName[0]) make_iso_name(it->isoName, sizeof(it->isoName), it->name, it->isDir);
        total += dir_record_length(kureha_strlen(it->isoName));
    }
    return align2048(total);
}

static u32 dir_table_number(ZENKI_CONTEXT* c, u32 itemIndex) {
    u32 i, n = 2;
    if (itemIndex == ISO_ROOT_PARENT) return 1;
    if (!c || itemIndex >= c->isoCount || !c->isoItems[itemIndex].isDir) return 1;
    for (i = 0; i < c->isoCount; ++i) {
        if (!c->isoItems[i].used || !c->isoItems[i].isDir) continue;
        if (i == itemIndex) return n;
        ++n;
    }
    return 1;
}

static u32 path_table_bytes(ZENKI_CONTEXT* c) {
    u32 i, total = 10; /* root path table entry */
    if (!c) return total;
    for (i = 0; i < c->isoCount; ++i) {
        ZENKI_ISO_ITEM* it = &c->isoItems[i];
        if (!it->used || !it->isDir) continue;
        total += 8 + kureha_strlen(it->isoName) + (kureha_strlen(it->isoName) & 1u);
    }
    return total;
}

static void layout_iso(ZENKI_CONTEXT* c) {
    u32 i, cur;
    if (!c) return;
    c->isoPvdSector = 16;
    c->isoVdstSector = 17;
    c->isoLPathSector = 18;
    c->isoMPathSector = 19;
    c->isoRootSector = 20;
    for (i = 0; i < c->isoCount; ++i) {
        ZENKI_ISO_ITEM* it = &c->isoItems[i];
        if (!it->used) continue;
        it->parent = find_parent_index(c, it->name);
        if (!it->isoName[0]) make_iso_name(it->isoName, sizeof(it->isoName), it->name, it->isDir);
    }
    c->isoPathTableBytes = path_table_bytes(c);
    c->isoPathTableSectors = divceil2048(c->isoPathTableBytes);
    if (c->isoPathTableSectors == 0) c->isoPathTableSectors = 1;
    c->isoLPathSector = 18;
    c->isoMPathSector = c->isoLPathSector + c->isoPathTableSectors;
    c->isoRootSector = c->isoMPathSector + c->isoPathTableSectors;
    c->isoRootBytes = compute_dir_bytes(c, ISO_ROOT_PARENT);
    c->isoRootSectors = divceil2048(c->isoRootBytes);
    if (c->isoRootSectors == 0) c->isoRootSectors = 1;
    c->isoDataStart = c->isoRootSector + c->isoRootSectors;
    cur = c->isoDataStart;
    for (i = 0; i < c->isoCount; ++i) {
        ZENKI_ISO_ITEM* it = &c->isoItems[i];
        if (!it->used) continue;
        if (it->isDir) {
            it->dirBytes = compute_dir_bytes(c, i);
            it->dirSectors = divceil2048(it->dirBytes);
            if (it->dirSectors == 0) it->dirSectors = 1;
            it->sectorCount = it->dirSectors;
            it->extent = cur;
            cur += it->sectorCount;
        }
    }
    for (i = 0; i < c->isoCount; ++i) {
        ZENKI_ISO_ITEM* it = &c->isoItems[i];
        if (!it->used || it->isDir) continue;
        if (it->original[0] && it->sizeBytes == 0) it->sizeBytes = file_size_low(it->original);
        it->sectorCount = divceil2048(it->sizeBytes);
        it->extent = cur;
        cur += it->sectorCount;
    }
    c->isoVolumeSectors = cur;
    if (c->isoVolumeSectors < c->isoDataStart) c->isoVolumeSectors = c->isoDataStart;
    c->isoLayoutValid = 1;
}

static void put_ascii_padded(u8* p, u32 len, const char* s) {
    u32 i;
    for (i = 0; i < len; ++i) p[i] = ' ';
    if (!s) return;
    for (i = 0; i < len && s[i]; ++i) p[i] = (u8)s[i];
}

static void put_iso_datetime(u8* p) {
    /* YYYYMMDDHHMMSSccO, 17 bytes.  Static to avoid CRT/time imports. */
    const char* d = "202501010000000";
    u32 i;
    for (i = 0; i < 16; ++i) p[i] = (u8)d[i];
    p[16] = 0;
}

static void write_pvd(ZENKI_CONTEXT* c, u8* out) {
    u8 rootName = 0;
    kureha_zero(out, 2048);
    out[0] = 1;
    out[1] = 'C'; out[2] = 'D'; out[3] = '0'; out[4] = '0'; out[5] = '1';
    out[6] = 1;
    put_ascii_padded(out + 8, 32, "KUREHA_REBUILD");
    put_ascii_padded(out + 40, 32, "KUREHA_DISC");
    put_both32(out + 80, c->isoVolumeSectors);
    put_both16(out + 120, 1);
    put_both16(out + 124, 1);
    put_both16(out + 128, 2048);
    put_both32(out + 132, c->isoPathTableBytes);
    put_le32(out + 140, c->isoLPathSector);
    put_le32(out + 144, 0);
    kureha_put_be32(out + 148, c->isoMPathSector);
    kureha_put_be32(out + 152, 0);
    write_dir_record_raw(out + 156, c->isoRootSector, c->isoRootBytes, 2, &rootName, 1);
    put_ascii_padded(out + 190, 128, "KUREHA VB6 REBUILD");
    put_ascii_padded(out + 318, 128, "KUREHA VB6 REBUILD");
    put_ascii_padded(out + 446, 128, "OPENAI CLEAN-ROOM REMAKE");
    put_iso_datetime(out + 813);
    put_iso_datetime(out + 830);
    put_iso_datetime(out + 847);
    put_iso_datetime(out + 864);
    out[881] = 1;
}

static void write_vdst(u8* out) {
    kureha_zero(out, 2048);
    out[0] = 255;
    out[1] = 'C'; out[2] = 'D'; out[3] = '0'; out[4] = '0'; out[5] = '1';
    out[6] = 1;
}

static void write_path_table_entry(u8* out, u32 sectorIndex, u32* streamOffset, u32 bigEndian, u32 extent, u32 parentNo, const u8* name, u32 nameLen) {
    u8 rec[128];
    u32 pos = 0;
    u32 recLen;
    u32 secStart = sectorIndex * 2048u;
    u32 secEnd = secStart + 2048u;
    if (nameLen > 96) nameLen = 96;
    rec[pos++] = (u8)nameLen;
    rec[pos++] = 0;
    if (bigEndian) {
        kureha_put_be32(rec + pos, extent); pos += 4;
        kureha_put_be16(rec + pos, parentNo); pos += 2;
    } else {
        put_le32(rec + pos, extent); pos += 4;
        put_le16(rec + pos, parentNo); pos += 2;
    }
    kureha_copy(rec + pos, name, nameLen);
    pos += nameLen;
    if (nameLen & 1u) rec[pos++] = 0;
    recLen = pos;
    if (*streamOffset + recLen > secStart && *streamOffset < secEnd) {
        u32 dstOff = *streamOffset > secStart ? *streamOffset - secStart : 0;
        u32 srcOff = secStart > *streamOffset ? secStart - *streamOffset : 0;
        u32 cp = min_u32(recLen - srcOff, 2048u - dstOff);
        kureha_copy(out + dstOff, rec + srcOff, cp);
    }
    *streamOffset += recLen;
}

static void write_path_table(ZENKI_CONTEXT* c, u8* out, u32 sectorIndex, u32 bigEndian) {
    u32 i, offset = 0;
    u8 rootName = 0;
    kureha_zero(out, 2048);
    write_path_table_entry(out, sectorIndex, &offset, bigEndian, c->isoRootSector, 1, &rootName, 1);
    for (i = 0; i < c->isoCount; ++i) {
        ZENKI_ISO_ITEM* it = &c->isoItems[i];
        if (!it->used || !it->isDir) continue;
        write_path_table_entry(out, sectorIndex, &offset, bigEndian, it->extent, dir_table_number(c, it->parent), (const u8*)it->isoName, kureha_strlen(it->isoName));
    }
}

static u32 dir_extent_for(ZENKI_CONTEXT* c, u32 dirIndex) {
    if (dirIndex == ISO_ROOT_PARENT) return c->isoRootSector;
    if (!c || dirIndex >= c->isoCount) return c ? c->isoRootSector : 0;
    return c->isoItems[dirIndex].extent;
}

static u32 dir_size_for(ZENKI_CONTEXT* c, u32 dirIndex) {
    if (dirIndex == ISO_ROOT_PARENT) return c->isoRootBytes;
    if (!c || dirIndex >= c->isoCount) return c ? c->isoRootBytes : 2048;
    return c->isoItems[dirIndex].dirBytes;
}

static void write_dir_sector(ZENKI_CONTEXT* c, u32 dirIndex, u32 dirSectorIndex, u8* out) {
    u32 i;
    u32 streamOffset = 0;
    u32 secStart = dirSectorIndex * 2048u;
    u32 secEnd = secStart + 2048u;
    u8 rec[256];
    u8 dot0 = 0, dot1 = 1;
    u32 parentIndex = ISO_ROOT_PARENT;
    if (dirIndex != ISO_ROOT_PARENT && dirIndex < c->isoCount) parentIndex = c->isoItems[dirIndex].parent;
    kureha_zero(out, 2048);
#define EMIT_REC(expr_len) do { \
        u32 rl = (expr_len); \
        if (rl) { \
            if (streamOffset + rl > secStart && streamOffset < secEnd) { \
                u32 dstOff = streamOffset > secStart ? streamOffset - secStart : 0; \
                u32 srcOff = secStart > streamOffset ? secStart - streamOffset : 0; \
                u32 cp = min_u32(rl - srcOff, 2048u - dstOff); \
                kureha_copy(out + dstOff, rec + srcOff, cp); \
            } \
            streamOffset += rl; \
        } \
    } while (0)
    kureha_zero(rec, sizeof(rec));
    EMIT_REC(write_dir_record_raw(rec, dir_extent_for(c, dirIndex), dir_size_for(c, dirIndex), 2, &dot0, 1));
    kureha_zero(rec, sizeof(rec));
    EMIT_REC(write_dir_record_raw(rec, dir_extent_for(c, parentIndex), dir_size_for(c, parentIndex), 2, &dot1, 1));
    for (i = 0; i < c->isoCount; ++i) {
        ZENKI_ISO_ITEM* it = &c->isoItems[i];
        if (!it->used) continue;
        if (it->parent != dirIndex) continue;
        kureha_zero(rec, sizeof(rec));
        EMIT_REC(write_dir_record_name(rec, it->extent, it->isDir ? it->dirBytes : it->sizeBytes, it->isDir ? 2u : 0u, it->isoName));
        if (streamOffset >= secEnd) break;
    }
#undef EMIT_REC
}

static int read_file_sector(const char* path, u32 pos, u8* out) {
    HANDLE32 h;
    i32 hi = 0;
    u32 got = 0;
    if (!path || !path[0]) return 0;
    h = kureha_CreateFileA(path, K_GENERIC_READ, K_FILE_SHARE_READ | K_FILE_SHARE_WRITE, K_OPEN_EXISTING);
    if (h == K_INVALID_HANDLE) return 0;
    kureha_SetFilePointer(h, (i32)pos, &hi, K_FILE_BEGIN);
    kureha_ReadFile(h, out, 2048, &got);
    kureha_CloseHandle(h);
    return got > 0;
}

static void write_iso_sector(ZENKI_CONTEXT* c, u32 lba, u8* out2048) {
    u32 i;
    kureha_zero(out2048, 2048);
    if (!c->isoLayoutValid) layout_iso(c);
    if (lba < 16) return;
    if (lba == c->isoPvdSector) { write_pvd(c, out2048); return; }
    if (lba == c->isoVdstSector) { write_vdst(out2048); return; }
    if (lba >= c->isoLPathSector && lba < c->isoLPathSector + c->isoPathTableSectors) {
        write_path_table(c, out2048, lba - c->isoLPathSector, 0);
        return;
    }
    if (lba >= c->isoMPathSector && lba < c->isoMPathSector + c->isoPathTableSectors) {
        write_path_table(c, out2048, lba - c->isoMPathSector, 1);
        return;
    }
    if (lba >= c->isoRootSector && lba < c->isoRootSector + c->isoRootSectors) {
        write_dir_sector(c, ISO_ROOT_PARENT, lba - c->isoRootSector, out2048);
        return;
    }
    for (i = 0; i < c->isoCount; ++i) {
        ZENKI_ISO_ITEM* it = &c->isoItems[i];
        if (!it->used || !it->isDir || it->sectorCount == 0) continue;
        if (lba >= it->extent && lba < it->extent + it->sectorCount) {
            write_dir_sector(c, i, lba - it->extent, out2048);
            return;
        }
    }
    for (i = 0; i < c->isoCount; ++i) {
        ZENKI_ISO_ITEM* it = &c->isoItems[i];
        if (!it->used || it->isDir || it->sectorCount == 0) continue;
        if (lba >= it->extent && lba < it->extent + it->sectorCount) {
            u32 fileOff = (lba - it->extent) * 2048u;
            if (it->original[0]) read_file_sector(it->original, fileOff, out2048);
            return;
        }
    }
}

static void raw2352_from_2048(u32 lba, const u8* user, u8* raw) {
    static const u8 sync[12] = {0x00,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0xff,0x00};
    u32 msf = lba + 150u;
    u32 mm = msf / (75u * 60u);
    u32 ss = (msf / 75u) % 60u;
    u32 ff = msf % 75u;
    kureha_zero(raw, 2352);
    kureha_copy(raw, sync, 12);
    raw[12] = (u8)(((mm / 10u) << 4) | (mm % 10u));
    raw[13] = (u8)(((ss / 10u) << 4) | (ss % 10u));
    raw[14] = (u8)(((ff / 10u) << 4) | (ff % 10u));
    raw[15] = 1;
    kureha_copy(raw + 16, user, 2048);
}

static u32 normalize_track_no(u32 trackNo) {
    if (trackNo >= 100) return 99;
    return trackNo;
}

static const char* track_display_text(ZENKI_CONTEXT* c, u32 zeroBasedTrack) {
    u32 tn = zeroBasedTrack + 1;
    if (!c || tn >= 100) return "";
    if (c->trackTextEnabled) {
        if (c->textUsed[tn][0][0] && c->text[tn][0][0][0]) return c->text[tn][0][0];
        if (c->textUsed[tn][1][0] && c->text[tn][1][0][0]) return c->text[tn][1][0];
    }
    return "";
}

static int open_current_track(ZENKI_CONTEXT* c) {
    ZENKI_TRACK* t;
    if (!c || c->readTrack >= c->trackCount) return 0;
    t = &c->tracks[c->readTrack];
    close_read(c);
    if (!t->used || !t->path[0]) return 0;
    c->readHandle = kureha_CreateFileA(t->path, K_GENERIC_READ, K_FILE_SHARE_READ | K_FILE_SHARE_WRITE, K_OPEN_EXISTING);
    if (c->readHandle == K_INVALID_HANDLE) return 0;
    c->reading = 1;
    c->readMode = 1;
    c->readPhase = 0;
    c->readSectorInPhase = 0;
    c->readByteInData = 0;
    if (t->dataOffset) {
        i32 hi = 0;
        kureha_SetFilePointer(c->readHandle, (i32)t->dataOffset, &hi, K_FILE_BEGIN);
    }
    return 1;
}

static u32 sector_bytes_for_read(ZENKI_TRACK* t, u32 rawMode) {
    if (!t) return rawMode ? 2352u : 2048u;
    if (t->isWave) return 2352u;
    if (rawMode) return t->sectorSize >= 2352u ? t->sectorSize : 2352u;
    return t->sectorSize == 2048u ? 2048u : t->sectorSize;
}

u32 STDCALL GetEngineVersion(void) { return 10407u; }

ptr STDCALL Initialize(void) {
    int i;
    for (i = 0; i < MAX_CTX; ++i) {
        if (!g_ctx[i].used) {
            ZENKI_CONTEXT* c = &g_ctx[i];
            kureha_zero(c, sizeof(*c));
            c->magic = ZENKI_MAGIC;
            c->used = 1;
            c->readHandle = K_INVALID_HANDLE;
            reset_context(c);
            return (ptr)c;
        }
    }
    return (ptr)0;
}

void STDCALL Terminate(ptr h) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    if (!c) return;
    reset_context(c);
    kureha_zero(c, sizeof(*c));
}

u32 STDCALL GetTOCStructure(ptr h, u32 tocType, void* outBuffer, u32 outBytes) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    u32 i;
    (void)tocType;
    if (!c) return 0;
    if (!c->isoLayoutValid) layout_iso(c);
    if (outBuffer && outBytes >= 8) {
        u32* p = (u32*)outBuffer;
        u32 maxDwords = outBytes / 4;
        kureha_zero(outBuffer, outBytes);
        p[0] = c->trackCount ? c->trackCount : (c->isoCount ? 1u : 0u);
        p[1] = c->isoVolumeSectors;
        for (i = 0; i < c->trackCount && (2 + i * 4 + 3) < maxDwords; ++i) {
            p[2 + i * 4 + 0] = i + 1;
            p[2 + i * 4 + 1] = c->tracks[i].pregap;
            p[2 + i * 4 + 2] = c->tracks[i].sectorCount;
            p[2 + i * 4 + 3] = c->tracks[i].flag;
        }
    }
    return c->trackCount ? c->trackCount : (c->isoCount ? 1u : 0u);
}

u32 STDCALL InitISOFS(ptr h, u32 flags) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    if (!c) return 0;
    clear_iso(c);
    c->isoFlags = flags;
    return 1;
}

void STDCALL ClearISO(ptr h) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    if (c) clear_iso(c);
}

u32 STDCALL IsISOEmpty(ptr h) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    return (c && c->isoCount == 0) ? 1u : 0u;
}

u32 STDCALL AddISOFile(ptr h, const char* isoName, const char* originalFileName) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    ZENKI_ISO_ITEM* it = add_iso_item(c, isoName, 0);
    if (!it) return 0;
    kureha_strcopy(it->original, sizeof(it->original), originalFileName ? originalFileName : "");
    it->sizeBytes = file_size_low(originalFileName);
    return 1;
}

u32 STDCALL AddISODummyFile(ptr h, const char* isoName, u32 byteSizeOrAttr, u32 attr, u32 stamp) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    ZENKI_ISO_ITEM* it = add_iso_item(c, isoName, 0);
    if (!it) return 0;
    it->sizeBytes = byteSizeOrAttr;
    it->attr = attr;
    it->stamp = stamp;
    return 1;
}

u32 STDCALL MakeISODirectory(ptr h, const char* isoName, u32 attr, u32 stamp) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    ZENKI_ISO_ITEM* it = add_iso_item(c, isoName, 1);
    if (!it) return 0;
    it->attr = attr;
    it->stamp = stamp;
    return 1;
}

void STDCALL GetISONewFileDirectoryName(ptr h, char* outName, u32 outBytes) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    kureha_strcopy(outName, outBytes, c ? c->lastIsoName : "");
}

u32 STDCALL ChangeISODirectory(ptr h, const char* isoName) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    char fullPath[260];
    u32 idx;
    if (!c || !isoName) return 0;
    if (kureha_streq(isoName, ".")) return 1;
    if (kureha_streq(isoName, "..")) {
        parent_path_of(fullPath, sizeof(fullPath), c->currentDir);
        if (!fullPath[0]) kureha_strcopy(c->currentDir, sizeof(c->currentDir), "\\");
        else kureha_strcopy(c->currentDir, sizeof(c->currentDir), fullPath);
        return 1;
    }
    normalize_iso_path(fullPath, sizeof(fullPath), c->currentDir, isoName, 1);
    if (!fullPath[0] || kureha_streq(fullPath, "\\")) {
        kureha_strcopy(c->currentDir, sizeof(c->currentDir), "\\");
        return 1;
    }
    idx = find_iso_item_index(c, fullPath, 1, 1);
    if (idx == ISO_ROOT_PARENT) return 0;
    kureha_strcopy(c->currentDir, sizeof(c->currentDir), fullPath);
    return 1;
}

u32 STDCALL RemoveISOFile(ptr h, const char* isoName) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    u32 i, j, removed = 0;
    char fullPath[260];
    if (!c || !isoName) return 0;
    normalize_iso_path(fullPath, sizeof(fullPath), c->currentDir, isoName, 0);
    for (i = 0; i < c->isoCount; ++i) {
        if (iso_path_equal(c->isoItems[i].name, fullPath) || iso_path_equal(c->isoItems[i].isoName, isoName) || iso_is_descendant_path(c->isoItems[i].name, fullPath)) {
            for (j = i; j + 1 < c->isoCount; ++j) c->isoItems[j] = c->isoItems[j + 1];
            --c->isoCount;
            --i;
            removed = 1;
        }
    }
    if (removed) c->isoLayoutValid = 0;
    return removed ? 1u : 0u;
}

u32 STDCALL RenameISOFile(ptr h, const char* oldName, const char* newName) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    u32 i, oldLen, found = 0;
    char oldPath[260];
    char newPath[260];
    if (!c || !oldName || !newName) return 0;
    normalize_iso_path(oldPath, sizeof(oldPath), c->currentDir, oldName, 0);
    normalize_iso_path(newPath, sizeof(newPath), c->currentDir, newName, 0);
    oldLen = kureha_strlen(oldPath);
    for (i = 0; i < c->isoCount; ++i) {
        if (iso_path_equal(c->isoItems[i].name, oldPath) || iso_path_equal(c->isoItems[i].isoName, oldName)) {
            kureha_strcopy(c->isoItems[i].name, sizeof(c->isoItems[i].name), newPath);
            make_iso_name(c->isoItems[i].isoName, sizeof(c->isoItems[i].isoName), newPath, c->isoItems[i].isDir);
            kureha_strcopy(c->lastIsoName, sizeof(c->lastIsoName), newPath);
            c->isoLayoutValid = 0;
            found = 1;
        } else if (iso_is_descendant_path(c->isoItems[i].name, oldPath)) {
            char tmp[260];
            u32 k = 0, j = 0;
            while (newPath[k] && k + 1 < sizeof(tmp)) { tmp[k] = newPath[k]; ++k; }
            if (k + 1 < sizeof(tmp)) tmp[k++] = '\\';
            j = oldLen + 1;
            while (c->isoItems[i].name[j] && k + 1 < sizeof(tmp)) tmp[k++] = c->isoItems[i].name[j++];
            tmp[k] = 0;
            kureha_strcopy(c->isoItems[i].name, sizeof(c->isoItems[i].name), tmp);
            make_iso_name(c->isoItems[i].isoName, sizeof(c->isoItems[i].isoName), tmp, c->isoItems[i].isDir);
            c->isoLayoutValid = 0;
            found = 1;
        }
    }
    return found ? 1u : 0u;
}

u32 STDCALL ChangeISOProperties(ptr h, const char* isoName, u32 attr, u32 stamp) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    u32 i;
    char fullPath[260];
    if (!c || !isoName) return 0;
    normalize_iso_path(fullPath, sizeof(fullPath), c->currentDir, isoName, 0);
    for (i = 0; i < c->isoCount; ++i) {
        if (iso_path_equal(c->isoItems[i].name, fullPath) || iso_path_equal(c->isoItems[i].isoName, isoName)) {
            c->isoItems[i].attr = attr;
            c->isoItems[i].stamp = stamp;
            return 1;
        }
    }
    return 0;
}

void STDCALL GetISOCurrentDirectory(ptr h, char* outDir, u32 outBytes) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    kureha_strcopy(outDir, outBytes, c ? c->currentDir : "");
}

static u32 write_find_data(ZENKI_ISO_ITEM* it, void* outData) {
    u8* p = (u8*)outData;
    if (!it || !outData) return 0;
    kureha_zero(outData, 1088);
    kureha_strcopy((char*)p, 512, it->name);
    kureha_strcopy((char*)p + 512, 256, it->isoName);
    ((u32*)(p + 1024))[0] = it->sizeBytes;
    ((u32*)(p + 1024))[1] = it->attr;
    ((u32*)(p + 1024))[2] = it->isDir;
    ((u32*)(p + 1024))[3] = it->stamp;
    ((u32*)(p + 1024))[4] = it->extent;
    ((u32*)(p + 1024))[5] = it->sectorCount;
    kureha_strcopy((char*)(p + 1056), 32, it->original);
    return 1;
}

u32 STDCALL FindISOFirstFile(ptr h, void* outData) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    if (!c || c->isoCount == 0) return 0;
    if (!c->isoLayoutValid) layout_iso(c);
    c->isoEnumIndex = 0;
    return write_find_data(&c->isoItems[0], outData);
}

u32 STDCALL FindISONextFile(ptr h, void* outData) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    if (!c) return 0;
    ++c->isoEnumIndex;
    if (c->isoEnumIndex >= c->isoCount) return 0;
    return write_find_data(&c->isoItems[c->isoEnumIndex], outData);
}

void STDCALL ClearTrack(ptr h) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    if (c) clear_tracks(c);
}

u32 STDCALL AddTrack(ptr h, const char* fileName, u32 pregap, u32 postgap, u32 flag) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    ZENKI_TRACK* t;
    if (!c || !fileName || c->trackCount >= MAX_TRACKS) return 0;
    t = &c->tracks[c->trackCount];
    kureha_zero(t, sizeof(*t));
    t->used = 1;
    kureha_strcopy(t->path, sizeof(t->path), fileName);
    t->pregap = pregap;
    t->postgap = postgap;
    t->flag = flag;
    analyze_track_file(t);
    ++c->trackCount;
    return 1;
}

u32 STDCALL ResetTrack(ptr h, u32 trackNo, const char* fileName, u32 pregap, u32 postgap, u32 flag) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    ZENKI_TRACK* t;
    u32 idx = trackNo;
    if (idx > 0) idx--;
    if (!c || !fileName || idx >= MAX_TRACKS) return 0;
    t = &c->tracks[idx];
    kureha_zero(t, sizeof(*t));
    t->used = 1;
    kureha_strcopy(t->path, sizeof(t->path), fileName);
    t->pregap = pregap;
    t->postgap = postgap;
    t->flag = flag;
    analyze_track_file(t);
    if (idx >= c->trackCount) c->trackCount = idx + 1;
    return 1;
}

u32 STDCALL RemoveTrack(ptr h, u32 trackNo) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    u32 idx = trackNo, j;
    if (idx > 0) idx--;
    if (!c || idx >= c->trackCount) return 0;
    for (j = idx; j + 1 < c->trackCount; ++j) c->tracks[j] = c->tracks[j + 1];
    --c->trackCount;
    return 1;
}

u32 STDCALL GetTrackCount(ptr h) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    return c ? c->trackCount : 0;
}

void STDCALL GetTrackInformation(ptr h, u32 trackNo, void* outInfo) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    ZENKI_TRACK* t;
    u32 idx = trackNo;
    u8* p = (u8*)outInfo;
    if (!outInfo) return;
    kureha_zero(outInfo, 0x410);
    if (idx > 0) idx--;
    if (!c || idx >= c->trackCount) return;
    t = &c->tracks[idx];
    kureha_strcopy((char*)p, 0x400, t->path);
    ((u32*)(p + 0x400))[0] = t->pregap;
    ((u32*)(p + 0x400))[1] = t->postgap;
    ((u32*)(p + 0x400))[2] = t->flag;
    ((u32*)(p + 0x400))[3] = t->sectorCount;
}

i32 STDCALL ReadStart(ptr h, const char* optionalPath) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    if (!c) return -1;
    close_read(c);
    if (c->trackCount == 0 && optionalPath && optionalPath[0]) AddTrack(h, optionalPath, 0, 0, 0);
    if (c->trackCount > 0) {
        c->readTrack = 0;
        if (!open_current_track(c)) return -1;
        return 0;
    }
    if (c->isoCount > 0) {
        layout_iso(c);
        c->reading = 1;
        c->readMode = 2;
        c->isoReadSector = 0;
        return 0;
    }
    return -1;
}

u32 STDCALL Read(ptr h, void* sectorBuffer, u32 rawMode) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    ZENKI_TRACK* t;
    u32 sectorBytes, got, want, dataLeft;
    if (!c || !sectorBuffer) return 0;
    if (c->readMode == 2) {
        u8 user[2048];
        if (c->isoReadSector >= c->isoVolumeSectors) { close_read(c); return 0; }
        write_iso_sector(c, c->isoReadSector, user);
        if (rawMode) {
            kureha_zero(sectorBuffer, 2448);
            raw2352_from_2048(c->isoReadSector, user, (u8*)sectorBuffer);
        } else {
            kureha_zero(sectorBuffer, 2448);
            kureha_copy(sectorBuffer, user, 2048);
        }
        ++c->isoReadSector;
        return 1;
    }
    while (c->readTrack < c->trackCount) {
        if (!c->reading || c->readMode != 1) {
            if (!open_current_track(c)) { ++c->readTrack; continue; }
        }
        t = &c->tracks[c->readTrack];
        sectorBytes = sector_bytes_for_read(t, rawMode);
        if (sectorBytes < 2048) sectorBytes = 2048;
        if (sectorBytes > 2448) sectorBytes = 2448;
        kureha_zero(sectorBuffer, 2448);
        if (c->readPhase == 0) {
            if (c->readSectorInPhase < t->pregap) { ++c->readSectorInPhase; return 1; }
            c->readPhase = 1; c->readSectorInPhase = 0;
        }
        if (c->readPhase == 1) {
            if (c->readByteInData >= t->dataBytes) {
                c->readPhase = 2; c->readSectorInPhase = 0;
            } else {
                u8 user[2048];
                u32 trackSector = 0;
                dataLeft = t->dataBytes - c->readByteInData;
                want = sectorBytes;
                if (!t->isWave && rawMode && t->sectorSize == 2048) want = 2048;
                if (want > dataLeft) want = dataLeft;
                got = 0;
                if (!t->isWave && rawMode && t->sectorSize == 2048) {
                    kureha_zero(user, sizeof(user));
                    trackSector = t->sectorSize ? (c->readByteInData / t->sectorSize) : 0;
                    if (!kureha_ReadFile(c->readHandle, user, want, &got)) got = 0;
                    raw2352_from_2048(t->pregap + trackSector, user, (u8*)sectorBuffer);
                } else {
                    if (!kureha_ReadFile(c->readHandle, sectorBuffer, want, &got)) got = 0;
                }
                c->readByteInData += got;
                if (got) ++c->readSectorInPhase;
                if (got == 0 && want != 0) { c->readPhase = 2; c->readSectorInPhase = 0; }
                else return 1;
            }
        }
        if (c->readPhase == 2) {
            if (c->readSectorInPhase < t->postgap) { ++c->readSectorInPhase; return 1; }
            close_read(c);
            ++c->readTrack;
            c->readPhase = 0; c->readSectorInPhase = 0; c->readByteInData = 0;
            continue;
        }
    }
    close_read(c);
    return 0;
}

void STDCALL ReadEnd(ptr h) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    if (c) close_read(c);
}

u32 STDCALL EnabledTrackText(ptr h, u32 enable) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    if (!c) return 0;
    c->trackTextEnabled = enable ? 1u : 0u;
    return 1;
}

void STDCALL SetTrackText(ptr h, u32 trackNo, u32 language, u32 information, const char* text) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    u32 tn = normalize_track_no(trackNo);
    if (!c || language >= MAX_LANG || information >= MAX_INFO) return;
    kureha_strcopy(c->text[tn][language][information], sizeof(c->text[tn][language][information]), text ? text : "");
    c->textUsed[tn][language][information] = (text && text[0]) ? 1u : 0u;
    if (text && text[0]) {
        c->trackTextEnabled = 1;
        c->textEnabled[language][information] = 1;
    }
}

void STDCALL GetTrackText(ptr h, u32 trackNo, u32 language, u32 information, char* outText, u32 outBytes) {
    ZENKI_CONTEXT* c = ctx_from_handle(h);
    u32 tn = normalize_track_no(trackNo);
    if (!outText || outBytes == 0) return;
    if (!c || language >= MAX_LANG || information >= MAX_INFO) { outText[0] = 0; return; }
    if (!c->textUsed[tn][language][information]) { outText[0] = 0; return; }
    kureha_strcopy(outText, outBytes, c->text[tn][language][information]);
}
