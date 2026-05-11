#ifndef KUREHA_COMMON_WINAPI_H
#define KUREHA_COMMON_WINAPI_H

/*
   Import-free Win32 helper layer for 32-bit PE DLLs.
   The DLL is linked with /noentry and /nostdlib, so we cannot rely on CRT or
   import libraries.  Needed kernel32/kernelbase functions are resolved from
   the PEB export lists at runtime.
*/

typedef unsigned char  u8;
typedef unsigned short u16;
typedef unsigned int   u32;
typedef signed int     i32;
typedef void*          ptr;

typedef int BOOL32;
typedef void* HANDLE32;

typedef struct LIST_ENTRY32_TAG {
    struct LIST_ENTRY32_TAG* Flink;
    struct LIST_ENTRY32_TAG* Blink;
} LIST_ENTRY32;

typedef struct UNICODE_STRING32_TAG {
    u16 Length;
    u16 MaximumLength;
    u16* Buffer;
} UNICODE_STRING32;

typedef struct PEB_LDR_DATA32_TAG {
    u32 Length;
    u8 Initialized;
    u8 pad1[3];
    ptr SsHandle;
    LIST_ENTRY32 InLoadOrderModuleList;
    LIST_ENTRY32 InMemoryOrderModuleList;
    LIST_ENTRY32 InInitializationOrderModuleList;
} PEB_LDR_DATA32;

typedef struct LDR_DATA_TABLE_ENTRY32_TAG {
    LIST_ENTRY32 InLoadOrderLinks;
    LIST_ENTRY32 InMemoryOrderLinks;
    LIST_ENTRY32 InInitializationOrderLinks;
    ptr DllBase;
    ptr EntryPoint;
    u32 SizeOfImage;
    UNICODE_STRING32 FullDllName;
    UNICODE_STRING32 BaseDllName;
} LDR_DATA_TABLE_ENTRY32;

typedef struct PEB32_TAG {
    u8 Reserved1[0x0c];
    PEB_LDR_DATA32* Ldr;
} PEB32;

static ptr kureha_get_peb(void) {
    ptr p;
    __asm__("movl %%fs:0x30, %0" : "=r"(p));
    return p;
}

static void kureha_zero(ptr p, u32 n) {
    u8* d = (u8*)p;
    while (n--) *d++ = 0;
}

static void kureha_copy(ptr dst, const void* src, u32 n) {
    u8* d = (u8*)dst;
    const u8* s = (const u8*)src;
    while (n--) *d++ = *s++;
}

static u32 kureha_strlen(const char* s) {
    u32 n = 0;
    if (!s) return 0;
    while (s[n]) ++n;
    return n;
}

static int kureha_streq(const char* a, const char* b) {
    u32 i = 0;
    if (!a || !b) return 0;
    while (a[i] && b[i]) {
        if (a[i] != b[i]) return 0;
        ++i;
    }
    return a[i] == b[i];
}

static char kureha_lower_char(char c) {
    if (c >= 'A' && c <= 'Z') return (char)(c + 32);
    return c;
}

static int kureha_starts_with_i(const char* s, const char* pfx) {
    u32 i = 0;
    if (!s || !pfx) return 0;
    while (pfx[i]) {
        if (kureha_lower_char(s[i]) != kureha_lower_char(pfx[i])) return 0;
        ++i;
    }
    return 1;
}

static int kureha_ends_with_i(const char* s, const char* suffix) {
    u32 sl = kureha_strlen(s), tl = kureha_strlen(suffix), i;
    if (tl > sl) return 0;
    for (i = 0; i < tl; ++i) {
        if (kureha_lower_char(s[sl - tl + i]) != kureha_lower_char(suffix[i])) return 0;
    }
    return 1;
}

static void kureha_strcopy(char* dst, u32 cap, const char* src) {
    u32 i;
    if (!dst || cap == 0) return;
    if (!src) src = "";
    for (i = 0; i + 1 < cap && src[i]; ++i) dst[i] = src[i];
    dst[i] = 0;
}

static int kureha_modname_eq(UNICODE_STRING32* us, const char* ascii) {
    u32 i = 0, n;
    if (!us || !us->Buffer || !ascii) return 0;
    n = (u32)(us->Length / 2);
    while (i < n && ascii[i]) {
        char wc = (char)(us->Buffer[i] & 0xff);
        if (kureha_lower_char(wc) != kureha_lower_char(ascii[i])) return 0;
        ++i;
    }
    return (i == n && ascii[i] == 0);
}

static u16 kureha_be16(const u8* p) { return (u16)(((u16)p[0] << 8) | p[1]); }
static u32 kureha_be32(const u8* p) { return ((u32)p[0] << 24) | ((u32)p[1] << 16) | ((u32)p[2] << 8) | p[3]; }
static void kureha_put_be16(u8* p, u32 v) { p[0] = (u8)(v >> 8); p[1] = (u8)v; }
static void kureha_put_be24(u8* p, u32 v) { p[0] = (u8)(v >> 16); p[1] = (u8)(v >> 8); p[2] = (u8)v; }
static void kureha_put_be32(u8* p, u32 v) { p[0] = (u8)(v >> 24); p[1] = (u8)(v >> 16); p[2] = (u8)(v >> 8); p[3] = (u8)v; }

/* PE export resolver */
static ptr kureha_find_export_in_module(ptr moduleBase, const char* name, int depth);

static ptr kureha_find_module_base(const char* moduleName) {
    PEB32* peb = (PEB32*)kureha_get_peb();
    LIST_ENTRY32* head;
    LIST_ENTRY32* cur;
    if (!peb || !peb->Ldr) return (ptr)0;
    head = &peb->Ldr->InMemoryOrderModuleList;
    cur = head->Flink;
    while (cur && cur != head) {
        LDR_DATA_TABLE_ENTRY32* e = (LDR_DATA_TABLE_ENTRY32*)((u8*)cur - 8);
        if (kureha_modname_eq(&e->BaseDllName, moduleName)) return e->DllBase;
        cur = cur->Flink;
    }
    return (ptr)0;
}

static ptr kureha_resolve_any_loaded(const char* name, int depth) {
    PEB32* peb = (PEB32*)kureha_get_peb();
    LIST_ENTRY32* head;
    LIST_ENTRY32* cur;
    if (!peb || !peb->Ldr || depth > 4) return (ptr)0;
    head = &peb->Ldr->InMemoryOrderModuleList;
    cur = head->Flink;
    while (cur && cur != head) {
        LDR_DATA_TABLE_ENTRY32* e = (LDR_DATA_TABLE_ENTRY32*)((u8*)cur - 8);
        ptr fp = kureha_find_export_in_module(e->DllBase, name, depth + 1);
        if (fp) return fp;
        cur = cur->Flink;
    }
    return (ptr)0;
}

static ptr kureha_find_export_by_forwarder(const char* fwd, int depth) {
    char mod[40];
    char fn[96];
    u32 i = 0, j = 0;
    ptr base;
    if (!fwd || depth > 4) return (ptr)0;
    while (fwd[i] && fwd[i] != '.' && i + 5 < sizeof(mod)) {
        mod[i] = fwd[i];
        ++i;
    }
    if (fwd[i] != '.') return (ptr)0;
    mod[i++] = '.'; mod[i++] = 'd'; mod[i++] = 'l'; mod[i++] = 'l'; mod[i] = 0;
    while (fwd[i] && j + 1 < sizeof(fn)) fn[j++] = fwd[i++];
    fn[j] = 0;
    base = kureha_find_module_base(mod);
    if (base) return kureha_find_export_in_module(base, fn, depth + 1);
    return kureha_resolve_any_loaded(fn, depth + 1);
}

static ptr kureha_find_export_in_module(ptr moduleBase, const char* name, int depth) {
    u8* base = (u8*)moduleBase;
    u32 e_lfanew, exportRva, exportSize, numNames, funcsRva, namesRva, ordsRva, i;
    u8* opt;
    u8* exp;
    if (!base || !name || depth > 4) return (ptr)0;
    if (base[0] != 'M' || base[1] != 'Z') return (ptr)0;
    e_lfanew = *(u32*)(base + 0x3c);
    if (*(u32*)(base + e_lfanew) != 0x00004550u) return (ptr)0;
    opt = base + e_lfanew + 0x18;
    exportRva = *(u32*)(opt + 0x60);
    exportSize = *(u32*)(opt + 0x64);
    if (!exportRva) return (ptr)0;
    exp = base + exportRva;
    numNames = *(u32*)(exp + 0x18);
    funcsRva = *(u32*)(exp + 0x1c);
    namesRva = *(u32*)(exp + 0x20);
    ordsRva  = *(u32*)(exp + 0x24);
    for (i = 0; i < numNames; ++i) {
        const char* nm = (const char*)(base + ((u32*)(base + namesRva))[i]);
        if (kureha_streq(nm, name)) {
            u16 ord = ((u16*)(base + ordsRva))[i];
            u32 rva = ((u32*)(base + funcsRva))[ord];
            if (rva >= exportRva && rva < exportRva + exportSize) {
                return kureha_find_export_by_forwarder((const char*)(base + rva), depth + 1);
            }
            return (ptr)(base + rva);
        }
    }
    return (ptr)0;
}

static ptr kureha_api(const char* name) {
    static int initialized = 0;
    static ptr pCreateFileA, pCloseHandle, pDeviceIoControl, pReadFile, pWriteFile;
    static ptr pSetFilePointer, pGetFileSize, pFlushFileBuffers, pSetEndOfFile;
    if (!initialized) {
        pCreateFileA       = kureha_resolve_any_loaded("CreateFileA", 0);
        pCloseHandle       = kureha_resolve_any_loaded("CloseHandle", 0);
        pDeviceIoControl   = kureha_resolve_any_loaded("DeviceIoControl", 0);
        pReadFile          = kureha_resolve_any_loaded("ReadFile", 0);
        pWriteFile         = kureha_resolve_any_loaded("WriteFile", 0);
        pSetFilePointer    = kureha_resolve_any_loaded("SetFilePointer", 0);
        pGetFileSize       = kureha_resolve_any_loaded("GetFileSize", 0);
        pFlushFileBuffers  = kureha_resolve_any_loaded("FlushFileBuffers", 0);
        pSetEndOfFile      = kureha_resolve_any_loaded("SetEndOfFile", 0);
        initialized = 1;
    }
    if (kureha_streq(name, "CreateFileA")) return pCreateFileA;
    if (kureha_streq(name, "CloseHandle")) return pCloseHandle;
    if (kureha_streq(name, "DeviceIoControl")) return pDeviceIoControl;
    if (kureha_streq(name, "ReadFile")) return pReadFile;
    if (kureha_streq(name, "WriteFile")) return pWriteFile;
    if (kureha_streq(name, "SetFilePointer")) return pSetFilePointer;
    if (kureha_streq(name, "GetFileSize")) return pGetFileSize;
    if (kureha_streq(name, "FlushFileBuffers")) return pFlushFileBuffers;
    if (kureha_streq(name, "SetEndOfFile")) return pSetEndOfFile;
    return (ptr)0;
}

/* Win32 declarations as function-pointer typedefs. */
typedef HANDLE32 (__stdcall *PFN_CreateFileA)(const char*, u32, u32, ptr, u32, u32, HANDLE32);
typedef BOOL32   (__stdcall *PFN_CloseHandle)(HANDLE32);
typedef BOOL32   (__stdcall *PFN_DeviceIoControl)(HANDLE32, u32, ptr, u32, ptr, u32, u32*, ptr);
typedef BOOL32   (__stdcall *PFN_ReadFile)(HANDLE32, ptr, u32, u32*, ptr);
typedef BOOL32   (__stdcall *PFN_WriteFile)(HANDLE32, const void*, u32, u32*, ptr);
typedef u32      (__stdcall *PFN_SetFilePointer)(HANDLE32, i32, i32*, u32);
typedef u32      (__stdcall *PFN_GetFileSize)(HANDLE32, u32*);
typedef BOOL32   (__stdcall *PFN_FlushFileBuffers)(HANDLE32);
typedef BOOL32   (__stdcall *PFN_SetEndOfFile)(HANDLE32);

#define K_INVALID_HANDLE ((HANDLE32)(-1))
#define K_GENERIC_READ   0x80000000u
#define K_GENERIC_WRITE  0x40000000u
#define K_FILE_SHARE_READ  0x00000001u
#define K_FILE_SHARE_WRITE 0x00000002u
#define K_OPEN_EXISTING  3u
#define K_CREATE_ALWAYS  2u
#define K_FILE_BEGIN     0u
#define K_FILE_CURRENT   1u
#define K_FILE_END       2u

static HANDLE32 kureha_CreateFileA(const char* p, u32 access, u32 share, u32 createMode) {
    PFN_CreateFileA f = (PFN_CreateFileA)kureha_api("CreateFileA");
    if (!f) return K_INVALID_HANDLE;
    return f(p, access, share, (ptr)0, createMode, 0, (HANDLE32)0);
}
static BOOL32 kureha_CloseHandle(HANDLE32 h) {
    PFN_CloseHandle f = (PFN_CloseHandle)kureha_api("CloseHandle");
    if (!f || !h || h == K_INVALID_HANDLE) return 0;
    return f(h);
}
static BOOL32 kureha_DeviceIoControl(HANDLE32 h, u32 code, ptr inb, u32 inBytes, ptr outb, u32 outBytes, u32* ret) {
    PFN_DeviceIoControl f = (PFN_DeviceIoControl)kureha_api("DeviceIoControl");
    if (!f) return 0;
    return f(h, code, inb, inBytes, outb, outBytes, ret, (ptr)0);
}
static BOOL32 kureha_ReadFile(HANDLE32 h, ptr buf, u32 bytes, u32* got) {
    PFN_ReadFile f = (PFN_ReadFile)kureha_api("ReadFile");
    if (!f) { if (got) *got = 0; return 0; }
    return f(h, buf, bytes, got, (ptr)0);
}
static BOOL32 kureha_WriteFile(HANDLE32 h, const void* buf, u32 bytes, u32* done) {
    PFN_WriteFile f = (PFN_WriteFile)kureha_api("WriteFile");
    if (!f) { if (done) *done = 0; return 0; }
    return f(h, buf, bytes, done, (ptr)0);
}
static u32 kureha_SetFilePointer(HANDLE32 h, i32 lo, i32* hi, u32 method) {
    PFN_SetFilePointer f = (PFN_SetFilePointer)kureha_api("SetFilePointer");
    if (!f) return 0xffffffffu;
    return f(h, lo, hi, method);
}
static u32 kureha_GetFileSize(HANDLE32 h, u32* hi) {
    PFN_GetFileSize f = (PFN_GetFileSize)kureha_api("GetFileSize");
    if (!f) { if (hi) *hi = 0; return 0xffffffffu; }
    return f(h, hi);
}
static BOOL32 kureha_FlushFileBuffers(HANDLE32 h) {
    PFN_FlushFileBuffers f = (PFN_FlushFileBuffers)kureha_api("FlushFileBuffers");
    if (!f) return 0;
    return f(h);
}

#endif
