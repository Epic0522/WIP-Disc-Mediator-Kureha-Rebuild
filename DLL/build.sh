#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "$0")" && pwd)"
mkdir -p "$ROOT/bin" "$ROOT/build"
TMPDIR="$(mktemp -d "$ROOT/build/dllbuild.XXXXXX")"
trap 'rm -rf "$TMPDIR"' EXIT
BREW_LLVM="/opt/homebrew/opt/llvm/bin"
BREW_LLD="/opt/homebrew/opt/lld/bin"
if [[ -x "$BREW_LLVM/clang" ]]; then
    CLANG="$BREW_LLVM/clang"
else
    CLANG="${CLANG:-clang}"
fi
if [[ -d "$BREW_LLD" ]]; then
    export PATH="$BREW_LLD:$PATH"
fi
if [[ -d "$BREW_LLVM" ]]; then
    export PATH="$BREW_LLVM:$PATH"
fi
CFLAGS=(--target=i686-pc-windows-msvc -fuse-ld=lld -O2 -ffreestanding -fno-builtin -fno-stack-protector -nostdlib)
LDFLAGS=(-Wl,/dll -Wl,/noentry -Xlinker /subsystem:windows,5.01)

"$CLANG" "${CFLAGS[@]}" "${LDFLAGS[@]}" -Wl,/def:"$ROOT/src/momiji.def" "$ROOT/src/momiji_reimpl.c" -o "$TMPDIR/Momiji.dll" 2>"$ROOT/build/momiji_build.log"
"$CLANG" "${CFLAGS[@]}" "${LDFLAGS[@]}" -Wl,/def:"$ROOT/src/zenki.def" "$ROOT/src/zenki_reimpl.c" -o "$TMPDIR/Zenki.dll" 2>"$ROOT/build/zenki_build.log"

cp "$TMPDIR/Momiji.dll" "$ROOT/bin/Momiji.dll"
cp "$TMPDIR/Zenki.dll" "$ROOT/bin/Zenki.dll"

if command -v llvm-objdump >/dev/null 2>&1; then
    llvm-objdump -p "$ROOT/bin/Momiji.dll" > "$ROOT/build/momiji_objdump.txt"
    llvm-objdump -p "$ROOT/bin/Zenki.dll" > "$ROOT/build/zenki_objdump.txt"
else
    printf 'llvm-objdump not found; export dump skipped.\n' > "$ROOT/build/momiji_objdump.txt"
    printf 'llvm-objdump not found; export dump skipped.\n' > "$ROOT/build/zenki_objdump.txt"
fi
