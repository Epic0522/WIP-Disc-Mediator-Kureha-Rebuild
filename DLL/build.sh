#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "$0")" && pwd)"
mkdir -p "$ROOT/bin" "$ROOT/build"
CFLAGS=(--target=i686-pc-windows-msvc -fuse-ld=lld -O2 -ffreestanding -fno-builtin -fno-stack-protector -nostdlib)
LDFLAGS=(-Wl,/dll -Wl,/noentry)
clang "${CFLAGS[@]}" "${LDFLAGS[@]}" -Wl,/def:"$ROOT/src/momiji.def" "$ROOT/src/momiji_reimpl.c" -o "$ROOT/bin/Momiji.dll" 2>"$ROOT/build/momiji_build.log"
clang "${CFLAGS[@]}" "${LDFLAGS[@]}" -Wl,/def:"$ROOT/src/zenki.def" "$ROOT/src/zenki_reimpl.c" -o "$ROOT/bin/Zenki.dll" 2>"$ROOT/build/zenki_build.log"
llvm-objdump -p "$ROOT/bin/Momiji.dll" > "$ROOT/build/momiji_objdump.txt"
llvm-objdump -p "$ROOT/bin/Zenki.dll" > "$ROOT/build/zenki_objdump.txt"
