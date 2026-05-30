# Kureha VB6 Rebuild

`Kureha VB6 Rebuild` is a clean-room reconstruction project for an older VB6-based CD authoring and disc utility application.

The long-term goal is not to patch decompiler output into something barely buildable, but to rebuild the software into a maintainable, compilable, and publishable VB6 codebase while preserving the original application's workflow and disc behavior as closely as possible.

## Project Purpose

This project exists mainly to study and reproduce how the original software handles Japanese `CD-TEXT` authoring and burning.

That includes:

- understanding how track and album text is organized internally
- reconstructing the original `TOC`, raw `TOC`, and `CD-TEXT` data model
- comparing generated structures against known-good sample discs
- studying when the original software chooses normal write paths versus raw write paths
- rebuilding the user interface and workflow so behavior can be validated step by step

## Why This Matters

Japanese `CD-TEXT` support is one of the most interesting parts of the original program. Older disc tools often expose only fragments of this workflow, while the target application appears to manage:

- bilingual track metadata
- disc and track level text fields
- raw disc layout details
- writing modes related to preserving subchannel and text data

Reconstructing those pieces in a clean VB6 project makes it easier to inspect, test, and eventually document the exact path from editable metadata to a burnable disc layout.

## Current Direction

The rebuild is being developed in stages:

1. recreate a VB6 project that opens cleanly in the IDE
2. rebuild the core data structures used for `TOC` and `CD-TEXT`
3. reconstruct the original user interface and workflow
4. fill in behavior module by module by studying the reversed application
5. validate output against reference samples and working burn paths

## Current Contents

- `Project.vbp`: the rebuilt VB6 project
- `MainForm.frm`: the main application scaffold
- `PropertyTrackForm.frm`: track `CD-TEXT` editor
- `DiscWriteForm.frm`, `DiscReadForm.frm`, `DiscCopyForm.frm`, `DiscEraseForm.frm`, `DiscAnalyzeForm.frm`: primary disc-operation dialogs
- `PropertyConfigurationForm.frm`, `PropertyISOFile.frm`, `PropertyReadParameterForm.frm`, `PropertyWriteParameterForm.frm`: configuration and property dialogs
- `TrackRippingForm.frm`, `TrackRippingSubForm.frm`, `WellSaveForm.frm`, `WellWriteForm.frm`, `IsReadyForm.frm`, `ListStatusForm.frm`: ripping, image, media-ready, and status/progress scaffolds
- `TOCTrack.cls`: track model and time helpers
- `TOCRawTrack.cls`: raw `TOC` entry model
- `TOCCDText.cls`: bilingual `CD-TEXT` field container
- `TOCInformation.cls`: aggregate disc metadata container
- `TrackEntry.cls`: track-facing UI model
- `FileEntry.cls`: file list UI model
- `DLL/`: current working dependency workspace for reconstructed `Momiji.dll` and `Zenki.dll`
- `DLL_stub/`: first-pass static-reconstruction workspace used to map exports, calling conventions, and early stub behavior
- `tools/analyze_raw96.py`: helper for inspecting RAW+96 test images, exporting simulated capture bundles, capturing real discs through `drutil`/`cdrdao`, and comparing bundle-shaped outputs

## DLL Reconstruction Notes

The public repository contains two DLL-focused workspaces that represent different stages of the native dependency reconstruction effort.

- `DLL_stub/`: the initial static reverse-engineering pass, where exports, stdcall stack cleanup, ABI guesses, and minimal stub implementations were established
- `DLL/`: the current active workspace, where those early ABI notes have been promoted into rebuilt DLLs, VB6 wrapper classes, and a more realistic reimplementation layer

The untouched original vendor DLLs are intentionally kept out of the public repository and used only as a private local reference during analysis.

At the moment, the dependency work appears to split cleanly into two roles:

- `Momiji.dll`: drive, media, tray, TOC, LBA, read/write, speed, and device-control functions
- `Zenki.dll`: ISO filesystem authoring, track collection management, and track-text (`CD-TEXT`) related helpers

The current `DLL/` tree already includes:

- `bin/`: rebuilt 32-bit DLL and import-library outputs
- `src/`: `.def` files and active native reimplementation sources (`momiji_reimpl.c`, `zenki_reimpl.c`)
- `build/`: build logs and post-build object-dump output
- `vb/`: VB6 declare module plus thin wrapper classes (`MomijiEngine.cls`, `ZenkiEngine.cls`)
- `analysis/`: ABI comparison artifacts including the original export map

This is important because the main VB6 rebuild can now be developed against a documented native boundary instead of treating those DLL calls as opaque black boxes.

The most significant change in the current state is that the rebuilt DLLs no longer appear to be simple passive stubs. The project now has:

- a static reverse-engineered ABI baseline from `DLL_stub/`
- a newer active `DLL/` implementation that is intended to receive calls from the decompiled or rebuilt VB6 side and forward them into real device-facing behavior
- a safer build script that writes new DLLs to a temporary directory first and only replaces `DLL/bin/*.dll` after a successful link

That means the next stage is no longer just export reconstruction. It is integration and behavioral verification.

The active native sources are currently being developed with an import-free / no-CRT design. They resolve the small Win32 surface they need at runtime, so the final DLLs should not require a Visual C++ runtime redistributable on the XP test system. The intended build path is to use LLVM/`lld-link` from the development machine, then copy only the generated `DLL/bin/Momiji.dll` and `DLL/bin/Zenki.dll` into the VB6 test directory.

## Planned Next Steps

The most likely path forward is:

1. validate the new `DLL/vb` wrapper classes against the rebuilt DLLs before wiring them deeper into the main project
2. verify the end-to-end call surface first in safe read-only areas: version queries, initialization, track enumeration, ISO state, and `CD-TEXT` getters/setters
3. validate `Zenki` first, since track management, ISO layout, and text behavior are the easiest parts to exercise from the current UI without touching hardware writes
4. validate `Momiji` next in progressively higher-risk stages: device open/close, media interrogation, TOC transfer, read path, then write-related operations
5. keep comparing behavior across the two public workspaces while validating against a private local original-binary reference outside the repository
6. only after that replace the remaining UI-side placeholder logic with real DLL-backed operations module by module

## Capture Bundle Workflow

The project now compares generated output and physical-disc captures through the same capture-bundle shape instead of comparing incompatible raw files directly. Capture bundles are local analysis artifacts and are intentionally ignored through `captures/` because they can contain real disc audio.

To inspect a generated RAW+96 debug image:

```sh
./tools/analyze_raw96.py analyze-bin path/to/kureha_filebacked_raw96_disc_test.bin
```

To export a simulated capture bundle from a generated RAW+96 debug image:

```sh
./tools/analyze_raw96.py bundle-from-raw96 path/to/kureha_filebacked_raw96_disc_test.bin captures/simulated_current
./tools/analyze_raw96.py analyze-capture captures/simulated_current
```

To preview the physical-disc capture commands before connecting or using a drive:

```sh
./tools/analyze_raw96.py capture-disc captures --name original_good_disc --dry-run
```

To capture a known-good disc from a real drive once hardware is connected:

```sh
./tools/analyze_raw96.py capture-disc captures --name original_good_disc
./tools/analyze_raw96.py analyze-capture captures/original_good_disc
```

To compare the known-good physical-disc bundle against the simulated bundle:

```sh
./tools/analyze_raw96.py compare-capture captures/original_good_disc captures/simulated_current
```

The main comparison target is the bundle structure: `disc.toc`, `data.bin`, optional `subchannel.bin`, `metadata.json`, and optional `raw96-debug.bin`. `raw96-debug.bin` is useful for internal diagnostics, but it should not be compared directly against a physical-disc `data.bin`.

## Status

The current build is still in an early reconstruction stage, but the project has moved beyond a static mockup.

As of the current milestone:

- the rebuilt project opens and runs in the VB6 IDE without the original decompiler-related load failures
- the main window layout has been rebuilt into a resizable, testable scaffold that now roughly matches the original workflow
- track addition now uses the standard Windows file picker instead of a temporary text prompt
- file and folder addition now use native Windows picker dialogs, retain original source paths, and persist those paths through project save/load
- the track `CD-TEXT` editor supports bilingual metadata editing and preserves Japanese text correctly in the rebuilt data model
- the lower track list has been cleaned up enough to validate titles, performers, source labels, gaps, durations, and flags during testing
- the disc usage area now includes a working preview ring and remaining-capacity estimate driven by imported track durations
- basic ISO file-list actions now synchronize into the rebuilt `Zenki.dll` layer, including real files, dummy files, directories, clears, removals, and renames
- the required native dependency layer is now organized in-repo as `DLL_stub/` and `DLL/`, preserving both the first static reverse-engineering pass and the current active reimplementation workspace
- the current `DLL/` workspace already includes rebuilt DLL outputs, native reimplementation sources, and VB6 wrapper classes intended to receive and relay calls from the application side
- the rebuilt VB6 project now has form-level coverage for the original application's main window, disc-operation dialogs, configuration/property dialogs, image save/write progress windows, media-ready prompt, list status dialog, and track-ripping subwindow
- the major menu and toolbar entry points now open concrete rebuilt windows instead of stopping at a single placeholder message, so the next development pass can debug behavior inside each screen
- view-menu quality-of-life actions have started to behave like application features: always-on-top now toggles through the Win32 API, Explorer opens at the selected file source or project folder, and the mastering log reports current project/DLL state
- `Zenki` now has a more complete in-memory ISO model: normalized ISO paths, directory parent/child relationships, multi-directory directory sectors, path table generation, multi-sector path table layout, file/dummy/directory enumeration, and safer remove/rename/property behavior
- `Zenki` track reading now understands ordinary file tracks and WAV data chunks, can stream track data with pregap/postgap handling, and can wrap 2048-byte sectors into raw 2352-byte sectors for raw-mode reads
- `Zenki` `CD-TEXT` storage now tracks whether each field was actually set, keeps album/track text separated by track number, and can return track-title text through `GetTrackInformation`
- `Momiji` now has a safer file-backed test mode for non-hardware verification: ordinary files can be opened as simulated CD-R targets, written by LBA, flushed, erased, queried for empty/non-empty state, and inspected for last-written LBA/cache-used state
- the rebuilt DLLs can be loaded from the VB6 IDE test directory through the `DLL/bin` path, and the VB6 wrapper layer can call into both `Zenki.dll` and `Momiji.dll`
- file-backed simulated writing has progressed from a short sector-transfer smoke test into a DAO RAW+SUB96 semantic test image: the generated image now contains lead-in, program area, and lead-out regions
- `CD-TEXT` is now emitted into the simulated RAW+96 lead-in area as Shift-JIS-compatible bytes, and the returned test image successfully decodes a Japanese track title from the generated packs
- the `tools/analyze_raw96.py` helper validates the returned test image structure, including sector count, logical LBA range, lead-in/program/lead-out samples, Q subchannel CRC, `CD-TEXT` pack CRC, and decoded Japanese text
- the latest verified file-backed test image produced `4500` lead-in sectors, `4922` program sectors, `6750` lead-out sectors, zero sampled Q CRC failures, zero `CD-TEXT` CRC failures, and decoded the title `アンティーカ & ストレイライト - Killer×Mission`
- comparison is being moved to a capture-bundle model so simulated output and physical-disc captures share the same shape: `disc.toc`, `data.bin`, optional `subchannel.bin`, `metadata.json`, and optional `raw96-debug.bin`
- `captures/` is intentionally local-only because physical-disc capture bundles can contain real audio data

What is still missing is the original application's full authoring and burn pipeline. The current rebuild can validate metadata, `CD-TEXT`, and a DAO RAW+SUB96-shaped file-backed image, but the program area still needs true audio-sector generation from the original application's supported input scope: 44.1 kHz / 16-bit MP3 or WAV sources converted into proper CDDA PCM sectors before any real hardware write path should be attempted.
