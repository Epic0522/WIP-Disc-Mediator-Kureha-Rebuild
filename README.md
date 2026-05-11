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
- `TOCTrack.cls`: track model and time helpers
- `TOCRawTrack.cls`: raw `TOC` entry model
- `TOCCDText.cls`: bilingual `CD-TEXT` field container
- `TOCInformation.cls`: aggregate disc metadata container
- `TrackEntry.cls`: track-facing UI model
- `FileEntry.cls`: file list UI model
- `DLL/`: current working dependency workspace for reconstructed `Momiji.dll` and `Zenki.dll`
- `DLL_stub/`: first-pass static-reconstruction workspace used to map exports, calling conventions, and early stub behavior

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

That means the next stage is no longer just export reconstruction. It is integration and behavioral verification.

## Planned Next Steps

The most likely path forward is:

1. validate the new `DLL/vb` wrapper classes against the rebuilt DLLs before wiring them deeper into the main project
2. verify the end-to-end call surface first in safe read-only areas: version queries, initialization, track enumeration, ISO state, and `CD-TEXT` getters/setters
3. validate `Zenki` first, since track management, ISO layout, and text behavior are the easiest parts to exercise from the current UI without touching hardware writes
4. validate `Momiji` next in progressively higher-risk stages: device open/close, media interrogation, TOC transfer, read path, then write-related operations
5. keep comparing behavior across the two public workspaces while validating against a private local original-binary reference outside the repository
6. only after that replace the remaining UI-side placeholder logic with real DLL-backed operations module by module

## Status

The current build is still in an early reconstruction stage, but the project has moved beyond a static mockup.

As of the current milestone:

- the rebuilt project opens and runs in the VB6 IDE without the original decompiler-related load failures
- the main window layout has been rebuilt into a resizable, testable scaffold that now roughly matches the original workflow
- track addition now uses the standard Windows file picker instead of a temporary text prompt
- the track `CD-TEXT` editor supports bilingual metadata editing and preserves Japanese text correctly in the rebuilt data model
- the lower track list has been cleaned up enough to validate titles, performers, source labels, gaps, durations, and flags during testing
- the disc usage area now includes a working preview ring and remaining-capacity estimate driven by imported track durations
- the required native dependency layer is now organized in-repo as `DLL_stub/` and `DLL/`, preserving both the first static reverse-engineering pass and the current active reimplementation workspace
- the current `DLL/` workspace already includes rebuilt DLL outputs, native reimplementation sources, and VB6 wrapper classes intended to receive and relay calls from the application side

What is still missing is the original application's full authoring and burn pipeline: the current rebuild is suitable for UI, metadata, and workflow reconstruction, but not yet for final disc-writing equivalence.
