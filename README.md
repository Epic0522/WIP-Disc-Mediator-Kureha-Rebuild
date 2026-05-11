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
- `DLL/`: reverse-engineered dependency workspace for `Momiji.dll` and `Zenki.dll`, including binaries, export maps, ABI notes, C stubs, and VB6 declare modules

## DLL Reconstruction Notes

The repository also contains a dedicated `DLL/` workspace for the two native libraries required by the original application.

At the moment, the dependency work appears to split cleanly into two roles:

- `Momiji.dll`: drive, media, tray, TOC, LBA, read/write, speed, and device-control functions
- `Zenki.dll`: ISO filesystem authoring, track collection management, and track-text (`CD-TEXT`) related helpers

The current `DLL/` tree already includes:

- `bin/`: rebuilt or stubbed 32-bit DLL and import-library outputs
- `src/`: `.def`, `.h`, and `.c` source files for the reconstructed exports
- `analysis/`: export tables, disassembly notes, and ABI reconstruction artifacts
- VB6 declare modules for calling both DLLs from the rebuilt IDE project

This is important because the main VB6 rebuild can now be developed against a documented native boundary instead of treating those DLL calls as opaque black boxes.

## Planned Next Steps

The most likely path forward is:

1. keep refining the VB6 UI and workflow until it matches the original application's visible behavior closely enough for side-by-side testing
2. introduce a thin VB6 wrapper layer around the `Momiji` and `Zenki` declares so the rebuild stops calling raw DLL entry points directly
3. validate `Zenki` first, since track management, ISO layout, and `CD-TEXT` behavior are the most immediately testable pieces in the current UI
4. bring `Momiji` online after that for drive state, media interrogation, TOC transfer, and eventually write-path experiments
5. replace the current placeholder or stub behavior module by module as each reconstructed native call is verified

## Status

The current build is still in an early reconstruction stage, but the project has moved beyond a static mockup.

As of the current milestone:

- the rebuilt project opens and runs in the VB6 IDE without the original decompiler-related load failures
- the main window layout has been rebuilt into a resizable, testable scaffold that now roughly matches the original workflow
- track addition now uses the standard Windows file picker instead of a temporary text prompt
- the track `CD-TEXT` editor supports bilingual metadata editing and preserves Japanese text correctly in the rebuilt data model
- the lower track list has been cleaned up enough to validate titles, performers, source labels, gaps, durations, and flags during testing
- the disc usage area now includes a working preview ring and remaining-capacity estimate driven by imported track durations
- the required native dependency layer is now organized in-repo under `DLL/`, with reconstructed ABI notes and callable VB6 declare modules for both `Momiji.dll` and `Zenki.dll`

What is still missing is the original application's full authoring and burn pipeline: the current rebuild is suitable for UI, metadata, and workflow reconstruction, but not yet for final disc-writing equivalence.
