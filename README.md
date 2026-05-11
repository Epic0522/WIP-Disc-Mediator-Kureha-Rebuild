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

## Status

The current build is still in an early reconstruction stage, but the project has moved beyond a static mockup.

As of the current milestone:

- the rebuilt project opens and runs in the VB6 IDE without the original decompiler-related load failures
- the main window layout has been rebuilt into a resizable, testable scaffold that now roughly matches the original workflow
- track addition now uses the standard Windows file picker instead of a temporary text prompt
- the track `CD-TEXT` editor supports bilingual metadata editing and preserves Japanese text correctly in the rebuilt data model
- the lower track list has been cleaned up enough to validate titles, performers, source labels, gaps, durations, and flags during testing
- the disc usage area now includes a working preview ring and remaining-capacity estimate driven by imported track durations

What is still missing is the original application's full authoring and burn pipeline: the current rebuild is suitable for UI, metadata, and workflow reconstruction, but not yet for final disc-writing equivalence.
