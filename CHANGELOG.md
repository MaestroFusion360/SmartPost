# Changelog

All notable changes to SmartPost are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [1.2.0.0] - 2026-08-21

### Added

- Added the current Autodesk `xml.cps` revision as `xml_last.cps` for the
  experimental Personal-license pipeline.
- Added the **Old xml.cps** option to switch back to the legacy intermediate
  postprocessor in Personal mode.
- Added explicit logging of the selected Personal intermediate postprocessor.
- Added Ruff and Pylint configuration and a reproducible development dependency
  list.
- Added reproducible PowerShell tests for comparing current and legacy XML
  cycle round-trips through unmodified downstream postprocessors.
- Added pytest regression coverage for Personal XML post selection, persisted
  booleans, Fusion output-unit mapping, XML merging, tool metadata, and grouped
  canned-cycle serialization.
- Added the optional `diagnosticSectionType` property to `xml_last.cps` for
  logging the source `currentSection.type` without modifying intermediate XML.
- Added a reproducible `test-xml-section-type.ps1` test for comparing turning
  CNC and XML imports through unchanged diagnostic and turning posts.

### Changed

- Split the oversized Smart Post command implementation into focused Fusion
  helper and Personal XML pipeline modules while preserving its public entry
  points and behavior.
- Made `xml_last.cps` the default intermediate postprocessor in Personal mode.
- Isolated intermediate XML, progress, and log files in a unique directory for
  each SmartPost run.
- Disabled `post.exe` backup generation with its native `--nobackup` option.
- Updated the README feature table to describe the tested Personal rapid-move
  behavior accurately.
- Generalized Git ignore rules for Python caches and local development files.

### Fixed

- Restricted postprocessor selection and persisted-path validation to existing
  `.cps` files, preventing archives and unrelated files from reaching Fusion or
  Autodesk `post.exe`.
- Removed XML 1.0-forbidden control characters from values serialized by both
  intermediate CPS variants, preventing corrupted Fusion parameters from
  causing `post.exe` load failures or stalls.
- Fixed tool metadata in `xml_last.cps` so downstream posts receive the actual
  Fusion tool type instead of `unspecified` / `UNKNOWN TOOL TYPE`.
- Fixed `xml_last.cps` parameter serialization for vector and array values used
  by drilling operations.
- Fixed grouped-cycle numeric parameter types so Autodesk's XML importer can
  restore cycle callbacks and section movement metadata without an internal
  error.
- Restored native canned drilling cycles in Personal mode through
  `xml_last.cps`, including verified G81, G83, G84, G86, and G87 output from the
  standard Autodesk Fanuc postprocessor.
- Fixed Personal intermediate post-processing passing UI unit indexes as Fusion
  API enums, which could convert metric cycle parameters to inches before the
  final metric post-processing stage.
- Fixed the **Old xml.cps** control dependency on **License Personal
  (Testing)**.
- Fixed persisted string values such as `"false"` being interpreted as enabled
  boolean options.
- Fixed cleanup of SmartPost-owned temporary files after successful and failed
  postprocessing without deleting unrelated user files.
- Fixed merged XML validation incorrectly treating a small buffered output file
  as empty before the file stream was flushed or closed.
- Separated actual `post.exe` timing from total batch-processing time.
- Fixed Python lint violations reported by Ruff and Pylint.
- Fixed Markdown table alignment warnings in the README.

### Known Limitations

- Any downstream `.cps` may exhibit severe performance degradation when it
  repeatedly calls `Section.getProperty()` or `Section.getParameter()` on
  sections reconstructed by the Autodesk XML importer. With Post Engine
  5.388.0, Autodesk Fanuc revision 44236 took about 30 seconds for the XML
  fixture versus less than one second after substituting its older
  `initializeSmoothing` implementation. No `post.exe` command-line option
  bypasses these section-scoped lookups; use a compatible post revision or a
  vendor-provided fix. SmartPost does not patch user-selected posts.
- Autodesk Post Engine 5.388.0 does not serialize or restore `Section.type`
  through its intermediate XML format. The XML importer does not parse
  attributes on `<section>`, so turning sections are reconstructed as the
  default `TYPE_MILLING`. This cannot be fixed in `xml_last.cps` alone.
- Manual NC is only partially represented by the Autodesk XML format. Comments
  and dwell records have supported XML elements, while Stop, Optional Stop, and
  Pass-through do not have a verified round-trip representation. An
  intermediate program containing only Manual NC records is rejected because
  the XML importer requires at least one section.
