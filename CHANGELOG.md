# Changelog

All notable changes to SmartPost are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [1.1.0.0] - 2026-08-21

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

### Changed

- Made `xml_last.cps` the default intermediate postprocessor in Personal mode.
- Isolated intermediate XML, progress, and log files in a unique directory for
  each SmartPost run.
- Disabled `post.exe` backup generation with its native `--nobackup` option.
- Updated the README feature table to describe the tested Personal rapid-move
  behavior accurately.
- Generalized Git ignore rules for Python caches and local development files.

### Fixed

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
- Fixed Python lint violations reported by Ruff and Pylint.
- Fixed Markdown table alignment warnings in the README.
