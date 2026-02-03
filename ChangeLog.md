# Changelog

All notable changes to this project will be documented in this file.

The format follows a simplified version of Keep a Changelog.

---

## [1.0] – 2026-02-03

### Added
- Initial public release of RehaCompareGUI
- GUI-based folder selection
- Recursive relative-path comparison
- Optional filename-only comparison
- SHA-256 hash comparison
- Detection of same-content files with different paths/names
- Case number and operator metadata capture
- Local and UTC run timestamps
- Script self-hash recorded in Summary.txt
- Detailed Summary.txt output with header, body, and footer
- Handling and logging of locked/in-use files
- Progress bar and real-time log window

### Notes
- Designed for forensic review of provider-returned datasets
- Read-only analysis; no file modification
- PowerShell script is the authoritative implementation

---

## [Unreleased]

### Planned / Ideas
- Optional hashing of only non-overlapping files
- Consolidated CSV output for agent review
- README screenshots
- Optional script signing guidance
