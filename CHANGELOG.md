# Changelog

All notable changes to this project will be documented in this file.

The format is based on Keep a Changelog and is maintained automatically by Semantic Release.

## [0.1.10] - 2026-08-04

### Added

- Optional `mcp` extra that provides an MCP server directly from the `easy_uiauto` package.
- `easy_uiauto` and `easy_uiauto_service` command-line entry points with `--help` and `--version`.
- Batch action execution through the MCP `run_actions` tool and TCP client/service.
- MCP CLI regression tests and CI coverage for the `mcp` extra.

### Changed

- Moved MCP implementation to `easy_uiauto.mcp`; use `python -m easy_uiauto.mcp.server`.
- Made control-vector persistence opt-in through `EASY_UIAUTO_CONTROL_VECTOR_DB_DIR` instead of using a machine-specific path.

## [0.1.0] - 2026-03-16

### Added

- Initial public release of `easy-uiauto`
- `Controller` class for mouse and keyboard control
- Control location using XPath-based positioning strategy
- Action recording with visual feedback (ScreenLineBox)
- Utility functions for control caching and window management
- Support for multiple UI frameworks (Win32, Qt)
- CI, release, and PyPI publishing automation
