# Changelog

All notable changes to this project will be documented in this file.

The format is based on Keep a Changelog and is maintained automatically by Semantic Release.

## [0.1.20] - 2026-08-10

### Added

- Added `--quick-setup-codex` for a minimal remote-vision deployment with hidden API-key input, user-scoped environment persistence, idempotent MCP replacement, and no OCR or UI test overhead.

## [0.1.19] - 2026-08-10

### Changed

- Added installation checks, MCP smoke tests, source test commands, and OCR/AI vision diagnostics to `easy_uiauto --help`.

## [0.1.18] - 2026-08-10

### Changed

- Expanded `easy_uiauto --help` with LOCATION, recording, replay, batch, mode, visual fallback, and MCP client configuration workflows.

## [0.1.17] - 2026-08-10

### Added

- Let MCP and TCP `find_control` accept the canonical `LOCATION` object from recorded actions or coordinate inspection.
- Include a reusable canonical `LOCATION` in coordinate and control-capture results.

### Fixed

- Preserve `foundIndex` and `searchDepth` when converting captured XPath records.
- Keep legacy flat control lookup from failing when XPath is empty.
- Read `ControlType` correctly when disassembling a location.

## [0.1.16] - 2026-08-10

### Fixed

- Send an explicit HTTP User-Agent for remote vision API requests so channels that reject the Python default client can accept them.

## [0.1.15] - 2026-08-10

### Added

- Optional `mcp` extra that provides an MCP server directly from the `easy_uiauto` package.
- `easy_uiauto` and `easy_uiauto_service` command-line entry points with `--help` and `--version`.
- Batch action execution through the MCP `run_actions` tool and TCP client/service.
- MCP CLI regression tests and CI coverage for the `mcp` extra.
- Opt-in commands to install, inspect, and uninstall the MCP server in Codex and Claude Code.
- Visual fallback MCP tools for image templates, OCR text, and remote multimodal API location.

### Changed

- Moved MCP implementation to `easy_uiauto.mcp`; use `python -m easy_uiauto.mcp.server`.
- Made control-vector persistence opt-in through `EASY_UIAUTO_CONTROL_VECTOR_DB_DIR` instead of using a machine-specific path.
- Replaced product-specific examples with generic application and control names.
- Added the optional `vision` dependency group for OpenCV template matching and OCR.

## [0.1.0] - 2026-03-16

### Added

- Initial public release of `easy-uiauto`
- `Controller` class for mouse and keyboard control
- Control location using XPath-based positioning strategy
- Action recording with visual feedback (ScreenLineBox)
- Utility functions for control caching and window management
- Support for multiple UI frameworks (Win32, Qt)
- CI, release, and PyPI publishing automation
