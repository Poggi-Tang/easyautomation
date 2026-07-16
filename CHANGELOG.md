# Changelog

All notable changes to this project will be documented in this file.

The format is based on Keep a Changelog and is maintained automatically by Semantic Release.

## Unreleased

### Added

- Working `easy-uiauto` CLI and optional local stdio MCP server
- Dedicated MCP UIAutomation worker thread, JSON control snapshots, TTL references, and safety policy
- Native Win32 and Qt fixtures plus Explorer, desktop/taskbar, and reversible SimuNPS integration tests
- Semantic Toggle, Selection, Expand/Collapse, scroll, and image-control actions

### Changed

- XPath replay validates structure, uses exact depth, preserves desktop-root semantics, and verifies indexed cache entries
- Recording preserves click/drag offsets, handles consecutive shortcuts and horizontal scrolling, and rejects incomplete drags
- UI highlighting and recording cleanup tolerate transient command, listener, and Tk refresh failures
- Release workflow is manual; normal development pushes do not publish a Python package

### Fixed

- Unsafe `eval` parsing, cross-thread UIA object use, unconditional cache writes, and global search-timeout pollution
- Mouse release cleanup, background-window physical input, clipboard restoration, and failed-locator keyboard continuation
- False double-click merging, stale indexed cache reuse, parent-window image lookup, and Win+D lookup side effects
- Desktop/taskbar XPath depth, scoped combo-box selection, tri-state toggle verification, and text replacement fallback

## [0.1.0] - 2026-03-16

### Added

- Initial public release of `easy-uiauto`
- `Controller` class for mouse and keyboard control
- Control location using XPath-based positioning strategy
- Action recording with visual feedback (ScreenLineBox)
- Utility functions for control caching and window management
- Support for multiple UI frameworks (Win32, Qt)
- CI, release, and PyPI publishing automation
