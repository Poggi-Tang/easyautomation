# Changelog

All notable changes to this project will be documented in this file.

The format is based on Keep a Changelog and is maintained automatically by Semantic Release.

## [0.6.0] - 2026-08-13

### Added

- Added a single-window, click-through overlay that can draw and number a complete page
  of controls without creating four tracking windows per control.
- Added numbered annotated page images for every learning scan and returned their vault
  paths in scan results.
- Added `show_ui_controls` and `easy_uiauto_ui show` to match the current learned page,
  box its executable or known controls, and return the corresponding command legend.
- Added target previews before single and batch semantic UI command execution, with
  configurable duration and an option to disable highlighting.
- Added MCP progress heartbeats, per-control verification counts, persistent in-process
  learning task status, and `get_ui_learning_status` for distinguishing a client wait
  timeout from a failed or still-running scan.
- Added automatic migration from absolute desktop rectangles and process-local window
  identifiers to window-relative visual hints.

### Changed

- Updated the bundled learning and operation Skills to show learned controls, answer
  current-page capability questions visually, and highlight verified runtime targets.
- Ordinary learn or scan requests now stop after the basic page scan. Recursive safe
  interaction runs only for an explicit deep-learning or autonomous-exploration request.
- Scan verification reuses shared XPath prefixes, and scan/exploration workers continue
  recording status after an MCP client wait timeout without blocking status queries.
- Knowledge records no longer persist PIDs, window handles, or absolute screen rectangles.
  Normalized geometry is marked as a visual hint only; live overlays resolve current
  `LOCATION`, image-template, or OCR positions before drawing.

## [0.5.1] - 2026-08-12

### Added

- Added `get_ui_learning_readiness` for a fast, non-network preflight that reports
  learning configuration without exposing API credentials.

### Fixed

- Read vision settings dynamically for every MCP request so configuration changes are
  visible to an already-running current-version server.
- Stop the learning Skill immediately on configuration, authentication, network, timeout,
  or model failures instead of retrying the slower `full-uia` strategy.
- Clarified that Codex needs one restart for package, MCP registration, or Skill updates,
  but not a second restart for vision environment changes.

## [0.5.0] - 2026-08-12

### Added

- Added visual-first learning that asks one multimodal call for page regions and key
  controls, then maps visual targets to stable local UIA controls without a full tree walk.
- Added `learn_ui_command_effect`, `explore_ui_workflows`, and matching
  `easy_uiauto_ui learn-effect`, `explore`, and `interactions` commands.
- Added bounded recursive exploration of newly opened pages and dialogs with
  page/command deduplication and inner-to-outer state recovery.
- Added local before/after target and desktop snapshots, localized pixel differences,
  delayed stability checks, changed-control inspection, top-level popup/transient-window
  detection, action-property changes, success conditions, state IDs, and Escape recovery.
- Added multi-state control templates and runtime fallback through LOCATION, image,
  OCR, and opt-in remote AI vision.
- Added right-click and hover commands and effect learning. Scrolling and dragging remain
  intentionally excluded from automated exploration.
- Added virtual-desktop window capture for applications on secondary monitors and streaming
  completion parsing compatible with OpenAI SSE responses and empty usage frames.

### Changed

- Store interaction knowledge as Obsidian-compatible Markdown and PNG source files while
  keeping the JSON index disposable.
- Send only target-window before/after images and related popup crops for effect analysis;
  complete desktop snapshots stay local.
- Stop safe exploration on user interference or failed state recovery and never
  automatically execute external, destructive, or confirmation-required commands.
- Keep visual-first scanning to one remote request. An empty visual target list no longer
  triggers a full UIA traversal and second semantic-model request.

## [0.4.0] - 2026-08-12

### Added

- Added `run_ui_commands` and `easy_uiauto_ui batch` for ordered, structured
  same-page semantic UI command sequences.
- Added per-batch timing, completed-step, unique-control, and knowledge-write
  diagnostics.

### Changed

- Batch execution now loads knowledge once, shares one window screenshot across
  all unique-control checks, completes every preflight before the first action,
  and rebuilds the knowledge index once.
- Batch clicks skip UIAutomation cursor animation and default per-click waits;
  single-command execution keeps its existing interaction behavior.
- The bundled operation Skill now groups stable same-page workflows and splits
  batches at page navigation boundaries.

## [0.3.0] - 2026-08-12

### Added

- Added batched multimodal understanding of control meaning using the original page,
  numbered control overlays, UIA metadata, region context, and supported actions.
- Added semantic intent, descriptions, aliases, evidence, ambiguity, confidence, risk,
  confirmation requirements, and separate function-verification metadata to control knowledge.
- Added `teach_ui_control` and `easy_uiauto_ui teach` for human corrections that survive
  rescans without bypassing LOCATION or image verification.
- Added explicit confirmation for external or destructive semantic UI commands.

### Changed

- Require high-confidence or human-taught semantics in addition to locator/image validation
  before publishing executable UI commands.
- Configure pytest to import the working `src` tree so tests cannot silently exercise an
  older installed package.

## [0.2.0] - 2026-08-12

### Added

- Added an Obsidian-compatible application UI knowledge vault using Markdown/YAML and PNG files as the source of truth, with a disposable JSON search index.
- Added full-window UIA scanning, remote AI page/region segmentation, per-control image capture, stable page/region identities, and scan completeness reporting.
- Added LOCATION and unique image validation, automatic quarantine of stale or ambiguous controls, and rescan-based repair.
- Added generated application-specific UI CLI commands through `easy_uiauto_ui` and equivalent MCP tools.
- Added bundled `easy-uiauto-learning` and `easy-uiauto-operate` Codex skills with installation and update commands.

## [0.1.21] - 2026-08-11

### Added

- Added `--full-setup-codex` to install missing vision requirements and Tesseract, configure the global Codex MCP entry, and time deterministic UIA, OCR, and remote AI vision validation.
- Added synthetic-image OCR and multimodal diagnostics so deployment validation does not expose the current desktop.

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
