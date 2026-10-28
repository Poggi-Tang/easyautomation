# Changelog

All notable changes to this project will be documented in this file.

The format is based on Keep a Changelog and is maintained automatically by Semantic Release.

## [Unreleased]

### Added

- Add a standalone Windows Tk configuration EXE and `easy_uiauto_config` GUI
  entry point for the knowledge directory, API URL, visible API Key and editable
  model dropdown. Fetch the provider's model list in a background thread.
- Persist the four settings to the Windows user environment. Read the knowledge
  directory dynamically, matching the existing API configuration behavior.
- Add a reproducible PyInstaller build script for the standalone EXE.
- Bundle the Windows x64 configuration EXE in wheels and source distributions,
  launch it through `easy_uiauto_config`, and verify package contents in CI and
  the tag publication workflow before uploading distributions.

## [0.9.8] - 2026-10-28

### Fixed

- Align automation test doubles with the shared Controller/uiautomation input,
  preview lifecycle and foreground activation interfaces introduced in 0.9.7.
- Check that an unconfirmed ValuePattern write is never repeated via clipboard
  paste. Keep the existing Windows test matrix as the PyPI publication gate.

## [0.9.7] - 2026-10-28

### Changed

- Route learned UI commands and recorder replay through shared Controller actions
  backed by uiautomation, including verified targets, coordinate targets, keyboard,
  text clipboard, recovery input and window activation.
- Show Controller/ScreenLineBox target previews during deep exploration by default.
  Acknowledge preview visibility and closure before screenshots; expose
  `highlight=false` / `--no-highlight` for effect learning and exploration.
- Preserve action failures and distinguish attempted input from confirmed results.
  Do not repeat an unconfirmed ValuePattern write using a second input method.
- Remove emoji decorations from the English and Chinese READMEs.
- Refresh the contact QR code image used by both READMEs.

### Fixed

- Bound recorder highlight queue processing, coalesce display updates, cancel Tk
  callbacks on shutdown, guard joins of unstarted UI threads and always signal
  cleanup completion. Keep recording independent of the UI knowledge vault.
- Initialize UIAutomation in the MCP background scan/exploration threads.

## [0.9.6] - 2026-10-26

### Fixed

- Restore live recorder highlighting by tracking UIA controls in the recorder
  process, so the outline follows each selected control and avoids cross-process
  XPath lookups.
- Cancel pending highlight tracking callbacks during cleanup so ESC closes the
  highlight window cleanly.

## [0.9.5] - 2026-10-26

### Fixed

- Start the isolated highlight worker without hiding the active Python
  interpreter's configured site-packages. This supports PyCharm environments
  that install runtime dependencies in the user site while still loading the
  exact easy-uiauto source package selected by the parent process.

## [0.9.4] - 2026-10-26

### Fixed

- Keep recording responsive when cross-process XPath tracking cannot resolve
  a live UIA window. Send a serializable control rectangle first so the
  highlight is shown immediately while the recorded mouse and keyboard
  listeners continue working.

## [0.9.3] - 2026-10-26

### Fixed

- Ship the recording highlighter and its worker together. Run Tk in a private
  interpreter and bootstrap the same installed package under isolated Python,
  including editable installations and calls from unrelated working directories.
- Detect missing workers and early child exits immediately, preserve startup
  errors, and stop/join only listeners that actually started. Always clean up the
  overlay, make repeated stop calls harmless and propagate recording failures
  to the caller. Honor `run_record(write_file=False)` without creating a file.
- Keep highlight bars click-through and non-activating. Use native window hit
  testing for occlusion instead of repeatedly discovering unrelated UIA controls.

## [0.9.2] - 2026-10-26

### Fixed

- Export all MCP tools from `easy_uiauto.mcp` with editor-visible signatures,
  lazy loading and optional progress contexts for ordinary async Python calls.
  Add import documentation and a runnable example. Resolve named controls and
  UIA readiness probes through native window handles instead of desktop traversal.
- Match contiguous OCR phrases across Tesseract word boxes without joining
  separate rows, distant labels or low-confidence words. Keep omitted drag
  source/destination XPaths as valid empty arrays.
- Inspect control trees through Win32-selected window handles in disposable UIA
  workers. Enforce a total node/time budget, stop child enumeration at the depth
  limit, report partial checkpoints on timeout and preserve foreground/layout.
- Bound control image states on disk, deduplicate new interaction PNGs by content,
  and automatically expire old interaction image sets while retaining textual
  history. Preserve referenced evidence, manual retention and recent failures;
  serialize cooperating writers and collection across MCP processes. Add
  `cleanup_ui_knowledge_images` with a dry-run default for existing vaults.
- Cross-check native action labels with visual semantics before generating or
  running commands. Preserve conflicting observations without executing their
  inferred intent; recover named menu actions from bounded native UIA evidence.
- Limit interactive candidates to controls observed in the current scan, count
  parent actions toward the exploration budget, and invalidate learned effects
  when a control's intent changes.
- Recover through an enabled owned modal when its owner is disabled; do not send
  browser-back keys as a generic recovery step. Preserve the application's main
  window name while learning owned dialogs.
- Extend native identity checks to checkbox/radio actions and labeled search /
  replacement inputs. Reject unresolved known action identities and recheck
  parent candidate lists after recursive exploration.
- Preserve runtime quarantine and failure evidence across rescans, semantic
  teaching and index rebuilds. Add explicit non-mutating control revalidation.
- Skip previously attempted interactions during automatic exploration, preserve
  learned effect evidence, exclude window-state controls and stop on interference.
- Restore only minimized windows during scanner activation; preserve maximized
  layout and verify that the target actually obtained the foreground.
- Skip disabled actions and new-document/window commands without an Escape
  reversal during exploration. Withhold parent recovery input after nested
  discovery stops, including interference and failed recovery.

## [0.9.1] - 2026-10-23

### Fixed

- Find learning targets, desktop shell roots and new popup handles with Win32
  enumeration before binding target UIA controls. Read window-state snapshots
  directly from Win32, avoiding desktop UIA root enumeration that can block on
  unrelated unresponsive providers. Report window lookup as a separate stage.
- Recover colored diagram bodies with hue-independent foreground components,
  including multicolor symbols, polygons and adjacent closed ports. Keep OCR
  text and neutral connecting wires separate from component bounds.
- Suppress busbar details misread as grounds and caption-only wire-loop tags;
  detect neutral mux bars and tag outlines interrupted by touching text.
- Retain unmatched YOLO graphic proposals with non-text pixel evidence and
  expose candidate rejection reasons, color evidence and fusion timings.

## [0.9.0] - 2026-10-23

### Added

- Add a local `easy_uiauto_image` CLI and checkout CMD launcher, exporting text,
  graphic and combined annotations, coordinates, raw detections, logs and an HTML
  viewer using cached OmniParser YOLOv9-E and RapidOCR PP-OCRv6 models on CPU.
- Add `analyze_ui_image` for saved UI screenshots, returning numbered annotations,
  semantic observations, pixel/normalized rectangles, centers, and a JSON artifact
  without desktop access, UIA verification, generated commands, or vault updates.
- Add local `correct_ui_image_annotations` with selected-icon pixel fitting and
  explicit rectangles; preserve IDs, raw model evidence, styles and old artifacts.
- Add circuit-symbol correction from original pixels, geometry-based duplicate
  fusion and opt-in recovery of missing busbars without reference annotations.

### Fixed

- Recover missed small text inside closed control frames using cached detector
  proposals and agreeing original/enlarged recognition; preserve baseline OCR IDs,
  reject connected strokes and export recovery evidence and timing.
- Distinguish saved-image graphic/text boxes with blue `G` and orange `T` labels;
  keep labels adjacent to their own edges and record canvas offsets when needed.
- Preserve text as independent controls, including labels/values overlapping
  graphics; keep text observations separate from circuit-outline fitting.
- Avoid near-identical trailing image tiles; preserve cross-tile evidence while
  deduplicating strongly overlapping geometry despite model role/type variation.
- Fit complete block-diagram outlines, tags, ellipses, terminals and grounds;
  exclude frame shadows from busbar discovery and report unresolved predictions.
- Draw one-pixel image annotations and compact labels after preview enlargement,
  retain raw model bounds, and include margins/known attached parts in full bounds.
- Prefer complete detections over clipped duplicates at screenshot tile boundaries.
- Analyze large screenshots in overlapping native-resolution tiles and supplement
  desktop coverage with Explorer's exposed items, including secondary monitors.
  Use native icon inventory and visual taskbar analysis for desktops; exclude
  occluded icons with a screen hit test instead of relying on UIA IsOffscreen.
- Reject oversized UIA ancestor matches; deduplicate only compatible overlapping
  visual detections and retain distinct entities sharing a container separately.
- Draw independent observation markers for saved visual components and preserve
  multiple repeated visual entities without overwriting their records.
- Resolve desktop locators across Explorer shell roots and never substitute an
  ancestor AutomationId for the requested leaf control.
- Isolate Win32 callback signatures from pygetwindow and detach readiness probes
  from MCP stdin to avoid Windows enumeration failures and probe timeouts.

## [0.8.0] - 2026-10-13

### Added

- Read-only MCP tools to wait for control presence, absence and enabled state,
  verify accessible text, and observe window changes from an explicit baseline.
- Separate readiness reports for UIA, OCR, image matching and remote vision.
- Typed strict control resolution and portable regression tests.

### Changed

- MCP desktop actions and locator tools return a consistent JSON outcome with
  `ok`, `code`, `message`, `verified`, `retryable` and `data`. Existing plain-text
  responses move to `message`; existing JSON payload fields remain available.
- Distinguish action execution from verified effects. Ambiguous default selectors
  require additional context or an explicit index instead of choosing the first match.
- `mouse_scroll` is a compatibility alias for `scroll`, including coordinates.

### Fixed

- Correct reversed scroll direction and propagate scrolling failures.
- Stop action batches on structured failures and preserve completed-step evidence.
- Do not send keyboard input when the requested window cannot be activated.
- Restore tag-based PyPI publishing of the tagged commit, gated by Windows tests.

## [0.7.5] - 2026-09-28

### Changed

- UIAutomation mouse and keyboard actions now pass zero operation wait time.
- Explicit mouse movements and drag actions skip gradual cursor movement, and
  text input uses zero inter-key interval.
- MCP UI clicks use zero operation wait time in both regular and fast modes.

## [0.7.4] - 2026-09-28

### Changed

- All UIAutomation control clicks and cursor moves now pass `simulateMove=False`
  by default, including text and keyboard focus actions and the MCP UI CLI.
- Window activation now locates visible top-level windows through Win32 enumeration
  and restores minimized windows before bringing them to the foreground.

## [0.7.3] - 2026-09-28

### Changed

- Top-level window lookup now checks `ClassName` and `Name` with `WindowControl`
  first, then retries by `Name` before returning `False`.
- AutomationId lookup verifies that the selector is unique within the window;
  duplicate IDs fall through to XPath, and ambiguous property fallbacks are rejected.

## [0.7.2] - 2026-08-13

### Changed

- Control lookup now uses window-scoped `AutomationId` first, contextual XPath second,
  and the `Name`/`ClassName`/`ControlType` property combination third.
- AutomationId lookup searches the complete current window instead of trusting a recorded
  XPath depth or dynamic `foundIndex`.
- Added the public `easy_uiauto.visual.segment_screen_rect` and `segment_image` Python APIs
  so local rectangle segmentation does not require MCP setup or imports.

### Fixed

- Empty XPath input no longer resolves to the desktop root as a false successful control.
- Public `find_control` now enumerates top-level UIA windows directly, avoiding slow or
  unsuccessful descendant queries for Qt application windows.
- Runtime command lookup falls back to a unique property combination only after both
  AutomationId and XPath fail.

## [0.7.1] - 2026-08-13

### Added

- Added a size-independent layout graph to `detect_visual_elements`, including sibling
  `right-of` and `below` relations, edge alignment, projection overlap, and adaptive gaps.
- Added variable-size composite groups whose bounds are recomputed from current members and
  whose layout signatures exclude element width, height, and observed content.

### Fixed

- Variable-width names and variable-height multiline content no longer need fixed geometry to
  belong to the same visual structure.
- Vertically repeated list items and adjacent table rows are kept separate instead of being
  merged into one oversized visual group.

## [0.7.0] - 2026-08-13

### Added

- Added `detect_visual_elements`, a local class-agnostic rectangle detector for coarse UIA
  canvases whose internally drawn elements have no UIA nodes.
- The detector accepts direct screen bounds or `find_control` / `get_control_at_position`
  results, captures only that rectangle, and returns absolute rectangles, parent-relative
  normalized rectangles, containment relationships, confidence, and a numbered preview.

### Changed

- Visual canvas segmentation runs locally without UIA traversal, OCR, remote AI, semantic
  classification, or model downloads. Multi-mask rectangle proposals are merged and small
  internal texture or character fragments are suppressed.

## [0.6.1] - 2026-08-13

### Added

- Added window-scoped AutomationId fast lookup with uniqueness checking before XPath traversal.
- Added `search_ui_knowledge_batch` for retrieving all controls needed by a multi-step workflow
  in one MCP call.
- Added semantic entity, observed value, dynamic-context, current-state, state-reason,
  enabling-condition, relationship, and visual-component knowledge fields.
- Visual-only names, avatars, badges, content areas, and status elements are now retained as
  searchable observations when Windows UIA has no distinct node.

### Changed

- Text commands verify UIA `SetValue` and fall back to Unicode-safe clipboard paste. A text
  command is successful only when an accessible value or exact clipboard readback confirms it.
- Disabled controls retain their potential commands and learned enabling conditions; runtime
  execution refuses them until UIA reports that the condition is satisfied.
- Scanned LOCATION records omit volatile names when AutomationId is available, while retaining
  XPath as a duplicate-ID and hierarchy fallback. Existing control IDs are reused when the
  AutomationId is unique on a learned page.
- The operation Skill batches knowledge searches, performs at most one direct runtime-context
  check, accepts a precise user send request as one-operation confirmation, and never switches
  to learning or retries an uncertain external action automatically.

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
