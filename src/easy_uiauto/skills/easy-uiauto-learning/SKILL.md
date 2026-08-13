---
name: easy-uiauto-learning
description: Visually scan and safely explore any visible Windows desktop application into an Obsidian-compatible easy_uiauto knowledge vault with key controls, stable UIA locations, control images, semantic commands, and learned operation effects. Use when the user asks to learn, scan, deeply explore, map, catalog, or repair an application's interface or generate a UI CLI.
---

# easy_uiauto Learning

Build reusable UI knowledge through the `easy_uiauto` MCP tools.

## Workflow

1. Call `get_ui_learning_readiness` before listing or scanning windows. If `ready` is false,
   stop immediately and return its setup instruction. Do not attempt either scan strategy.
2. Call `list_windows` once and identify the exact target window. Do not scan unrelated windows.
3. Ask the user to navigate to the page that should be learned when the requested scope is unclear.
4. Call `scan_window_knowledge(window_name=..., strategy="visual-first", show_overlay=true)`
   exactly once. The scan draws one click-through overlay for all understood controls and
   saves the matching numbered `annotated_page_image` in the vault. Surface its MCP progress
   messages so the user sees the task ID, stage, elapsed time, and verification count.
   Never retry with `full-uia` after configuration, authentication, timeout, network, or model
   errors: both strategies require the same remote endpoint. Use `full-uia` only when the user
   explicitly requests diagnostic UIA coverage after a successful readiness check.
5. If the tool wait times out or loses its final response, call `get_ui_learning_status()`.
   When the state is `running`, report the current stage and elapsed time and continue checking
   the same task; never start a duplicate scan. Treat only `state="failed"` as a scan failure.
   When the state is `completed`, use its result summary instead of rescanning.
6. Inspect `strategy`, `visual_targets`, `truncated`, `uia_controls_seen`, `controls_saved_this_scan`,
   `knowledge_controls_total`, `commands`, `status_counts`, `semantic_counts`,
   `annotated_controls`, `annotated_page_image`, `command_items`, `quarantine_summary`,
   and `stage_timings`. These fields are sufficient to finish a basic learning request.
7. Treat the scan as incomplete when `truncated` is true. Rescan with a larger limit or split the work across pages.
8. Report the returned `command_items` and quarantine summary. Do not call
   `list_ui_commands` or `search_ui_knowledge` again after a successful basic scan unless
   the user asks for a separate full listing, search, or repair operation.
9. Search by user intent and aliases, not only by the original UIA name.
10. For an explicit repair request, call
   `search_ui_knowledge(app_id=..., include_quarantine=true)` and inspect
   `semantic_ambiguity`, evidence, and confidence.
11. Use `teach_ui_control` only when the user or direct observation establishes the real
   function. Teaching semantics never overrides failed LOCATION or image verification.
12. Rescan the same page after the UI becomes stable. Manually taught meanings survive rescans.
13. Call `learn_ui_command_effect` only when the user explicitly requests command-effect
   learning. Use `recover=true` only when Escape is a valid reversal.
14. Call `explore_ui_workflows(policy="safe", max_depth=3)` only when the user explicitly
   requests deep learning, interactive exploration, or autonomous workflow learning. A plain
   "learn", "scan", or "map" request ends after the basic visual-first scan. Deep exploration
   tests known reversible commands and recursively learns newly opened pages or dialogs. Use
   `policy="supervised"` only with active user supervision. Stop if interference or recovery
   failure is reported. Never request automated scrolling or dragging; they are unsupported.
15. After explicitly requested deep exploration, call `list_ui_interactions` and report
   learned before/after states, changed regions, popup windows, effects, success conditions,
   unstable regions, and recovery results.
16. Repeat visual scanning for every important page or dialog that remains unknown.

## Storage Rules

- Treat Markdown/YAML and PNG files under `~/easy_uiauto_vault/applications/<app-id>` as the source of truth.
- Treat `.easy_uiauto/index.json` as disposable. Use `rebuild_ui_knowledge_index` after manual Markdown edits.
- Keep page, region, and semantic names generic and application-derived. Do not inject unrelated product knowledge.
- Keep locator, image, semantic, and function verification separate. An image match does not
  prove a control's purpose, and an AI inference does not prove the operation was executed.
- Never probe sending, publishing, deleting, purchasing, closing without saving, or other
  externally visible/destructive controls during unattended learning.
- Preserve quarantined records as diagnostics. Do not delete them merely to improve pass rates.
- Full desktop before/after images are local evidence. Only target-window images and related
  popup crops may be sent to the configured vision endpoint.
- Treat dynamic regions, transient windows, and user-interference records as uncertainty, not
  as deterministic application behavior.

## Completion Criteria

- No requested page remains unscanned.
- No scan is silently truncated.
- Verified commands are listed.
- Every actionable control has a high-confidence contextual meaning or a reported quarantine reason.
- Explicitly requested command effects have a before/after record and an evidence-based success condition.
- Quarantined controls and their reasons are reported.
- The vault path, source page screenshot, and numbered annotated page screenshot are returned.
