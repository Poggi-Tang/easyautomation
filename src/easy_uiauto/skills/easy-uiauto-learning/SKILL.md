---
name: easy-uiauto-learning
description: Visually scan and safely explore any visible Windows desktop application into an Obsidian-compatible easy_uiauto knowledge vault with key controls, stable UIA locations, control images, semantic commands, and learned operation effects. Use when the user asks to learn, scan, deeply explore, map, catalog, or repair an application's interface or generate a UI CLI.
---

# easy_uiauto Learning

Build reusable UI knowledge through the `easy_uiauto` MCP tools.

## Workflow

1. Call `list_windows` and identify the exact target window. Do not scan unrelated windows.
2. Ask the user to navigate to the page that should be learned when the requested scope is unclear.
3. Call `scan_window_knowledge(window_name=..., strategy="visual-first")`. Use
   `strategy="full-uia"` only when visual targeting fails or diagnostic coverage is required.
4. Inspect `strategy`, `visual_targets`, `truncated`, `uia_controls_seen`, `controls_saved_this_scan`,
   `knowledge_controls_total`, `commands`, `status_counts`, and `semantic_counts`.
5. Treat the scan as incomplete when `truncated` is true. Rescan with a larger limit or split the work across pages.
6. Call `list_ui_commands(app_id=...)` and report the generated command surface.
7. Search by user intent and aliases, not only by the original UIA name.
8. Call `search_ui_knowledge(app_id=..., include_quarantine=true)` when quarantine or
   uncertain semantics are nonzero. Inspect `semantic_ambiguity`, evidence, and confidence.
9. Use `teach_ui_control` only when the user or direct observation establishes the real
   function. Teaching semantics never overrides failed LOCATION or image verification.
10. Rescan the same page after the UI becomes stable. Manually taught meanings survive rescans.
11. Call `learn_ui_command_effect` for a user-requested command when its real response and
   success condition must be learned. Use `recover=true` only when Escape is a valid reversal.
12. Call `explore_ui_workflows(policy="safe", max_depth=3)` to rescan the current state,
   test known reversible commands, and recursively learn newly opened pages or dialogs. Use
   `policy="supervised"` only with active user supervision. Stop if interference or recovery
   failure is reported. Never request automated scrolling or dragging; they are unsupported.
13. Call `list_ui_interactions` and report learned before/after states, changed regions,
   popup windows, effects, success conditions, unstable regions, and recovery results.
14. Repeat visual scanning for every important page or dialog that remains unknown.

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
- Requested command effects have a before/after record and an evidence-based success condition.
- Quarantined controls and their reasons are reported.
- The vault path and page screenshot paths are returned to the user.
