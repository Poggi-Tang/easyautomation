---
name: easy-uiauto-learning
description: Scan and explore Windows applications into an easy_uiauto UI knowledge vault, or analyze a supplied UI image into numbered annotations and element coordinates. Use for application learning, interface mapping, UI CLI generation, or UI screenshot annotation.
---

# easy_uiauto Learning

Build reusable UI knowledge through the `easy_uiauto` MCP tools.

## Supplied UI Images

For standalone local YOLO + OCR exports from CMD, use the checkout's
`annotate-image.cmd "input-image" "output-parent"` or installed
`easy_uiauto_image "input-image" "output-parent"`. This creates an independent run
folder with `text.png`, `controls.png`, `combined.png`, coordinates, raw detections,
logs and `index.html`. Default profile is `block-diagram`; use `--profile circuit`
for electrical symbols or `--profile ui` for general UI proposals. Inference is
local CPU with cached models; it does not use the configured remote vision API.
The CMD pipeline automatically checks missed small text inside closed frames,
using cached detection proposals and agreeing original/enlarged recognition.
Review `ocr-recovery.json` and source `rapidocr-small-text-recovery`; recovered
observations are appended while baseline text and display IDs remain stable.
`--no-small-text-recovery` disables that pass for comparison.

For an attached or saved screenshot, call `analyze_ui_image(image_path=..., output_dir=...)`
with its local file path on the MCP server's machine. This path uses the configured vision
endpoint directly; window enumeration and live-desktop readiness checks are unnecessary.
Return `annotated_image`, `coordinates_json`, and the matching numbered `elements`.
Coordinates are pixels relative to the EXIF-oriented image's top-left corner, with
exclusive right/bottom bounds. Honor `truncated`; all detections remain observations.
Saved-image graphics use blue boxes and `G1…` labels, text uses orange boxes and
`T1…` labels. Keep each label two pixels outside its own box, sliding only along
the four edges for collision avoidance. `annotation_label` maps the display label
to the original numeric `id` used by correction tools. Preserve labels across layers.
If image-edge labels need canvas margins, `annotation_layout.image_origin` records
the input-image offset; element coordinates remain relative to the original input.
Small images are enlarged automatically; returned coordinates still refer to the original
displayed image. Use `analysis_scale=1` only when enlargement is unwanted.
Display `annotated_preview_image` when present; its strokes are drawn at one pixel
after enlargement. `rect`/`bbox` include `box_padding` (default two original pixels)
and known attached components; `detected_rect` preserves the raw prediction. Use
`analysis_context` for the user's requested scope and component granularity.
Inferred names, functions, and suggested actions have no UIA or execution verification.
This operation does not update the application vault or generate commands.
Visible text is also a control: include independent captions, titles, numeric
values and pin/port labels unless the user explicitly asks for graphics only.
Use a separate ID and `entity_type=text-label`, with literal content in
`observed_value`; default to one independent label or text line per box. Associate
it with its graphic using `text-of`/`label-of`, not `attached-to`, so an external
caption does not enlarge the graphic's bounds. Text and graphic boxes can overlap
without being duplicates. Circuit-outline fitting does not validate text bounds;
its `text_control_ids` remain observations needing a text detector/OCR or review.

To repair existing boxes, call `correct_ui_image_annotations(coordinates_json=...,`
`refine_ids=[...])` for local pixel fitting of selected isolated icons on light,
plain backgrounds. This uses the saved original image and makes no API request.
For exact corrections, pass `corrections=[{"id":3,"rect":{"left":110,"top":65,`
`"right":125,"bottom":81}}]` instead; right/bottom are exclusive original-image
coordinates and receive no extra padding. Preserve IDs and semantic content.
Return the new PNG/JSON pair and review `correction.updates` for skipped targets.
For connected circuit diagrams, set `profile="circuit"` and pass the original
model output as `coordinates_json`. It detects rectangular frames, paired circles,
grounded load branches, and thick busbars in the original pixels. Selected
observations sharing one geometry candidate are merged under their primary ID;
merged IDs and conflicting captions stay in metadata. Set `discover_missing=true`
only when geometric busbar additions are wanted. Other missing component kinds
are not added. Inspect candidate counts, merged/added IDs, skipped targets and the
new annotation. A correction test must not substitute previous reviewed geometry
for the original model output. Originals and `detected_rect` remain available;
correction does not establish UIA or functional verification. Unsupported symbols
or ambiguous backgrounds may still require explicit bounds.
The circuit profile also fits connected control block diagrams using complete
frames, pointed tags, machine/meter ellipses, terminal rings and grounds. Review
`diagram_family`, `needs_review_ids` and each skipped update: unsupported model
boxes remain unchanged. Do not count them as successful corrections. Coverage
includes tile bounds and raw/merged observation counts; conflicting tile evidence
is retained in `tile_observations`. Model confidence is not a calibrated geometry
score, so a high score alone does not establish accurate bounds.

## Live Application Workflow

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
   Also inspect key records for `entity_type`, `observed_value`, `dynamic_context`,
   `current_state`, `state_reason`, `enabling_condition`, `relationships`, and
   `visual_components`. Basic learning must understand visible UI meaning and state; these
   fields do not require interactive deep exploration.
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
- Keep stable function separate from volatile values. Contact names, document names, selected
  rows, and current results belong in `observed_value` with `dynamic_context=true`; generic
  inputs and actions must not be renamed for whichever item happened to be visible during scan.
- Retain meaningful visual-only entities such as avatars, names, message areas, badges, and
  status fields even when Windows UIA exposes no distinct node. They remain observations and
  never become executable commands without a verified locator.
- When UIA exposes one coarse Pane, Document, or Custom control for several internally drawn
  elements, pass that control result or its `bounds` once to `detect_visual_elements`. Use the
  returned rectangles and containment tree as local geometry candidates. This tool is local,
  class-agnostic, and non-semantic; do not run it over native child controls already exposed by
  UIA, and do not mistake a rectangle candidate for a verified action.
- Use its `relations`, `groups`, and `layout_signature` for self-drawn composite structures.
  Treat every group's `size_policy=variable` literally: current width, height, and text content
  are observations, not identity. Match repeated structures by relation topology and parent
  anchor, and keep separate rows or vertically repeated items as separate instances.
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
- Disabled actions include a grounded state reason and enabling condition when visible context
  supports one, such as a Send button disabled because the message draft is empty.
- Explicitly requested command effects have a before/after record and an evidence-based success condition.
- Quarantined controls and their reasons are reported.
- The vault path, source page screenshot, and numbered annotated page screenshot are returned.
