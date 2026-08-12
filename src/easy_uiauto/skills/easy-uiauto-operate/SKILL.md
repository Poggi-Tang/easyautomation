---
name: easy-uiauto-operate
description: Operate scanned Windows desktop applications through verified easy_uiauto semantic UI commands, search reusable control knowledge, and stop safely when knowledge is stale or quarantined. Use when the user asks to control or automate an application that has been scanned into the UI knowledge vault.
---

# easy_uiauto Operate

Prefer verified application knowledge over rediscovering the interface.

## Workflow

1. Call `list_ui_knowledge_apps` to identify the application ID.
2. Call `list_ui_commands(app_id=...)` or `search_ui_knowledge(app_id=..., query=...)`.
3. Select a command whose intent, aliases, description, page, region, and action match the request.
4. Inspect `risk`, `requires_confirmation`, semantic confidence/source, and function-verification
   status. Do not claim inferred meanings were function-tested.
5. For two or more commands on the same stable page, call
   `run_ui_commands(app_id=..., steps=[...])` once. Use command strings for clicks and
   `{ "command": "...set-text", "text": "..." }` for text entry. The batch verifies
   all unique controls before the first action and stops at the first action error.
6. Use `run_ui_command(app_id=..., command=..., text=...)` for one command or when a
   prior command navigates to another page.
   Pass `confirm=true` only after the user explicitly approves an external or destructive effect.
7. For `set-text`, pass the user text through the `text` field. Never place it inside the command name.
8. After navigation changes the page, query commands again before the next operation.
9. If execution quarantines a control, stop using that command. Ask the user to expose the correct page and invoke the learning workflow to rescan it.
10. Use `list_ui_interactions` when an effect record exists and verify its success condition
    against the observed post-operation state. Do not infer success only because a click ran.

## Safety Rules

- Execute only commands returned by `list_ui_commands`; those records passed scan-time verification.
- Do not bypass quarantine with raw coordinates unless the user explicitly requests a one-off operation.
- Do not assume a visual match is correct when identical icons exist. Require the stored LOCATION and image evidence.
- For destructive or externally visible operations, inspect the semantic command and current page before executing.
- Never combine commands from different pages in one batch. Treat a page-changing command as the end of a batch.

## Fallback Order

1. Verified stored LOCATION.
2. A unique stored multi-state control image.
3. Unique local OCR text.
4. Remote AI vision only when deterministic methods fail and the user permits
   `allow_vision_fallback=true`.
5. Fresh visual-first learning scan when all stored evidence is stale or ambiguous.
