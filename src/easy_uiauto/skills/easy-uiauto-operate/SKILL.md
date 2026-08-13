---
name: easy-uiauto-operate
description: Operate scanned Windows desktop applications through verified easy_uiauto semantic UI commands, search reusable control knowledge, and stop safely when knowledge is stale or quarantined. Use when the user asks to control or automate an application that has been scanned into the UI knowledge vault.
---

# easy_uiauto Operate

Prefer verified application knowledge over rediscovering the interface.

## Workflow

1. Call `list_ui_knowledge_apps` only when the application ID is not already established in
   the current task.
2. When the user asks what can be operated on the current page, call
   `show_ui_controls(app_id=..., include="executable")`. Report its numbered legend and do
   not run a separate UIA scan or remote vision request.
3. Otherwise make one targeted `search_ui_knowledge` call for one intent, or one
   `search_ui_knowledge_batch` call for a workflow with multiple intents. For example, message
   sending should batch-search the recipient/conversation, message input, and Send action.
   Do not also list all commands, and do not repeat broad searches merely to reconfirm fields
   already returned in the records.
4. Select a command whose intent, aliases, description, page, region, and action match the request.
5. Inspect `risk`, `requires_confirmation`, semantic confidence/source, and function-verification
   status. Do not claim inferred meanings were function-tested.
6. For two or more commands on the same stable page, call
   `run_ui_commands(app_id=..., steps=[...], highlight=true)` once. Use command strings for clicks and
   `{ "command": "...set-text", "text": "..." }` for text entry. The batch verifies
   all unique controls, highlights all batch targets in one overlay before the first action,
   and stops at the first action error.
7. Use `run_ui_command(app_id=..., command=..., text=..., highlight=true)` for one command
   or when a prior command navigates to another page. The red frame identifies the verified
   runtime target; it is click-through and does not remain in the input path.
   Pass `confirm=true` only after the user explicitly approves an external or destructive effect.
8. For `set-text`, pass the user text through the `text` field. Never place it inside the command name.
   Continue only when `action_verification.verified=true`; a tool return without verified text
   evidence is not successful input.
9. After navigation, reuse already returned generic input and action commands when they belong
   to the same learned page. If recipient or document context must be checked, call `find_control`
   once with the stored full LOCATION and inspect the live returned properties. Do not guess by
   name, dump the UIA tree, or launch multiple low-level discovery calls.
10. If execution quarantines a control or input verification fails, stop. Report the exact stale
   locator, precondition, or write evidence. Never start `scan_window_knowledge` from an operation
   request; ask the user before switching to a repair-learning workflow.
11. Use `list_ui_interactions` when an effect record exists and verify its success condition
    against the observed post-operation state. Do not infer success only because a click ran.

## Safety Rules

- Execute only commands returned by `list_ui_commands`; those records passed scan-time verification.
- Do not bypass quarantine with raw coordinates unless the user explicitly requests a one-off operation.
- Do not assume a visual match is correct when identical icons exist. Require the stored LOCATION and image evidence.
- For destructive or externally visible operations, inspect the semantic command and current page before executing.
- A direct request such as "send X to Y" is explicit confirmation for exactly that recipient,
  content, and one send attempt. Do not ask again, but do not broaden or repeat the action.
- Never automatically retry an external action when its outcome is uncertain.
- Never combine commands from different pages in one batch. Treat a page-changing command as the end of a batch.

## Fallback Order

1. Verified stored LOCATION.
2. A unique stored multi-state control image.
3. Unique local OCR text.
4. Remote AI vision only when deterministic methods fail and the user permits
   `allow_vision_fallback=true`.
5. User-approved repair learning when all stored evidence is stale or ambiguous. Stop the
   operation first; do not silently change modes or wait on a scan.
