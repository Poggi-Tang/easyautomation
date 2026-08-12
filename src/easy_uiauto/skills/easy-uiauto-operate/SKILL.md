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
5. Call `run_ui_command(app_id=..., command=..., text=...)`.
   Pass `confirm=true` only after the user explicitly approves an external or destructive effect.
6. For `set-text`, pass the user text through the `text` argument. Never place it inside the command name.
7. After navigation changes the page, query commands again before the next operation.
8. If execution quarantines a control, stop using that command. Ask the user to expose the correct page and invoke the learning workflow to rescan it.

## Safety Rules

- Execute only commands returned by `list_ui_commands`; those records passed scan-time verification.
- Do not bypass quarantine with raw coordinates unless the user explicitly requests a one-off operation.
- Do not assume a visual match is correct when identical icons exist. Require the stored LOCATION and image evidence.
- For destructive or externally visible operations, inspect the semantic command and current page before executing.

## Fallback Order

1. Verified semantic UI command.
2. Fresh UIA inspection and a learning rescan.
3. OCR or image matching for one-off recovery.
4. Remote AI vision only when deterministic methods cannot identify the target.
