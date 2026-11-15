# Models, offline use and API keys

English | [简体中文](model-and-privacy.zh-CN.md)

## Three ways to run

- **Local features only**: leave the API settings empty. Manage the vault, add controls manually, scan with local YOLO / RapidOCR, and compose or execute saved workflows.
- **Local features with a cloud multimodal model**: configure a multimodal model with image input and JSON output through an OpenAI-compatible API for semantic descriptions, reference interpretation and deep learning. Pricing and availability depend on the provider.
- **Local features with an intranet multimodal model**: configure a local or intranet multimodal model service with the same capabilities. The offline bundle does not include this service; deploy it separately.

The first source installation through `setup.bat` downloads dependencies and models. Complete GUI
and offline deployment ZIPs from Releases include the runtimes, detection weights and OCR models.
Their local features run offline after extraction and installation. The target application,
remote reference pages and cloud model services may still need network access.

## Data used during learning

| Operation | Local processing | Data sent to the configured model endpoint |
| --- | --- | --- |
| GUI basic learning without semantic descriptions | Screenshots, YOLO, OCR, UIA matching, vault writes | None |
| Added semantic descriptions | Local scanning and matching | Window images, control properties and region information; reference-assisted descriptions also include relevant excerpts |
| Reference import | File reading, text extraction, progress caching | Reference text chunks, source information and application context for indexing and citation checks |
| Deep learning | UI interaction, observation, saving regions/controls/workflows | Current images, exploration progress, user instructions, relevant references and test inputs |
| CLI / MCP visual-first scan | UIA matching, locator checks, knowledge storage | Selected-window images or image tiles |
| Effect interpretation | Before/after observation and change detection | Target-window images, related popup crops and control changes |
| Explicit remote vision lookup | Validation of returned coordinates | Requested screenshot area and target description |
| Fetch model list | Update the model selector | Authenticated request; no screenshots or references |

GUI reference import currently includes both text extraction and model interpretation. Installing
OCR alone does not complete that workflow. Basic learning does not click the target application,
but still calls the API when semantic descriptions are enabled. YOLO detects elements and
RapidOCR reads text; neither independently understands business tasks or plans interactions.

Vault Markdown, images and reference sources are ordinary local files without automatic encryption.
Version control requires your own Git setup. Team access management, workflow version rollback
and multi-machine scheduling are not built-in services.

## API key storage and upgrades

Starting with 1.1.0, the Qt GUI and standalone configuration EXE mask API keys by default and
provide a reveal button. Saving from either GUI or the CLI stores the key in the current
Windows user's Credential Manager under `easy-uiauto/vision-api-key`. Paths, API URL and model
name remain in user environment variables. Credential storage protects the saved value;
it does not replace access controls at the model service.

Read priority is Windows Credential Manager, then user environment variables, process environment
variables, and the old `mcp_servers.easy_uiauto.env` entry. Only this application's settings are
read. A saved credential takes precedence over an old environment variable; update it in Settings.

Existing configuration remains readable. A successful save removes the old
`EASY_UIAUTO_VISION_API_KEY` user environment variable and its current-process copy. Failed saves
restore the previous settings and never fall back to plaintext storage. A normal GUI save does
not rewrite existing client configuration files or erase historical plaintext copies. Re-registering
the Codex MCP server migrates this application's settings and replaces its old entry.

Upgrade the GUI, MCP and Python together to 1.1.0 before saving settings: older versions cannot
read Credential Manager. Updated processes under the same user read changes dynamically.
Another user or machine needs its own key configuration. Restart older clients after upgrading
to discard inherited environment values.

Keys are not written into the vault or distribution bundles. Model request errors report status
codes or fixed messages without provider error bodies, which can echo invalid keys. This does
not control request logging performed by custom proxies, third-party SDKs or user scripts.
