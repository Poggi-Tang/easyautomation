"""Current locator evidence, independent of semantic approval and quarantine."""
from __future__ import annotations

import hashlib
import json

from .locator_images import record_location


def signature(record):
    context = {"location": record_location(record),
               "window_scope": record.get("window_scope") or {"kind": "main"},
               "image": record.get("image") or "", "image_variants": record.get("image_variants") or [],
               "image_sha256": record.get("image_sha256") or ""}
    return hashlib.sha256(json.dumps(context, sort_keys=True, ensure_ascii=False).encode("utf-8")).hexdigest()


def current(record):
    check = record.get("locator_verification", {})
    if not check:
        return {}
    try:
        return check if check.get("locator_signature") == signature(record) else {}
    except (TypeError, ValueError):
        return {}


def apply(record, result, *, source):
    """Only record evidence for the locator that was actually checked."""
    if result.get("locator_signature") != signature(record):
        return False
    if result.get("ok") is True and result.get("resolved_by") not in {"image", "location"}:
        return False
    from .mcp.knowledge import utc_now
    record["locator_verification"] = {
        **{key: result[key] for key in ("locator_signature", "resolved_by", "score", "uniqueness", "template", "error")
           if key in result},
        "status": "passed" if result.get("ok") is True else "failed",
        "source": source, "verified_at": utc_now(),
    }
    return True
