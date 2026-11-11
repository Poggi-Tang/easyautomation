"""Reference images and explicit visual targets shared by GUI replay and preview."""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


def is_image_location(location):
    return (not location.get("Xpath") and
            isinstance(location.get("Img"),str) and bool(location["Img"].strip()) and
            ((location.get("PARAMETERS") or {}).get("locator_kind")=="visual-observation" or
             not any(location.get(k) for k in ('Name','ClassName','ControlType','AutomationId','foundIndex'))))


def record_location(record):
    """Adapt image-only observations without manufacturing or persisting UIA identity."""
    from .mcp.protocol import normalize_location
    location=normalize_location(record.get('location') or {})
    image=location.get('Img') or record.get('image')
    if (not location['Xpath'] and isinstance(image,str) and image.strip() and
            not any(location.get(k) for k in ('Name','ClassName','ControlType','AutomationId','foundIndex'))):
        location['Img']=image
        location['PARAMETERS'].setdefault('locator_kind','visual-observation')
    return location


def reference_images(directory, record):
    """An explicit LOCATION.Img overrides learned templates; relative paths stay local."""
    directory = Path(directory).resolve()
    explicit = (record.get("location") or {}).get("Img")
    values = [explicit] if explicit else [record.get("image"), *(record.get("image_variants") or [])]
    result = []
    for value in values:
        if not isinstance(value, str) or not value.strip():
            continue
        path = Path(value)
        try:
            resolved = (directory / path).resolve()
            if not path.is_absolute() and not resolved.is_relative_to(directory):
                continue
            if resolved not in result and resolved.is_file():
                result.append(resolved)
        except (OSError, ValueError):
            continue
    return result


@dataclass
class ImageTarget:
    rectangle: dict
    reference: Path
    score: float
    uniqueness: float
    name: str = ""


def target_arguments(target):
    return {"rectangle": target.rectangle} if isinstance(target, ImageTarget) else {"control": target}


def target_evidence(target):
    if isinstance(target, ImageTarget):
        return dict(resolved_by="image", rect=target.rectangle, template=str(target.reference),
                    score=target.score, uniqueness=target.uniqueness)
    return dict(resolved_by="location")
