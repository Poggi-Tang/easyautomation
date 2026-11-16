"""Compare UI state evidence without treating pixel dimensions as identity."""
from __future__ import annotations

import base64
import io

from PIL import Image, ImageChops, ImageFilter, ImageStat


def visual_evidence(image):
    preview = image.convert("RGB")
    preview.thumbnail((768, 768), Image.Resampling.LANCZOS)
    buffer = io.BytesIO()
    preview.save(buffer, format="PNG")
    return base64.b64encode(buffer.getvalue()).decode("ascii")


def same_image(first, second):
    """Tolerate scaling/rendering noise, but check local changes as well as the mean."""
    if min(*first.size, *second.size) < 2:return False
    if abs(first.width/first.height - second.width/second.height) > .01:return False
    # Use one common aspect-preserving grid, not each image's native pixel grid.
    width = min(first.width, second.width, 384)
    height = max(2, round(width*first.height/first.width))
    if height > 384:width=max(2,round(width*384/height));height=384
    size = (width,height)
    a, b = [im.convert("RGB").resize(size, Image.Resampling.LANCZOS).filter(ImageFilter.GaussianBlur(.4))
            for im in (first, second)]
    difference = ImageChops.difference(a, b)
    if max(ImageStat.Stat(difference).mean) > 1.25:return False
    for top in range(0, size[1], 16):
        for left in range(0, size[0], 16):
            tile = difference.crop((left, top, min(left+16, size[0]), min(top+16, size[1])))
            if max(ImageStat.Stat(tile).mean) > 5:return False
    return True


def matches(first, second):
    if not first or not second or first.get("version") != 1 or second.get("version") != 1:return False
    if any(first.get(k) != second.get(k) for k in ("scope", "native")):return False
    if not first.get("native"):return False
    try:
        with Image.open(io.BytesIO(base64.b64decode(first["image"]))) as a, \
                Image.open(io.BytesIO(base64.b64decode(second["image"]))) as b:
            return same_image(a, b)
    except (KeyError, ValueError, OSError):
        return False
