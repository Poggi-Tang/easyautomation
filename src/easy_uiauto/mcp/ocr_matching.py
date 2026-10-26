"""Match Tesseract words or contiguous phrases without joining separate UI rows."""

from __future__ import annotations


def find_match(data, text, match_mode, min_confidence, offset=(0, 0)):
    needle = " ".join(text.casefold().split())
    groups = {}
    words = []
    hierarchy = ("page_num", "block_num", "par_num", "line_num")
    for index, candidate in enumerate(data.get("text", [])):
        candidate = (candidate or "").strip()
        if not candidate:
            continue
        try:
            key = tuple(data[name][index] for name in hierarchy)
        except (KeyError, IndexError, TypeError):
            key = ("single", index)  # Missing hierarchy must not combine unrelated labels.
        try:
            word = dict(
                text=candidate,
                confidence=float(data["conf"][index]),
                left=int(data["left"][index]),
                top=int(data["top"][index]),
                width=int(data["width"][index]),
                height=int(data["height"][index]),
            )
            if not 0 <= word["confidence"] <= 100 or min(word["width"], word["height"]) <= 0:
                word = None
        except (KeyError, IndexError, TypeError, ValueError):
            word = None
        groups.setdefault(key, []).append(word)
        if word is not None:
            words.append(word)

    def match(span):
        if any(word["confidence"] < min_confidence for word in span):
            return None
        phrase = " ".join(word["text"] for word in span)
        haystack = " ".join(phrase.casefold().split())
        if not (haystack == needle if match_mode == "exact" else needle in haystack):
            return None
        left, top = min(word["left"] for word in span), min(word["top"] for word in span)
        width = max(word["left"] + word["width"] for word in span) - left
        height = max(word["top"] + word["height"] for word in span) - top
        left, top = left + offset[0], top + offset[1]
        return dict(
            text=phrase,
            confidence=min(word["confidence"] for word in span),
            left=left,
            top=top,
            width=width,
            height=height,
            center_x=left + width // 2,
            center_y=top + height // 2,
        )

    # Preserve tight single-word boxes instead of returning a larger containing phrase.
    for word in words:
        found = match([word])
        if found:
            return found
    for line in groups.values():
        for start, first in enumerate(line):
            if first is None:
                continue
            span = [first]
            for following in line[start + 1 : start + 64]:
                previous = span[-1]
                if following is None:
                    break
                gap = following["left"] - (previous["left"] + previous["width"])
                if gap < -max(previous["width"], following["width"]) or gap > max(
                    12, 2 * max(previous["height"], following["height"])
                ):
                    break
                span.append(following)
                found = match(span)
                if found:
                    return found
    return None
