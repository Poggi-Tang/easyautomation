"""Phrase OCR regression: words on one UI label share the matching bounds."""

import pytest

from easy_uiauto.mcp.ocr_matching import find_match


def data():
    return dict(
        text=["MCP", "TARGET"],
        conf=[95, 90],
        left=[10, 50],
        top=[20, 20],
        width=[30, 60],
        height=[20, 20],
        page_num=[1, 1],
        block_num=[1, 1],
        par_num=[1, 1],
        line_num=[1, 1],
    )


@pytest.mark.parametrize("mode", ["contains", "exact"])
def test_multiword_label_matches_union_of_contiguous_words(mode):
    match = find_match(data(), "MCP TARGET", mode, 60, (-100, 200))
    assert match == dict(
        text="MCP TARGET",
        confidence=90,
        left=-90,
        top=220,
        width=100,
        height=20,
        center_x=-40,
        center_y=230,
    )


@pytest.mark.parametrize("change", ["line", "gap", "weak", "malformed", "missing-hierarchy"])
def test_unrelated_or_untrusted_words_are_not_combined(change):
    value = data()
    if change == "line":
        value["line_num"] = [1, 2]
    elif change == "gap":
        value["left"][1] = 300
    elif change == "weak":
        value["conf"][1] = 20
    elif change == "malformed":
        value["width"][0] = 0
    else:
        del value["line_num"]
    assert find_match(value, "MCP TARGET", "exact", 60) is None


def test_single_word_query_keeps_its_tight_box():
    match = find_match(data(), "TARGET", "contains", 60)
    assert match["text"] == "TARGET" and match["left"] == 50 and match["width"] == 60
