import math

import pytest

from easy_uiauto.mcp.manual_layout import assign_region, segmentation, validate_regions


def region(identifier="top", name="工具栏", bounds=(0, 0, 1, .5)):
    return {"id": identifier, "name": name,
            "normalized_rect": dict(zip(("left", "top", "right", "bottom"), bounds))}


@pytest.mark.parametrize("bounds", [(-.1, 0, 1, 1), (0, 0, 1.1, 1), (0, .5, 1, .5),
                                    (0, 0, math.nan, 1), (0, False, 1, 1)])
def test_invalid_regions_are_rejected(bounds):
    with pytest.raises(ValueError): validate_regions([region(bounds=bounds)])


def test_overlaps_and_duplicate_ids_fail_but_touching_edges_are_valid():
    with pytest.raises(ValueError, match="重叠"):
        validate_regions([region(), region("other", bounds=(.2, .3, 1, 1))])
    with pytest.raises(ValueError, match="不重复"):
        validate_regions([region(), region(bounds=(0, .5, 1, 1))])
    assert len(validate_regions([region(), region("bottom", bounds=(0, .5, 1, 1))])) == 2


@pytest.mark.parametrize("bounds,expected", [((.1, .1, .3, .2), ("top", "contained")),
                                             ((.1, .4, .3, .8), ("bottom", "largest-overlap")),
                                             ((.1, .4, .3, .6), ("unassigned", "cross-region-ambiguous"))])
def test_control_region_assignment_is_deterministic(bounds, expected):
    regions = [region(), region("bottom", bounds=(0, .5, 1, 1))]
    assert assign_region(region(bounds=bounds)["normalized_rect"], regions) == expected


def test_uncovered_controls_are_not_silently_discarded():
    assert assign_region(region(bounds=(0, .8, 1, 1))["normalized_rect"], [region()]) == ("unassigned", "outside-defined-regions")


def test_region_coordinates_are_window_relative_and_size_changes_require_review():
    layout = {"regions": [region()], "image_size": {"width": 400, "height": 300}}
    result = segmentation(layout, {"left": -300, "top": 250, "width": 400, "height": 300})
    assert result["regions"][0]["rect"] == dict(left=0, top=0, right=400, bottom=150, width=400, height=150)
    with pytest.raises(ValueError, match="尺寸"):
        segmentation(layout, {"width": 800, "height": 600})
