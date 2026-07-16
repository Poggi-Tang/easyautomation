from __future__ import annotations

import os
import uuid

import pytest
import uiautomation

from easy_uiauto.utils import find_control_by_xpath, get_control_xpath

pytestmark = [
    pytest.mark.windows_integration,
    pytest.mark.skipif(
        os.environ.get("EASY_UIAUTO_RUN_SIMUNPS_TESTS") != "1",
        reason="set EASY_UIAUTO_RUN_SIMUNPS_TESTS=1 for reversible SimuNPS smoke tests",
    ),
]


def test_simunps_search_toggle_log_and_xpath_are_reversible():
    window = uiautomation.WindowControl(
        Name="SimuNPS",
        ClassName="EmtMainUI",
        AutomationId="SimuNPS",
        searchDepth=1,
    )
    if not window.Exists(2):
        pytest.skip("SimuNPS is not running")

    search = window.EditControl(
        Name="model_serach",
        ClassName="SearchLineEdit",
        searchDepth=15,
    )
    message_filter = window.CheckBoxControl(
        Name="\u6d88\u606f",
        ClassName="QPushButton",
        searchDepth=15,
    )
    log_table = window.TableControl(
        Name="table_log",
        ClassName="LogTableView",
        searchDepth=15,
    )
    assert search.Exists(2)
    assert message_filter.Exists(2)
    assert log_table.Exists(2)

    value_pattern = search.GetPattern(uiautomation.PatternId.ValuePattern)
    toggle_pattern = message_filter.GetPattern(uiautomation.PatternId.TogglePattern)
    grid_pattern = log_table.GetPattern(uiautomation.PatternId.GridPattern)
    assert value_pattern is not None
    assert toggle_pattern is not None
    assert grid_pattern is not None

    original_value = value_pattern.Value
    original_toggle = int(toggle_pattern.ToggleState)
    probe = f"easy-uiauto-smoke-{uuid.uuid4().hex[:8]}"
    try:
        assert value_pattern.SetValue(probe, waitTime=0)
        assert value_pattern.Value == probe
        assert toggle_pattern.Toggle(waitTime=0)
        assert int(toggle_pattern.ToggleState) != original_toggle

        xpath = get_control_xpath(search)
        assert xpath
        replayed = find_control_by_xpath(xpath, use_cache=False)
        assert replayed is not None
        assert uiautomation.ControlsAreSame(search, replayed)
        assert int(grid_pattern.RowCount) >= 0
    finally:
        value_pattern.SetValue(original_value, waitTime=0)
        if int(toggle_pattern.ToggleState) != original_toggle:
            toggle_pattern.Toggle(waitTime=0)

    assert value_pattern.Value == original_value
    assert int(toggle_pattern.ToggleState) == original_toggle
