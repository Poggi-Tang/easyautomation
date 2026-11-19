"""Conservative identity checks for literal, named native UI actions.

These aliases describe common UI actions, not application-specific selectors.
Editable values and dynamic labels are deliberately not used as action identity.
"""

from __future__ import annotations

import re
import unicodedata


def normalize_label(value: str) -> str:
    value = unicodedata.normalize("NFKC", str(value or "")).split("\t", 1)[0]
    value = re.sub(r"\([&]?[a-z0-9]\)", "", value, flags=re.I)
    value = value.replace("&", "").replace("…", "").replace("...", "")
    return re.sub(r"[\s_\-/]+", "", value).casefold().strip(".:：")


# identity: (canonical intent, risk, literal translations / common intent aliases)
_ACTIONS = {
    "file-menu": ("open-file-menu", "safe", ["文件", "File", "File menu"]),
    "edit-menu": ("open-edit-menu", "safe", ["编辑", "Edit", "Edit menu"]),
    "format-menu": ("open-format-menu", "safe", ["格式", "Format", "Format menu"]),
    "view-menu": ("open-view-menu", "safe", ["查看", "View", "View menu"]),
    "help-menu": ("open-help-menu", "safe", ["帮助", "Help", "Help menu"]),
    "new-file": (
        "new-file",
        "state-changing",
        ["新建", "新建文件", "新建文档", "New", "New document"],
    ),
    "new-window": ("new-window", "state-changing", ["新窗口", "新建窗口", "New window"]),
    "open-file": ("open-file", "state-changing", ["打开", "Open", "Open document"]),
    "save": ("save-file", "state-changing", ["保存", "Save", "save-document"]),
    "save-as": ("save-file-as", "state-changing", ["另存为", "Save as", "save-document-as"]),
    "page-setup": ("page-setup", "state-changing", ["页面设置", "Page setup"]),
    "print": ("print-document", "external", ["打印", "Print"]),
    "exit": ("exit-application", "destructive", ["退出", "Exit"]),
    "undo": ("undo", "state-changing", ["撤销", "撤消", "Undo"]),
    "cut": ("cut", "state-changing", ["剪切", "Cut"]),
    "copy": ("copy", "safe", ["复制", "Copy"]),
    "paste": ("paste", "state-changing", ["粘贴", "Paste"]),
    "delete": ("delete", "destructive", ["删除", "Delete"]),
    "find": ("find", "safe", ["查找", "Find", "find-text"]),
    "find-next": ("find-next", "safe", ["查找下一个", "Find next"]),
    "find-previous": ("find-previous", "safe", ["查找上一个", "Find previous"]),
    "replace": (
        "replace",
        "state-changing",
        ["替换", "Replace", "replace-text", "replace-current-match"],
    ),
    "replace-all": (
        "replace-all",
        "state-changing",
        ["全部替换", "Replace all", "replace-all-matches"],
    ),
    "goto": ("goto", "safe", ["转到", "Go to", "go-to-line", "goto-line"]),
    "select-all": ("select-all", "safe", ["全选", "Select all"]),
    "date-time": (
        "insert-date-time",
        "state-changing",
        ["时间/日期", "日期/时间", "Time/Date", "Date/Time", "insert-time-date"],
    ),
    "search-bing": ("search-bing", "external", ["使用 Bing 搜索", "Search with Bing"]),
    "word-wrap": ("toggle-word-wrap", "state-changing", ["自动换行", "Word wrap"]),
    "font": ("open-font-dialog", "safe", ["字体", "Font"]),
    "status-bar": ("toggle-status-bar", "state-changing", ["状态栏", "Status bar"]),
    "zoom": ("open-zoom-menu", "safe", ["缩放", "Zoom"]),
    "zoom-in": ("zoom-in", "state-changing", ["放大", "Zoom in"]),
    "zoom-out": ("zoom-out", "state-changing", ["缩小", "Zoom out"]),
    "zoom-reset": (
        "restore-default-zoom",
        "state-changing",
        ["恢复默认缩放", "Restore default zoom"],
    ),
    "about": ("open-about-dialog", "safe", ["关于", "About"]),
    "close-window": ("close-window", "destructive", ["关闭", "关闭窗口", "Close", "Close window"]),
    "maximize-window": (
        "maximize-window",
        "state-changing",
        ["最大化", "Maximize", "Maximize window"],
    ),
    "minimize-window": (
        "minimize-window",
        "state-changing",
        ["最小化", "Minimize", "Minimize window"],
    ),
    "restore-window": ("restore-window", "state-changing", ["还原", "Restore", "Restore window"]),
    "cancel": ("cancel", "safe", ["取消", "Cancel"]),
    "match-case": (
        "toggle-match-case",
        "state-changing",
        ["区分大小写", "Match case", "Case sensitive"],
    ),
    "wrap-around": (
        "toggle-wrap-around",
        "state-changing",
        ["循环", "循环查找", "Wrap around", "Wrap search"],
    ),
    "search-up": ("select-search-up", "state-changing", ["向上", "Up", "Search up"]),
    "search-down": ("select-search-down", "state-changing", ["向下", "Down", "Search down"]),
    "find-input": (
        "enter-search-text",
        "state-changing",
        [
            "查找内容",
            "Find what",
            "Search content input",
            "Find what text input",
            "Find what input",
        ],
    ),
    "replace-input": (
        "enter-replacement-text",
        "state-changing",
        ["替换为", "Replace with", "Replacement text input", "Replacement input"],
    ),
}
_ALIASES = {
    normalize_label(alias): identity
    for identity, (intent, _risk, labels) in _ACTIONS.items()
    for alias in [identity, intent, *labels]
}
NATIVE_TYPES = {"MenuItemControl", "ButtonControl", "CheckBoxControl", "RadioButtonControl"}
_INPUT_LABELS = {
    normalize_label(label): key
    for key, labels in {
        "find-input": ["查找内容", "Find what"],
        "replace-input": ["替换为", "Replace with"],
    }.items()
    for label in labels
}


# These are user-facing function descriptions. Locator provenance and effect
# verification belong in separate metadata fields, never in the description.
_ACTION_DESCRIPTIONS = {
    "file-menu": "打开文件和文档相关的命令菜单。",
    "edit-menu": "打开编辑相关的命令菜单。",
    "format-menu": "打开格式相关的命令菜单。",
    "view-menu": "打开视图相关的命令菜单。",
    "help-menu": "打开帮助和软件信息相关的命令菜单。",
    "new-file": "新建文件或文档。",
    "new-window": "打开一个新的软件窗口。",
    "open-file": "打开已有的文件或文档。",
    "save": "保存当前文件、文档或工程。",
    "save-as": "以新的名称或位置保存当前文件、文档或工程。",
    "page-setup": "设置页面或打印相关选项。",
    "print": "打印当前文件或文档。",
    "exit": "退出当前软件。",
    "undo": "撤销最近一次编辑操作。",
    "cut": "剪切当前选中的内容。",
    "copy": "复制当前选中的内容。",
    "paste": "将剪贴板内容粘贴到当前位置。",
    "delete": "删除当前选中的内容或对象。",
    "find": "查找指定的文字或内容。",
    "find-next": "查找下一个匹配项。",
    "find-previous": "查找上一个匹配项。",
    "replace": "替换当前匹配的文字或内容。",
    "replace-all": "替换全部匹配的文字或内容。",
    "goto": "跳转到指定的位置、行或对象。",
    "select-all": "选中当前范围内的全部内容。",
    "date-time": "在当前位置插入日期和时间。",
    "search-bing": "使用 Bing 搜索当前选中的内容。",
    "word-wrap": "切换文本的自动换行显示。",
    "font": "打开字体设置对话框。",
    "status-bar": "显示或隐藏状态栏。",
    "zoom": "打开缩放选项菜单。",
    "zoom-in": "放大当前内容的显示比例。",
    "zoom-out": "缩小当前内容的显示比例。",
    "zoom-reset": "恢复默认的显示比例。",
    "about": "打开关于窗口，查看软件信息。",
    "close-window": "关闭当前窗口。",
    "maximize-window": "将当前窗口最大化。",
    "minimize-window": "将当前窗口最小化。",
    "restore-window": "将当前窗口恢复为普通大小。",
    "cancel": "取消当前对话框或操作。",
    "match-case": "切换查找或替换时是否区分大小写。",
    "wrap-around": "切换查找到末尾后是否从开头继续。",
    "search-up": "将查找方向设为向上。",
    "search-down": "将查找方向设为向下。",
    "find-input": "输入要查找的文字或内容。",
    "replace-input": "输入用于替换的文字或内容。",
}


def generated_description(value: object, limit: int = 0) -> str:
    """Discard generated prose without Chinese text; keep exact UI terms within it."""
    text = value.strip() if isinstance(value, str) else ""
    text = text[:limit] if limit else text
    return text if re.search(r"[\u3400-\u9fff]", text) else ""


def native_action_identity(native: dict) -> str:
    if native.get("control_type") == "EditControl":
        return _INPUT_LABELS.get(normalize_label(native.get("name", "")), "")
    if native.get("control_type") in NATIVE_TYPES:
        return identity(native.get("name", ""))
    return ""


def identity(value: str) -> str:
    label = normalize_label(value)
    if label.startswith(("关于", "about")):
        return "about"
    return _ALIASES.get(label, "")


def native_metadata(control) -> dict:
    return {
        "name": str(getattr(control, "Name", "") or ""),
        "control_type": str(getattr(control, "ControlTypeName", "") or ""),
        "automation_id": str(getattr(control, "AutomationId", "") or ""),
        "accelerator_key": str(getattr(control, "AcceleratorKey", "") or ""),
        "access_key": str(getattr(control, "AccessKey", "") or ""),
    }


def check(semantic: dict, native: dict) -> dict:
    label = native.get("name", "")
    actual = native_action_identity(native)
    predictions = {key: identity(semantic.get(key, "")) for key in ("intent", "semantic_name")}
    if not semantic.get("dynamic_context") and semantic.get("entity_type") == "text-label":
        predictions["observed_value"] = identity(semantic.get("observed_value", ""))
    conflicts = {
        key: value for key, value in predictions.items() if actual and value and value != actual
    }
    unresolved = (
        actual
        and (semantic.get("intent") or semantic.get("semantic_name"))
        and not any(predictions.get(key) for key in ("intent", "semantic_name"))
    )
    return {
        "status": (
            "conflict"
            if conflicts
            else "unresolved"
            if unresolved
            else "consistent"
            if actual
            else "not-comparable"
        ),
        "native_label": label,
        "native_identity": actual,
        "predicted_identities": predictions,
        "conflicts": conflicts,
    }


def literal_semantics(native: dict) -> dict | None:
    """Return user-facing function semantics for a recognized native action."""
    key = native_action_identity(native)
    if not key:
        return None
    intent, risk, _labels = _ACTIONS[key]
    control_type = native.get("control_type")
    return dict(
        semantic_name=normalize_display(native["name"]),
        intent=intent,
        description=_ACTION_DESCRIPTIONS[key],
        role="action",
        actions=["click", "set-text"] if control_type == "EditControl" else ["click"],
        aliases=[],
        risk=risk,
        requires_confirmation=risk in {"external", "destructive"},
        confidence=1.0,
        source="uia-native-label",
        description_source="uia-native-label",
        description_status="inferred",
        entity_type="input" if control_type == "EditControl" else "ui-element",
        observed_value=native["name"],
        dynamic_context=False,
        evidence=[f"UIA Name: {native['name']}"],
        ambiguity="",
    )


def normalize_display(value: str) -> str:
    value = unicodedata.normalize("NFKC", value).split("\t", 1)[0]
    return re.sub(r"\([&]?[a-z0-9]\)", "", value, flags=re.I).replace("&", "").strip()


def assert_consistent(record: dict, control=None) -> None:
    stored = check(record, record)
    if record.get("semantic_binding", {}).get("status") in {"conflict", "unresolved"} or stored[
        "status"
    ] in {"conflict", "unresolved"}:
        raise RuntimeError(f"Semantic identity conflict: {stored}")
    if control is None:
        return
    current = native_metadata(control)
    live = check(record, current)
    if live["status"] in {"conflict", "unresolved"}:
        raise RuntimeError(f"Live semantic identity conflict: {live}")
    if (
        record.get("control_type") in NATIVE_TYPES or native_action_identity(record)
    ) and not record.get("dynamic_context"):
        expected = normalize_label(record.get("name", ""))
        actual = normalize_label(current["name"])
        if expected and actual and expected != actual:
            raise RuntimeError(
                f"Native action label changed: {record.get('name')!r} -> {current['name']!r}"
            )
