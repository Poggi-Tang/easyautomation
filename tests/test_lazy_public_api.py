"""Importing the package or manager must not initialize automation backends."""
import json
import subprocess
import sys


def test_manager_import_avoids_cv_recording_and_legacy_gui_in_a_fresh_process():
    code = '''
import json, sys
import easy_uiauto.library_launcher
import easy_uiauto.qt_ui.window
names = ("cv2", "numpy", "tkinter", "pynput", "uiautomation", "easy_uiauto.record",
         "easy_uiauto.knowledge_gui", "mcp.server", "pydantic")
print(json.dumps([name for name in names if name in sys.modules]))
'''
    result = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True, check=True)
    assert json.loads(result.stdout) == []


def test_public_exports_keep_original_function_and_class_identity():
    import easy_uiauto
    from importlib import import_module

    for name in [*easy_uiauto.__all__, "get_control_xpath"]:
        module = import_module(easy_uiauto._EXPORT_MODULES[name], easy_uiauto.__name__)
        assert getattr(easy_uiauto, name) is getattr(module, name)
        assert name in dir(easy_uiauto)


def test_legacy_star_import_keeps_all_exported_names():
    import easy_uiauto

    namespace = {}
    exec("from easy_uiauto import *", namespace)
    assert all(namespace[name] is getattr(easy_uiauto, name) for name in easy_uiauto.__all__)
