import pytest

pytest.importorskip("PySide6")
from PySide6.QtWidgets import QDialog

from easy_uiauto.qt_ui.dialogs import DeepLearningDialog
from easy_uiauto.qt_ui.widgets import StepEditor
from test_knowledge_library import library
from test_region_picker import picked
from test_qt_knowledge_gui import app, settle
from test_qt_responsiveness import window


def test_deep_dialog_and_toolbar_use_separate_job(app,library,window,tmp_path,monkeypatch):
    identifier=library.save_picked_region("demo",picked(tmp_path))
    library.confirm_layout("demo")
    window.reload("region:demo:"+identifier);settle(app,lambda:not window.jobs)
    buttons={b.text():b for b in window.context_buttons}
    assert "基本学习" in buttons and "深度学习" in buttons
    calls=[]
    monkeypatch.setattr(window,"bound_window",lambda:dict(handle=123))
    monkeypatch.setattr(window,"start_task",lambda *a,**k:calls.append((a,k)))
    def choose(dialog):
        dialog.inputs.setPlainText('{"工程名称":"测试"}')
        dialog.commit.setChecked(True)
        return QDialog.DialogCode.Accepted
    monkeypatch.setattr(DeepLearningDialog,"exec",choose)
    buttons["深度学习"].click()
    settle(app,lambda:bool(calls))
    request=calls[0][0][1]
    assert calls[0][0][0]=="deep-learn" and request["region_id"]==identifier
    assert request["learning_options"]["allow_commit"] and request["learning_options"]["inputs"]=={"工程名称":"测试"}


@pytest.mark.parametrize("entry", ["toolbar", "context-menu"])
def test_control_deep_learning_uses_fixed_entry_and_restores_selection(app,library,window,tmp_path,monkeypatch,entry):
    from PySide6.QtGui import QAction
    from easy_uiauto.qt_ui import window as window_module
    region=library.save_picked_region("demo",picked(tmp_path))
    control=library.save_captured_control("demo",picked(tmp_path),name="新建工程",region_id=region)
    library.confirm_layout("demo")
    key="control:demo:"+control["id"]
    window.reload(key);settle(app,lambda:not window.jobs)
    calls=[];reloads=[]
    monkeypatch.setattr(window,"bound_window",lambda:dict(handle=123))
    monkeypatch.setattr(window,"start_task",lambda *args,**kwargs:calls.append((args,kwargs)))
    def choose(dialog):
        assert "新建工程" in dialog.windowTitle()
        assert dialog.scope.isHidden()
        for mode in ("new","resume","improve","update"):
            dialog.mode.setCurrentIndex(dialog.mode.findData(mode))
            assert dialog.values()["entry_control_id"]==control["id"]
        dialog.mode.setCurrentIndex(dialog.mode.findData("new"))
        return QDialog.DialogCode.Accepted
    monkeypatch.setattr(DeepLearningDialog,"exec",choose)
    if entry=="toolbar":
        value=next(b for b in window.context_buttons if b.text()=="深度学习")
        assert not value.icon().isNull();value.click()
    else:
        class Menu:
            def __init__(self,parent):self.items=[]
            def addAction(self,*args):
                action=QAction(args[-1],window)
                if len(args)>1:action.setIcon(args[0])
                self.items.append(action);return action
            def exec(self,*args):
                action=next(a for a in self.items if a.text()=="深度学习")
                assert action.isEnabled() and not action.icon().isNull();action.trigger()
        monkeypatch.setattr(window_module,"QMenu",Menu)
        window.tree_menu(window.tree.visualRect(window.tree.currentIndex()).center())
    settle(app,lambda:bool(calls))
    args,kwargs=calls[0];request=args[1]
    assert args[0]=="deep-learn" and kwargs["hide"]
    assert request["region_id"]==region and request["learning_options"]["entry_control_id"]==control["id"]
    monkeypatch.setattr(window,"reload",lambda key,**kwargs:reloads.append(key))
    args[2](dict(message="完成"))
    assert reloads==[key]


def test_control_missing_region_reports_error_without_starting_task(app,library,window,monkeypatch):
    from test_knowledge_library import add_control
    control=add_control(library)
    window.reload("control:demo:"+control["id"]);settle(app,lambda:not window.jobs)
    errors=[]
    monkeypatch.setattr(window,"error",lambda title,error:errors.append(str(error)))
    monkeypatch.setattr(window,"bound_window",lambda:pytest.fail("should fail before desktop binding"))
    window.deep_learn()
    assert "所属区域" in errors[0]


def test_call_editor_preserves_reference_and_parameter_bindings(app):
    editor=StepEditor(lambda:None,choose_workflow=lambda:dict(id="shared",keyword="公共填写"))
    editor.add_call()
    assert editor.steps[0]["workflow_id"]=="shared" and editor.steps[0]["action"]=="call"
    editor.value.setText('{"名称":"${工程名称}"}');editor.edit()
    assert editor.steps[0]["parameters"]=={"名称":"${工程名称}"}
    editor.description.setText("复用填写步骤");editor.edit()
    assert editor.steps[0]["action"]=="call" and "target" not in editor.steps[0]
    assert editor.tree.currentItem().text(1)=="shared" and "复用填写步骤" in editor.tree.currentItem().text(0)
    assert not editor.value_source.isEnabled()
    editor.value.setText('{"broken":');editor.edit()
    with pytest.raises(ValueError,match="JSON"):editor.validated_steps()
    editor.select()
    assert editor.value.text()=='{"broken":'
    editor.value.setText('{}');editor.edit()
    assert editor.validated_steps()[0]["parameters"]=={}
    editor.deleteLater()


def test_entry_editor_does_not_offer_workflow_calls(app):
    editor=StepEditor(lambda:None)
    editor.select()
    assert editor.kind.findData("call")==-1
    editor.deleteLater()


def test_relearning_dialog_loads_session_and_passes_scope_and_explicit_retry(app):
    from PySide6.QtCore import Qt
    graph=dict(id="deep-one",updated_at="2026-11-13",status="partial",options=dict(inputs={"名称":"已有测试"},instructions="参考说明"),
        states={"root":dict(name="主界面",branches=[dict(id="pending",name="提交",status="uncertain",reason="结果不明"),
                                                 dict(id="done",name="已完成",status="done")])})
    dialog=DeepLearningDialog(dict(name="区域"),[dict(id="flow",keyword="流程")],sessions=[graph],
                              controls=[dict(id="control",name="控件")])
    assert dialog.values()["mode"]=="resume" and dialog.values()["inputs"]=={"名称":"已有测试"}
    assert dialog.pending.topLevelItemCount()==1
    dialog.mode.setCurrentIndex(dialog.mode.findData("improve"))
    assert dialog.pending.topLevelItemCount()==2
    dialog.scope.topLevelItem(0).setCheckState(0,Qt.CheckState.Checked)
    dialog.pending.topLevelItem(0).setCheckState(0,Qt.CheckState.Checked)
    result=dialog.values()
    assert result["session_id"]=="deep-one" and result["control_ids"]==["control"] and result["retry_branches"]==["pending"]
    dialog.mode.setCurrentIndex(dialog.mode.findData("update"))
    assert not dialog.scope.isEnabled() and dialog.values()["control_ids"]==[]
    dialog.mode.setCurrentIndex(dialog.mode.findData("new"))
    assert dialog.values()["session_id"]=="" and dialog.values()["retry_branches"]==[]
    dialog.deleteLater()
