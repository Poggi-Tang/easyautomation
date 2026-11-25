import pytest
pytest.importorskip('PySide6')

from PySide6.QtWidgets import QFileDialog
from easy_uiauto.qt_ui.window import KnowledgeWindow
from easy_uiauto.mcp import knowledge
from test_knowledge_library import library
from test_qt_knowledge_gui import app,settle


def test_reference_is_listed_under_usage_as_read_only_archive(app,library):
    directory=library.directory('demo')
    knowledge.write_markdown(directory/'references/guide.md',dict(id='guide',kind='reference',app_id='demo',name='用户手册',status='ready',enabled=True,source='manual.pdf'),
                             '# 原功能\n保存当前模型。')
    window=KnowledgeWindow(library.root,remember=False)
    try:
        settle(app,lambda:'reference:demo:guide' in window.model.nodes)
        item=window.model.nodes['reference:demo:guide']
        assert item.parent() is window.model.nodes['references-root:demo']
        assert item.parent().parent() is window.model.nodes['usage-root:demo']
        window.goto('reference:demo:guide');app.processEvents()
        assert '原功能' in window.usage.toPlainText() and window.usage.isReadOnly()
        assert {b.text() for b in window.context_buttons} == {"重新整理", "删除所选"}
        assert len(window.deletable_nodes([window.current]))==1
    finally:window.dirty=False;window.close();app.processEvents()


def test_import_files_uses_background_job_and_selected_software(app,library,monkeypatch):
    window=KnowledgeWindow(library.root,remember=False);calls=[]
    try:
        settle(app,lambda:'references-root:demo' in window.model.nodes)
        window.goto('references-root:demo');app.processEvents()
        assert {b.text() for b in window.context_buttons}>={'添加参考文件','添加网页'}
        monkeypatch.setattr(QFileDialog,'getOpenFileNames',lambda *a,**k:(['manual.pdf','case.docx'],''))
        monkeypatch.setattr(window,'start_task',lambda *a,**k:calls.append((a,k)))
        window.import_reference_files()
        args,kwargs=calls[0]
        assert args[0]=='import-reference' and args[1]['app_id']=='demo'
        assert args[1]['sources']==['manual.pdf','case.docx'] and not kwargs.get('hide')
    finally:window.dirty=False;window.close();app.processEvents()


def test_business_process_can_start_without_a_region(app,library,monkeypatch):
    from easy_uiauto.business_processes import from_topics
    from easy_uiauto.qt_ui.dialogs import DeepLearningDialog
    from PySide6.QtWidgets import QDialog

    process, = from_topics([dict(title="新建工程", description="创建本地工程", entry="文件菜单",
                                 steps=["点击文件菜单"], expected=["工程显示在树中"], quote="点击文件")],
                           dict(id="guide", name="测试手册"))
    knowledge.write_markdown(library.directory("demo") / "business-processes/guide.md", {"processes": [process]}, "")
    window=KnowledgeWindow(library.root,remember=False)
    calls=[]
    key=f"business:demo:{process['id']}"
    try:
        settle(app,lambda:key in window.model.nodes)
        window.goto(key)
        app.processEvents()
        assert "待学习" in window.usage.toPlainText() and "测试手册" in window.usage.toPlainText()
        assert "学习业务流程" in {b.text() for b in window.context_buttons}
        monkeypatch.setattr(window,"bound_window",lambda:{"handle": 1})
        monkeypatch.setattr(DeepLearningDialog,"exec",lambda self:QDialog.DialogCode.Accepted)
        monkeypatch.setattr(window,"start_task",lambda *args,**kwargs:calls.append((args,kwargs)))
        window.deep_learn()
        settle(app,lambda:bool(calls))
        args,kwargs=calls[0]
        assert args[0]=="deep-learn" and args[1]["region_id"]=="" and args[1]["business_process_id"]==process["id"]
        assert kwargs["hide"] and not library.read("demo")["regions"]
    finally:
        window.dirty=False
        window.close()
        app.processEvents()
