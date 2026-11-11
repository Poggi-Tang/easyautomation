import pytest
pytest.importorskip('PySide6')

from PySide6.QtWidgets import QApplication,QFileDialog
from easy_uiauto.qt_ui.window import KnowledgeWindow
from easy_uiauto.mcp import knowledge
from test_knowledge_library import library
from test_qt_knowledge_gui import settle


@pytest.fixture(scope='module')
def app():return QApplication.instance() or QApplication([])


def test_reference_is_listed_under_usage_and_edited_md_is_saved(app,library):
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
        assert '原功能' in window.usage.toPlainText() and not window.usage.isReadOnly()
        window.usage.setPlainText('# 新功能\n先选择模型，再保存。');assert window.dirty
        assert window.save();settle(app,lambda:not window.jobs)
        assert '先选择模型' in library.reference_documents('demo')[0]['body']
        window.goto('reference:demo:guide');window.toggle_reference();settle(app,lambda:not window.jobs)
        assert library.reference_documents('demo')[0]['enabled'] is False
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


def test_supplement_description_uses_saved_control_without_requiring_live_window(app,library,monkeypatch):
    from test_knowledge_library import add_control
    record=add_control(library);record.update(description_status='unmatched',description_reason='手册未找到明确依据')
    knowledge.save_control(library.directory('demo'),record)
    window=KnowledgeWindow(library.root,remember=False);calls=[]
    try:
        settle(app,lambda:'control:demo:control-one' in window.model.nodes)
        window.goto('control:demo:control-one');app.processEvents()
        assert '手册未找到' in window.description.placeholderText()
        assert '手册未找到' in window.reference_info.text()
        monkeypatch.setattr(window,'start_task',lambda *a,**k:calls.append((a,k)))
        window.supplement_descriptions()
        args,kwargs=calls[0]
        assert args[0]=='supplement-descriptions' and args[1]['control_ids']==['control-one']
        assert not kwargs.get('hide') and 'window' not in args[1]
    finally:window.dirty=False;window.close();app.processEvents()
