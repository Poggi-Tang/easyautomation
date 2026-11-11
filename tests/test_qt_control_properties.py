import json
from copy import deepcopy

import pytest
pytest.importorskip('PySide6')
from PySide6.QtWidgets import QApplication,QLineEdit

from easy_uiauto.qt_ui.window import KnowledgeWindow
from easy_uiauto.mcp import knowledge
from easy_uiauto.mcp.protocol import with_control_identity
from test_knowledge_library import library,add_control
from test_qt_knowledge_gui import settle


@pytest.fixture
def editor(library):
    app=QApplication.instance() or QApplication([])
    record=add_control(library,status='verified')
    record.update(verification=dict(location='passed',image='passed'),image='images/reference.png')
    knowledge.save_control(library.directory('demo'),record)
    window=KnowledgeWindow(library.root,remember=False)
    settle(app,lambda:'control:demo:control-one' in window.model.nodes)
    window.goto('control:demo:control-one');app.processEvents()
    try:yield app,window,library,record
    finally:window.dirty=False;window.close();app.processEvents()


def test_editable_properties_update_target_xpath_and_invalidate_location_check(editor):
    app,window,library,record=editor
    expected={'automation_id':'save-id','name':'Save','class_name':'NewButton','control_type':'SplitButtonControl'}
    for key,value in expected.items():
        assert isinstance(window.native_fields[key],QLineEdit) and not window.native_fields[key].isReadOnly()
        window.native_fields[key].setText(value)
    location=json.loads(window.locator.toPlainText())
    assert location['Xpath'][:-1]==record['location']['Xpath'][:-1]
    assert location['Xpath'][-1]['Name']=='Save' and location['Name']=='Save'
    assert location['Xpath'][-1]['AutomationId']=='save-id' and location['AutomationId']=='save-id'
    assert window.dirty and not window.identity_error
    assert window.save();settle(app,lambda:not window.jobs)
    saved=library.read('demo')['controls'][0]
    for key,value in expected.items():assert saved[key]==value
    assert saved['status']=='observed' and saved['verification']['location']=='pending'
    assert saved['image']==record['image'] and saved['actions']==record['actions']
    assert saved['location']['Xpath'][-1]['foundIndex']==record['location']['Xpath'][-1]['foundIndex']


@pytest.mark.parametrize('where',['top','leaf'])
def test_incremental_advanced_edits_update_fields_and_save_both_representations(editor,where):
    app,window,library,record=editor
    location=deepcopy(record['location'])
    for name in ['Open A','Open AB']:
        if where=='top':location['Name']=name
        else:location['Xpath'][-1]['Name']=name
        window.locator.setPlainText(json.dumps(location))
        assert window.native_fields['name'].text()==name and not window.identity_error
    assert window.save();settle(app,lambda:not window.jobs)
    saved=library.read('demo')['controls'][0]
    assert saved['location']['Name']==saved['location']['Xpath'][-1]['Name']=='Open AB'


def test_conflicting_advanced_identity_cannot_be_saved(editor,monkeypatch):
    app,window,library,record=editor
    before=knowledge.find_control_record(library.directory('demo'),record['id'])[0].read_bytes()
    errors=[];monkeypatch.setattr(window,'error',lambda title,error:errors.append(str(error)))
    location=deepcopy(record['location']);location['Name']='First';location['Xpath'][-1]['Name']='Second'
    window.locator.setPlainText(json.dumps(location))
    assert not window.save() and '不一致' in errors[-1]
    assert knowledge.find_control_record(library.directory('demo'),record['id'])[0].read_bytes()==before


def test_blank_control_type_cannot_silently_keep_old_locator(editor,monkeypatch):
    app,window,library,record=editor
    errors=[];monkeypatch.setattr(window,'error',lambda title,error:errors.append(str(error)))
    window.native_fields['control_type'].clear()
    assert not window.save() and 'ControlType' in errors[-1]
    assert library.read('demo')['controls'][0]['location']==record['location']


def test_clearing_criteria_removes_stale_xpath_constraints():
    location=dict(WindowName='UIA Window',Name='Old',AutomationId='old-id',
        Xpath=[dict(ControlType='WindowControl',Name='UIA Window'),
               dict(ControlType='ButtonControl',Name='Old',AutomationId='old-id',foundIndex=3,match_mode='position')],
        PARAMETERS=dict(match_mode='position',window_scope=dict(kind='owned')))
    previous=deepcopy(location)
    edited=with_control_identity(location,dict(name='',automation_id='',class_name='Button',control_type='ButtonControl'))
    assert location==previous and edited['WindowName']=='UIA Window'
    assert edited['Name']==edited['AutomationId']==''
    assert 'Name' not in edited['Xpath'][-1] and 'AutomationId' not in edited['Xpath'][-1]
    assert edited['PARAMETERS']==location['PARAMETERS']
    assert edited['Xpath'][0]==location['Xpath'][0]
    assert edited['Xpath'][-1]['foundIndex']==3 and edited['Xpath'][-1]['match_mode']=='position'


def test_locator_image_preview_uses_explicit_image_and_verify_passes_image_context(editor,monkeypatch):
    from PIL import Image
    app,window,library,record=editor
    image=library.directory('demo')/'images'/'preview.png';image.parent.mkdir(exist_ok=True)
    Image.new('RGB',(24,16),'blue').save(image)
    location=deepcopy(record['location']);location['Img']='images/preview.png'
    window.locator.setPlainText(json.dumps(location));app.processEvents()
    assert window.template_selector.count()==1 and window.template_preview.image_size==(24,16)
    assert window.template_path.text()==str(image.resolve()) and window.template_path.isReadOnly()
    calls=[]
    monkeypatch.setattr(window,'maybe_save',lambda:True)
    monkeypatch.setattr(window,'bound_window',lambda:dict(handle=1))
    monkeypatch.setattr(window,'start_task',lambda *a,**k:calls.append((a,k)))
    window.verify()
    args,options=calls[0]
    assert args[0]=='verify' and args[1]['vault']==str(library.root) and args[1]['app_id']=='demo'
    assert 'image' in args[1] and options['hide']
    location['Img']='missing.png';window.locator.setPlainText(json.dumps(location))
    assert window.template_selector.count()==0 and window.template_preview.image_size==(0,0)


def test_visual_only_locator_can_be_saved_without_creating_a_xpath(editor):
    from easy_uiauto.mcp.protocol import location_from_xpath
    app,window,library,record=editor
    location=location_from_xpath([],parameters={'locator_kind':'visual-observation'})
    location['Img']='images/reference.png'
    window.locator.setPlainText(json.dumps(location))
    assert not window.identity_error and window.save()
    settle(app,lambda:not window.jobs)
    saved=library.read('demo')['controls'][0]
    assert not saved['location']['Xpath'] and saved['location']['Img']=='images/reference.png'
    assert saved['control_type']=='VisualElement'
