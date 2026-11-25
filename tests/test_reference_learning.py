import pytest

from easy_uiauto import reference_library as refs
from easy_uiauto.mcp import knowledge
from test_manual_region_scan import setup
from test_local_region_learning import local,scan,records


def reference(root):
    directory=knowledge.app_dir('chosen-app',root)
    knowledge.write_markdown(directory/'references/guide.md',dict(id='guide',kind='reference',name='手册',status='ready',enabled=True),
                             '# 保存\n保存当前工程。')


def test_local_scanning_keeps_control_metadata_and_ignores_archived_reference(local,monkeypatch):
    layout,root,*_=local
    scan(layout,root);before=records(root)[0];reference(root)
    seen=[]
    def complete(system,payload,**kw):
        seen.append(payload)
        return dict(controls=[dict(id=1,description='保存当前工程',reference_id='guide',quote='保存当前工程。')])
    monkeypatch.setattr(refs,'complete',complete)
    scan(layout,root);after=records(root)[0]
    assert after['description']==before['description'] and not seen
    assert after['location']==before['location'] and after['actions']==before['actions']
    assert after['function_verification']['executed'] is False


def test_basic_scan_does_not_need_reference_api(local,monkeypatch):
    layout,root,*_=local
    scan(layout,root);reference(root)
    def fail(*a,**kw):raise RuntimeError('reference API unavailable')
    monkeypatch.setattr(refs,'complete',fail)
    result=scan(layout,root)
    assert result['controls_saved_this_scan']==1 and 'reference_description_error' not in result
    assert records(root)[0]['location']['Xpath']


def test_handwritten_control_description_is_not_replaced_by_reference(local,monkeypatch):
    layout,root,*_=local
    scan(layout,root);record=records(root)[0];reference(root)
    record.update(semantic_source='manual',description='我确认的功能说明')
    knowledge.save_control(knowledge.app_dir('chosen-app',root),record)
    monkeypatch.setattr(refs,'complete',lambda *a,**kw:dict(controls=[dict(id=1,description='模型的说明',reference_id='guide',quote='保存当前工程。')]))
    scan(layout,root)
    assert records(root)[0]['description']=='我确认的功能说明'
