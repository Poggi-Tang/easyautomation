import pytest

from easy_uiauto import reference_library as refs
from easy_uiauto.mcp import knowledge
from test_knowledge_library import library,add_control
from test_reference_library import SETTINGS


def long_reference(directory):
    body='# 用户手册\n\n### 保存\n保存当前工程到磁盘。\n\n### 加载\n加载模型到目标机。\n\n'
    body+='### 通信参数\n'+'该字段为通信参数，与工具栏操作无关。'*1200
    knowledge.write_markdown(directory/'references/guide.md',dict(id='guide',kind='reference',name='手册',status='ready',enabled=True),body)


def test_entire_document_is_understood_and_explanations_need_no_literal_match(library,monkeypatch):
    directory=library.directory('demo');long_reference(directory);seen=[]
    def complete(system,payload,**kw):
        seen.append(payload)
        if 'reference' in payload:return dict(understanding='软件管理工程并将模型部署到目标机，分别保存编辑成果和管理运行。')
        if 'reading_notes' in payload:return dict(understanding='工程持久化与目标机运行是两个不同阶段，加载后还需执行。')
        assert '软件' in system and '持久化' in payload['software_understanding']
        return dict(controls=[dict(id=1,description='将正在编辑的工程持久化，供以后继续使用。',
            basis='根据工程持久化流程和Core_save的动作含义归纳',reference_ids=['guide'],uncertainty=False),
            dict(id=2,uncertainty=True,reason='缺少可辨识的控件身份')])
    monkeypatch.setattr(refs,'complete',complete)
    semantics={1:dict(semantic_name='Core_save'),2:dict(semantic_name='ButtonControl')}
    report={};result=refs.descriptions(directory,semantics,report=report,settings=SETTINGS)
    assert set(result)=={1} and result[1]['description_method']=='software-understanding'
    assert 'quote' not in result[1]['reference_evidence'][0]
    assert report['matched']==1 and report['unmatched'][0]['id']==2
    parts=[p['reference']['text'] for p in seen if 'reference' in p]
    original=refs.documents(directory)[0]['body'].split('## 来源原文',1)[0]
    assert ''.join(parts)==original
    before=len(seen);refs.descriptions(directory,semantics,settings=SETTINGS)
    assert len(seen)==before+1 and 'software_understanding' in seen[-1]
    assert (directory/report['understanding_path']).is_file()


def test_edited_reference_invalidates_software_understanding(library,monkeypatch):
    from easy_uiauto.reference_understanding import understanding
    directory=library.directory('demo');long_reference(directory);seen=[]
    def complete(system,payload,**kw):
        seen.append(payload);return dict(understanding='理解结果')
    monkeypatch.setattr(refs,'complete',complete)
    first=understanding(directory,settings=SETTINGS);count=len(seen)
    record=refs.documents(directory)[0];refs.save(directory,'guide',record['body']+'\n修改后的操作知识。')
    second=understanding(directory,settings=SETTINGS)
    assert len(seen)>count and first['document_path']!=second['document_path']


def test_short_document_is_used_in_full_without_keyword_retrieval(library,monkeypatch):
    directory=library.directory('demo')
    knowledge.write_markdown(directory/'references/guide.md',dict(id='guide',status='ready',enabled=True),
                             '# 项目生命周期\n修改内容应在退出之前保留，后续可以接着编辑。')
    def complete(system,payload,**kw):
        assert '退出之前保留' in payload['software_understanding'] and 'catalog' not in payload
        return dict(controls=[dict(id=1,description='保存编辑成果，便于下次继续工作。',
                                  basis='根据项目生命周期与store动作的含义归纳',reference_ids=['guide'])])
    monkeypatch.setattr(refs,'complete',complete)
    result=refs.descriptions(directory,{1:dict(semantic_name='store')},settings=SETTINGS)
    assert result[1]['description'].startswith('保存编辑成果')


def test_code_namespace_is_not_expanded_into_business_meaning(library,monkeypatch):
    directory=library.directory('demo')
    knowledge.write_markdown(directory/'references/guide.md',dict(id='guide',status='ready',enabled=True),'# 软件\n提供配置界面。')
    semantic=dict(semantic_name='InternalFramework_configPage',name='InternalFramework_configPage')
    def complete(system,payload,**kw):
        assert payload['controls'][0]['name']=='configPage'
        assert payload['controls'][0]['native_name'] is None
        return dict(controls=[dict(id=1,description='打开配置页面。',basis='根据配置页动作和软件界面职责归纳')])
    monkeypatch.setattr(refs,'complete',complete)
    result=refs.descriptions(directory,{1:semantic},settings=SETTINGS)
    assert result[1]['description']=='打开配置页面。'
    assert semantic['name']=='InternalFramework_configPage'


def test_overlong_synthesis_retries_concisely_without_reading_notes_again(monkeypatch):
    from easy_uiauto.reference_understanding import _understand
    seen=[]
    def complete(system,payload,**kw):
        seen.append((system,payload))
        if len(seen)==1:raise refs.ReferenceResponseError('truncated',reason='length')
        return dict(understanding='核心软件理解')
    monkeypatch.setattr(refs,'complete',complete)
    notes=dict(reading_notes=['全部阅读结果'])
    assert _understand('整合软件理解',notes,SETTINGS)=='核心软件理解'
    assert len(seen)==2 and seen[0][1] is seen[1][1] and '1500' in seen[1][0]


def test_supplement_preserves_manual_observed_locators_and_images_and_records_unmatched(library,monkeypatch):
    directory=library.directory('demo');base=add_control(library)
    knowledge.write_markdown(directory/'references/guide.md',dict(id='guide',kind='reference',name='手册',status='ready',enabled=True),'# 手册\n保存当前工程。')
    for identifier,fields in [('manual',dict(semantic_source='manual',description='人工说明')),
                               ('effect',dict(description_source='interaction-observation',description='实际效果')),
                               ('unknown',dict(name='',semantic_name='ButtonControl',description='旧推断',description_source='reference-inference'))]:
        knowledge.save_control(directory,{**base,**fields,'id':identifier,'image':'images/keep.png'})
    before={r['id']:r for _,r in knowledge.iter_records(directory,'control')}
    monkeypatch.setattr(refs,'complete',lambda *a,**kw:dict(controls=[dict(id=base['id'],description='保存当前工程',reference_id='guide',quote='保存当前工程。')]))
    report=refs.supplement(directory,settings=SETTINGS)
    after={r['id']:r for _,r in knowledge.iter_records(directory,'control')}
    assert report['matched']==1 and report['preserved']==2
    assert after['manual']['description']=='人工说明' and after['effect']['description']=='实际效果'
    assert after['unknown']['description']=='' and after['unknown']['description_status']=='unmatched'
    for identifier,original in before.items():
        for key in ('location','actions','image','status'):
            assert after[identifier].get(key)==original.get(key)


def test_reference_api_failure_does_not_replace_existing_descriptions(library,monkeypatch):
    directory=library.directory('demo');r=add_control(library);r['description']='已有说明';knowledge.save_control(directory,r)
    knowledge.write_markdown(directory/'references/guide.md',dict(id='guide',status='ready',enabled=True),'# 手册\n保存当前工程。')
    before=knowledge.find_control_record(directory,r['id'])[0].read_bytes()
    def fail(*a,**kw):raise RuntimeError('API unavailable')
    monkeypatch.setattr(refs,'complete',fail)
    with pytest.raises(RuntimeError):refs.supplement(directory,settings=SETTINGS)
    assert knowledge.find_control_record(directory,r['id'])[0].read_bytes()==before
