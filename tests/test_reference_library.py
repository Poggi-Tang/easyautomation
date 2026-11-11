import io
import json
from email.message import Message
from zipfile import ZipFile

import pytest

from easy_uiauto import reference_library as refs
from easy_uiauto.mcp import knowledge
from test_knowledge_library import library


SETTINGS=dict(api_url="https://example.test/v1",api_key="secret-not-for-vault",api_model="fixture")


@pytest.fixture
def understand(monkeypatch):
    calls=[]
    def complete(system,payload,**kwargs):
        calls.append(payload)
        block=payload['blocks'][0]
        return dict(topics=[dict(title="保存",aliases=["Save"],description="保存当前工程",entry="文件 / 保存",
             preconditions=["已打开工程"],steps=["点击保存"],expected=["工程写入磁盘"],
             source_ids=[block['id']],quote=block['text'][:50])])
    monkeypatch.setattr(refs,"complete",complete)
    return calls


def source(tmp_path):
    p=tmp_path/'手册.md';p.write_text('# 用户手册\n\n保存 Save：点击文件菜单中的保存，将当前工程写入磁盘。',encoding='utf-8');return p


def test_first_import_interprets_and_persists_standard_md_source_and_citations(library,tmp_path,understand):
    original=source(tmp_path)
    record=library.import_reference('demo',original,settings=SETTINGS)
    assert record['status']=='ready' and record['topic_count']==1 and len(understand)==1
    assert '## 功能与操作索引' in record['body'] and '**前置条件**' in record['body']
    assert '**预期结果**' in record['body'] and 'S1' in record['body'] and '来源原文' in record['body']
    assert (library.directory('demo')/record['source_asset']).read_bytes()==original.read_bytes()
    assert library.read('demo')['references'][0]['id']==record['id']
    assert SETTINGS['api_key'] not in (library.directory('demo')/record['document_path']).read_text(encoding='utf-8')


def test_duplicate_import_reuses_edited_md_without_another_api_call(library,tmp_path,understand):
    original=source(tmp_path)
    first=library.import_reference('demo',original,settings=SETTINGS)
    library.save_reference('demo',first['id'],'# 自己修订\n按钮用于保存模型。')
    second=library.import_reference('demo',original,settings=SETTINGS)
    assert second['id']==first['id'] and second['body'].strip()=='# 自己修订\n按钮用于保存模型。'
    assert len(understand)==1
    assert '按钮用于保存模型' in refs.context(library.directory('demo'),'保存')[0]['text']


def test_forced_reinterpretation_preserves_original_name_and_source(library,tmp_path,understand):
    first=library.import_reference('demo',source(tmp_path),settings=SETTINGS)
    second=library.import_reference('demo',library.directory('demo')/first['source_asset'],force=True,settings=SETTINGS)
    assert second['name']==first['name'] and second['source']==first['source'] and len(understand)==2


def test_rejected_citation_never_overwrites_previous_md(library,tmp_path,understand,monkeypatch):
    original=source(tmp_path);first=library.import_reference('demo',original,settings=SETTINGS)
    path=library.directory('demo')/first['document_path'];before=path.read_bytes()
    monkeypatch.setattr(refs,'complete',lambda *a,**k:dict(topics=[dict(title='假功能',description='猜测',source_ids=['S1'],quote='不存在的原文')]))
    with pytest.raises(ValueError,match='来源无法核对'):library.import_reference('demo',original,force=True,settings=SETTINGS)
    assert path.read_bytes()==before


def test_all_long_document_chunks_are_interpreted_with_unique_origins(library,tmp_path,understand,monkeypatch):
    monkeypatch.setattr(refs,'CHUNK_CHARS',80)
    p=tmp_path/'long.md';p.write_text(('测试操作与预期结果。'*35),encoding='utf-8')
    record=library.import_reference('demo',p,settings=SETTINGS)
    assert len(understand)>2 and record['topic_count']==len(understand)
    ids=[b['id'] for call in understand for b in call['blocks']]
    assert len(ids)==len(set(ids)) and all(i in record['body'] for i in ids)


def test_disabled_reference_and_deleted_reference_no_longer_enter_context(library,tmp_path,understand):
    record=library.import_reference('demo',source(tmp_path),settings=SETTINGS)
    library.save_reference('demo',record['id'],record['body'],enabled=False)
    assert refs.context(library.directory('demo'))==[]
    archive=library.archive('demo','reference',record['id'])
    assert archive.is_file() and library.reference_documents('demo')==[]


def test_word_table_keeps_steps_and_expected_cells_together():
    data=io.BytesIO()
    with ZipFile(data,'w') as z:
        z.writestr('word/document.xml','''<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
        <w:p><w:r><w:t>项目管理</w:t></w:r></w:p><w:tbl><w:tr>
        <w:tc><w:p><w:r><w:t>点击新建</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>出现工程对话框</w:t></w:r></w:p></w:tc>
        </w:tr></w:tbl></w:body></w:document>''')
    chunks=refs.extract(data.getvalue(),'.docx')
    assert chunks[0][1]['text']=='点击新建 | 出现工程对话框'
    assert '表格' in chunks[0][1]['origin']


@pytest.mark.parametrize('suffix', ['.mm','.opml','.xmind'])
def test_mind_map_paths_are_retained(suffix):
    if suffix=='.xmind':
        data=io.BytesIO()
        with ZipFile(data,'w') as z:z.writestr('content.json',json.dumps([dict(rootTopic=dict(title='工程',children=dict(attached=[dict(title='新建')])))]))
        raw=data.getvalue()
    else: raw=b'<map><node TEXT="Project"><node TEXT="New"/></node></map>'
    blocks=[b for chunk in refs.extract(raw,suffix) for b in chunk]
    assert any(' / ' in b['origin'] for b in blocks)


def test_pdf_text_pages_preserve_page_origin():
    from pypdf import PdfWriter
    from pypdf.generic import DictionaryObject,NameObject,DecodedStreamObject
    writer=PdfWriter();page=writer.add_blank_page(width=400,height=300)
    font=DictionaryObject({NameObject('/Type'):NameObject('/Font'),NameObject('/Subtype'):NameObject('/Type1'),NameObject('/BaseFont'):NameObject('/Helvetica')})
    page[NameObject('/Resources')]=DictionaryObject({NameObject('/Font'):DictionaryObject({NameObject('/F1'):writer._add_object(font)})})
    stream=DecodedStreamObject();stream.set_data(b'BT /F1 12 Tf 10 250 Td (Save project to disk) Tj ET')
    page[NameObject('/Contents')]=writer._add_object(stream)
    output=io.BytesIO();writer.write(output)
    blocks=refs.extract(output.getvalue(),'.pdf')
    assert blocks[0][0]['origin']=='PDF 第 1 页' and 'Save project' in blocks[0][0]['text']


def test_scanned_pdf_renders_then_uses_local_ocr(monkeypatch):
    from PIL import Image
    from easy_uiauto import local_image_worker
    image=Image.new('RGB',(400,200),'white');out=io.BytesIO();image.save(out,format='PDF')
    sizes=[]
    def ocr(path,**kw):
        with Image.open(path) as frame:sizes.append(frame.size)
        return '扫描手册：保存模型'
    monkeypatch.setattr(local_image_worker,'read_text',ocr)
    blocks=refs.extract(out.getvalue(),'.pdf')
    assert sizes==[(800,400)] and '扫描手册' in blocks[0][0]['text']


def test_web_html_ignores_scripts_and_reads_text(library,understand,monkeypatch):
    class Response(io.BytesIO):
        headers=Message()
        def geturl(self):return 'https://manual.example.test/help'
    Response.headers['Content-Type']='text/html'
    monkeypatch.setattr(refs.request,'urlopen',lambda *a,**k:Response(b'<h1>Manual</h1><p>Save model</p><script>steal secrets</script>'))
    record=library.import_reference('demo','https://manual.example.test/help',settings=SETTINGS)
    assert 'Save model' in record['body'] and 'steal secrets' not in record['body']
    assert record['source']=='https://manual.example.test/help'


def test_control_description_requires_real_quote_and_does_not_change_location(library,tmp_path,understand,monkeypatch):
    record=library.import_reference('demo',source(tmp_path),settings=SETTINGS)
    semantic=dict(semantic_name='CoreManUI_save',observed_value='',location={'Xpath':['keep']},actions=['click'])
    monkeypatch.setattr(refs,'complete',lambda *a,**k:dict(controls=[dict(id=1,description='保存工程',reference_id=record['id'],quote='保存当前工程'),
                                  dict(id=2,description='猜测',reference_id=record['id'],quote='这句话不存在')]))
    result=refs.descriptions(library.directory('demo'),{1:semantic,2:dict(semantic_name='Unknown')},settings=SETTINGS)
    assert set(result)=={1} and result[1]['description_source']=='reference-inference'
    assert semantic['location']=={'Xpath':['keep']} and semantic['actions']==['click']


def test_unconfigured_api_fails_without_creating_ready_reference(library,tmp_path,monkeypatch):
    monkeypatch.setattr(refs.configuration,'load_gui_configuration',lambda:{})
    with pytest.raises(ValueError,match='配置'):library.import_reference('demo',source(tmp_path))
    assert library.reference_documents('demo')==[]


def test_doc_and_external_xml_entities_fail_explicitly():
    with pytest.raises(ValueError,match='另存为'):refs.extract(b'legacy','.doc')
    with pytest.raises(ValueError,match='外部实体'):refs.extract(b'<!DOCTYPE map [<!ENTITY x SYSTEM "file:///secret">]><map/>','.mm')


def test_verified_effect_description_survives_rescan_but_human_edit_wins():
    previous=dict(description='点击后打开新建窗口',description_source='interaction-observation',description_status='observed',description_effect_record='interactions/example.md')
    fresh=dict(description='参考资料推断',semantic_source='uia-metadata')
    assert knowledge.merge_control_state(previous,fresh)['description']==previous['description']
    assert knowledge.merge_control_state(previous,{**fresh,'semantic_source':'manual'})['description']=='参考资料推断'


@pytest.mark.parametrize('quote', [
    'SimuLab 采用上/下位机，支持（HIL）仿真。',
    'SimuLab采用上/下位机,支持(HIL)仿真。',
    'SimuLab 采用上/下位机，支持（HIL）\\n仿真。',
])
def test_pdf_layout_quote_uses_original_source_span(quote):
    text='产品介绍\nSimuLab   采用上/下位机，支持（HIL）\n\n仿真。\n后续内容'
    actual=refs._source_quote(text,quote)
    assert actual=='SimuLab   采用上/下位机，支持（HIL）\n\n仿真。'
    assert actual in text


@pytest.mark.parametrize('quote', ['支持 20 个节点。','不支持 2 个节点。','支持 2 个节点。支持离线仿真。'])
def test_quote_normalization_never_ignores_numbers_words_or_discontinuous_spans(quote):
    assert refs._source_quote('系统支持 2 个节点。中间的限制不能省略。支持离线仿真。',quote) is None


def test_invalid_citation_is_repaired_once_using_real_source(library,tmp_path,monkeypatch):
    p=source(tmp_path);calls=[]
    def complete(system,payload,**kw):
        calls.append(payload)
        if 'invalid_topics' in payload:
            evidence=next(e for e in payload['evidence'] if e['source_id']=='S2')
            return dict(corrections=[dict(index=0,evidence_id=evidence['id'])])
        return dict(topics=[dict(title='保存',description='保存工程',source_ids=['S999'],quote='模型概括的引文')])
    monkeypatch.setattr(refs,'complete',complete)
    result=library.import_reference('demo',p,settings=SETTINGS)
    assert result['topic_count']==1 and len(calls)==2
    assert '> 保存 Save：点击文件菜单中的保存，将当前工程写入磁盘。' in result['body']


def test_retry_reuses_validated_chunks_and_extraction_but_not_failed_chunk(library,tmp_path,monkeypatch):
    monkeypatch.setattr(refs,'CHUNK_CHARS',15)
    p=tmp_path/'resume.md';p.write_text('第一部分：保存模型到磁盘。\n\n第二部分：打开已经保存的模型。',encoding='utf-8')
    calls=[];fail=True
    def complete(system,payload,**kw):
        calls.append(payload)
        if 'invalid_topics' in payload:return dict(corrections=[])
        b=payload['blocks'][0]
        return dict(topics=[dict(title='功能',description='操作说明',source_ids=[b['id']],
             quote='不存在的引用' if fail and b['id']=='S2' else b['text'])])
    monkeypatch.setattr(refs,'complete',complete)
    with pytest.raises(ValueError,match='第 2/2 段'):library.import_reference('demo',p,settings=SETTINGS)
    assert library.reference_documents('demo')==[]
    fail=False;calls.clear()
    monkeypatch.setattr(refs,'extract',lambda *a,**k:pytest.fail('should reuse extracted content'))
    result=library.import_reference('demo',p,settings=SETTINGS)
    assert result['topic_count']==2 and len(calls)==1 and calls[0]['blocks'][0]['id']=='S2'
    cache=library.directory('demo')/'references/.imports'
    assert all(SETTINGS['api_key'] not in f.read_text(encoding='utf-8') for f in cache.rglob('*.json'))


def test_changed_model_does_not_reuse_completed_chunks(library,tmp_path,understand):
    p=source(tmp_path)
    result=library.import_reference('demo',p,settings=SETTINGS)
    library.archive('demo','reference',result['id'])
    library.import_reference('demo',p,settings={**SETTINGS,'api_model':'another'})
    assert len(understand)==2


def test_empty_cover_chunk_does_not_abort_valid_manual(library,tmp_path,monkeypatch):
    monkeypatch.setattr(refs,'CHUNK_CHARS',15)
    p=tmp_path/'cover.md';p.write_text('手册封面\n\n点击保存，将模型写入磁盘。',encoding='utf-8')
    def complete(system,payload,**kw):
        b=payload['blocks'][0]
        return dict(topics=[] if b['id']=='S1' else [dict(title='保存',description='保存模型',source_ids=[b['id']],quote=b['text'])])
    monkeypatch.setattr(refs,'complete',complete)
    result=library.import_reference('demo',p,settings=SETTINGS)
    assert result['topic_count']==1 and '### S1' in result['body'] and '### S2' in result['body']


def test_description_citation_restores_pdf_spacing_instead_of_losing_description(library,tmp_path,understand,monkeypatch):
    record=library.import_reference('demo',source(tmp_path),settings=SETTINGS)
    library.save_reference('demo',record['id'],'# 手册\n\n### 保存\nSimuLab  保存当前工程，\n写入磁盘。')
    monkeypatch.setattr(refs,'complete',lambda *a,**k:dict(controls=[dict(id=1,description='保存工程',
        reference_id=record['id'],quote='SimuLab保存当前工程，写入磁盘。')]))
    result=refs.descriptions(library.directory('demo'),{1:dict(semantic_name='保存')},settings=SETTINGS)
    assert result[1]['reference_evidence'][0]['quote']=='SimuLab  保存当前工程，\n写入磁盘。'


def test_repair_selects_actual_page_excerpt_when_steps_span_pdf_pages(monkeypatch):
    chunk=[dict(id='S1',origin='PDF 第 1 页',text='录制组\n\n1. 打开数据录制。\n\n2. 点击添加录制组，为该子系\n\n1'),
           dict(id='S2',origin='PDF 第 2 页',text='用户手册\n\n统创建一个录制组。\n\n3. 配置录制参数。\n\n2')]
    result=dict(topics=[dict(title='录制组',description='管理录制组',source_ids=['S1','S2'],quote='点击添加录制组，为该子系统创建一个录制组。')])
    def repair(system,payload,**kw):
        chosen=next(e for e in payload['evidence'] if e['text'].startswith('2. 点击'))
        return dict(corrections=[dict(index=0,evidence_id=chosen['id'])])
    monkeypatch.setattr(refs,'complete',repair)
    topics=refs._validated_topics(result,chunk,settings=SETTINGS,progress=lambda _:None,label='第1段')
    assert topics[0]['quote']=='2. 点击添加录制组，为该子系'
    assert topics[0]['source_ids']==['S1','S2']


def test_repair_cannot_invent_source_excerpt_identifier(monkeypatch):
    chunk=[dict(id='S1',origin='原文',text='保存当前工程到磁盘。')]
    result=dict(topics=[dict(title='假功能',description='无依据',source_ids=['S1'],quote='不存在')])
    monkeypatch.setattr(refs,'complete',lambda *a,**kw:dict(corrections=[dict(index=0,evidence_id='S1.E999')]))
    with pytest.raises(ValueError,match='来源无法核对'):
        refs._validated_topics(result,chunk,settings=SETTINGS,progress=lambda _:None,label='第1段')


@pytest.mark.parametrize('source_ids',[None,'S1',1,[None]])
def test_malformed_citation_ids_can_only_be_repaired_by_selecting_real_evidence(source_ids,monkeypatch):
    chunk=[dict(id='S1',origin='原文',text='保存当前工程到磁盘。')]
    result=dict(topics=[dict(title='保存',description='保存工程',source_ids=source_ids,quote='保存工程')])
    monkeypatch.setattr(refs,'complete',lambda *a,**kw:dict(corrections=[dict(index=0,evidence_id='S1.E1')]))
    topics=refs._validated_topics(result,chunk,settings=SETTINGS,progress=lambda _:None,label='第1段')
    assert topics[0]['source_ids']==['S1'] and topics[0]['quote']==chunk[0]['text']


@pytest.mark.parametrize('reason',['length','invalid_json'])
def test_incomplete_model_response_splits_chunk_without_losing_source_or_repeating_completed_work(library,tmp_path,monkeypatch,reason):
    monkeypatch.setattr(refs,'CHUNK_CHARS',4000)
    p=tmp_path/'large.md';p.write_text('点击保存，将模型写入磁盘。'*100+'\n\n'+'点击打开，打开已有模型。'*100,encoding='utf-8')
    calls=[]
    def complete(system,payload,**kw):
        calls.append(payload)
        if len(payload['blocks'])>1:
            raise refs.ReferenceResponseError('response cut off',reason=reason)
        b=payload['blocks'][0]
        return dict(topics=[dict(title='操作',description='模型操作',source_ids=[b['id']],quote=b['text'][:30])])
    monkeypatch.setattr(refs,'complete',complete)
    result=library.import_reference('demo',p,settings=SETTINGS)
    assert len(calls)==3 and result['topic_count']==2
    assert '### S1' in result['body'] and '### S2' in result['body']
    assert '点击保存' in result['body'] and '点击打开' in result['body']


def test_incomplete_single_long_page_is_split_with_original_page_id(library,tmp_path,monkeypatch):
    p=tmp_path/'one-page.md';p.write_text('保存工程到磁盘。'*200,encoding='utf-8')
    calls=[]
    def complete(system,payload,**kw):
        b=payload['blocks'][0];calls.append(b)
        if len(b['text'])>1200:raise refs.ReferenceResponseError('cut off',reason='length')
        return dict(topics=[dict(title='保存',description='保存工程',source_ids=[b['id']],quote=b['text'][:30])])
    monkeypatch.setattr(refs,'complete',complete)
    result=library.import_reference('demo',p,settings=SETTINGS)
    assert len(calls)==3 and result['topic_count']==2
    assert {b['id'] for b in calls}=={'S1'}
    assert calls[1]['text']+calls[2]['text']==calls[0]['text']


def test_complete_detects_length_even_if_truncated_response_contains_valid_json(monkeypatch):
    class Response(io.BytesIO):pass
    raw=json.dumps(dict(choices=[dict(finish_reason='length',message=dict(content='{"topics":[]}'))])).encode()
    monkeypatch.setattr(refs.request,'urlopen',lambda *a,**kw:Response(raw))
    with pytest.raises(refs.ReferenceResponseError) as exc:refs.complete('instruction',{},settings=SETTINGS)
    assert exc.value.reason=='length'


def test_incomplete_short_response_fails_instead_of_recursing_forever(library,tmp_path,monkeypatch):
    calls=[]
    def complete(*a,**kw):
        calls.append(True);raise refs.ReferenceResponseError('bad JSON',reason='invalid_json')
    monkeypatch.setattr(refs,'complete',complete)
    with pytest.raises(refs.ReferenceResponseError,match='已完成章节仍保留'):
        library.import_reference('demo',source(tmp_path),settings=SETTINGS)
    assert len(calls)==1 and library.reference_documents('demo')==[]
