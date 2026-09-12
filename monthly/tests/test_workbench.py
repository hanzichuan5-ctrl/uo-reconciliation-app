import collections
import json
import sqlite3
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from monthly import core, exporter
from monthly.server import app


@pytest.fixture
def state(tmp_path,monkeypatch):
    live=core.STATE/'workbench.sqlite3'
    target=tmp_path/'state';target.mkdir()
    with sqlite3.connect(live) as a,sqlite3.connect(target/'workbench.sqlite3') as b:
        a.backup(b)
    monkeypatch.setattr(core,'STATE',target)
    monkeypatch.setattr(exporter,'OUTPUT',tmp_path/'exports')
    return tmp_path


def test_periods_and_money():
    assert core.cents(542.70000000000005)==54270
    assert core.month_tokens('2025年12月-2026年3月')==['2025-12','2026-01','2026-02','2026-03']
    assert core.month_tokens('2026年4-5月')==['2026-04','2026-05']
    assert core.month_tokens('红字信息表编号2020550511223000')==[]
    assert core.month_tokens('202605#游戏')==['2026-05']


def test_actual_data_and_idempotency(state):
    july=core.build_month('2026-07');aug=core.build_month('2026-08')
    assert july['income']['inScopeCount']==53
    assert july['income']['inScopeAmountCents']==4890756364
    assert aug['income']['invoiceCount']==0
    assert aug['receipts']['stats']['included']==49
    assert aug['receipts']['amountCents']==4418260485
    assert all('1.代理游戏联运渠道回款情况' not in s['name'] for s in aug['sources']+july['sources'])
    assert core.build_month('2026-08')['fingerprint']==aug['fingerprint']
    assert core.build_month('2026-08')['builtAt']==aug['builtAt']
    source=next(s for s in core.selected_sources('2026-08') if s['role']=='bank')
    sid=core.register(Path(source['path']).read_bytes(),'重复上传.xlsx','bank',source['entity'],attach_month='2026-08')
    assert sid==source['id']
    assert core.build_month('2026-08')['fingerprint']==aug['fingerprint']


def test_confirmation_persists_and_foreign_currency_rejected(state):
    data=core.build_month('2026-08')
    row=next(r for r in data['receipts']['rows'] if r['status']=='review' and r['currency']=='CNY')
    updated=core.decide('2026-08','receipt',row['id'],row['revision'],'included')
    assert updated['receipts']['amountCents']==data['receipts']['amountCents']+row['amountCents']
    assert core.build_month('2026-08')['fingerprint']==updated['fingerprint']
    core.decide('2026-08','receipt',row['id'],row['revision'],'reset')
    assert core.build_month('2026-08')['receipts']['amountCents']==data['receipts']['amountCents']
    with pytest.raises(ValueError,match='已变化'):
        core.decide('2026-08','receipt',row['id'],'stale','included')
    foreign=next(r for r in data['receipts']['rows'] if r['currency']!='CNY')
    with pytest.raises(ValueError,match='人民币'):
        core.decide('2026-08','receipt',foreign['id'],foreign['revision'],'included')


def test_candidate_is_not_automatically_confirmed_and_can_be_undone(state):
    data=core.build_month('2026-07')
    row=next(c for c in data['income']['cases'] if c['company']=='北京慕远科技有限公司' and c['amountCents']==214200)
    assert row['status']=='review' and row['months']==[]
    result=core.decide('2026-07','income_link',row['id'],row['revision'],'2026-05')
    linked=next(c for c in result['income']['cases'] if c['id']==row['id'])
    assert linked['status']=='ready' and linked['deltaCents']==0
    completed=core.decide('2026-07','income',linked['id'],linked['revision'],'done')
    case=next(c for c in completed['income']['cases'] if c['id']==row['id'])
    assert case['localDone'] and case['systems'][0]['syncStatus']=='未同步'
    undone=core.decide('2026-07','income_link',case['id'],case['revision'],'reset')
    case=next(c for c in undone['income']['cases'] if c['id']==row['id'])
    assert case['status']=='review' and not case.get('localDone')


def test_stale_source_decision_is_not_reused(state,monkeypatch):
    data=core.build_month('2026-08');row=next(r for r in data['receipts']['rows'] if r['status']=='review' and r['currency']=='CNY')
    core.decide('2026-08','receipt',row['id'],row['revision'],'included')
    original=core.receipts
    def changed(sources,month):
        result=original(sources,month)
        target=next(r for r in result['rows'] if r['id']==row['id'])
        target['amountCents']+=100;target['revision']='new-source-revision'
        return result
    monkeypatch.setattr(core,'receipts',changed)
    result=core.build_month('2026-08');target=next(r for r in result['receipts']['rows'] if r['id']==row['id'])
    assert target['staleDecision'] and target['status']=='review'


def test_restricted_file_and_wrong_month_do_not_enter_batch(state):
    with pytest.raises(ValueError,match='已被排除'):
        core.register(b'not-read','1.代理游戏联运渠道回款情况(20260909）.xlsx','template')
    source=next(s for s in core.selected_sources('2026-08') if s['role']=='bank')
    with pytest.raises(ValueError,match='实际日期'):
        core.register(Path(source['path']).read_bytes(),'错月份.xlsx','bank',source['entity'],attach_month='2026-09')
    assert core.selected_sources('2026-09')==[]


def test_api_restart_read_and_cross_origin(state):
    with TestClient(app) as client:
        assert client.get('/api/health').json()['app']=='uo-monthly-workbench'
        assert client.get('/api/months/2026-08').json()['receipts']['stats']['included']==49
        assert client.post('/api/months/2026-08/run',headers={'Origin':'https://example.com'}).status_code==403
    with TestClient(app) as client:
        assert client.get('/api/months/2026-08').json()['receipts']['amountCents']==4418260485


def test_receipt_copy_and_expansion_preserve_other_parts(state):
    data=core.build_month('2026-08')
    source=next(s for s in core.selected_sources('2026-08') if s['role']=='template')
    selected=[r for r in data['receipts']['rows'] if r['status']=='included']
    no_op=state/'same.xlsx';v=exporter.write_receipt_copy(source,selected,no_op)
    assert v['changedParts']==[]
    assert no_op.read_bytes()==Path(source['path']).read_bytes()
    # Exercise real row expansion on disposable output with synthetic test-only receipts.
    added=[dict(selected[-1],id=f'test-{i}',payer=f'测试付款方{i}',amountCents=100+i) for i in range(5)]
    expanded=state/'expanded.xlsx'
    v=exporter.write_receipt_copy(source,selected+added,expanded)
    with zipfile.ZipFile(source['path']) as original,zipfile.ZipFile(expanded) as output:
        part=exporter.receipt_part(original)[0]
        assert all(original.read(n)==output.read(n) for n in original.namelist() if n not in (part,'xl/calcChain.xml'))
        def other_chain(z):
            result=[];current=None;receipt_sheetid=exporter.receipt_part(z)[2]
            for c in ET.fromstring(z.read('xl/calcChain.xml')):
                current=int(c.attrib['i']) if 'i' in c.attrib else current
                if current!=receipt_sheetid:result.append(dict(c.attrib))
            return result
        assert other_chain(original)==other_chain(output)
    assert v['count']==54
