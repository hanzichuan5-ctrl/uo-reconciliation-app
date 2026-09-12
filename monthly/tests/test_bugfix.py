from monthly.core import *

def test_combined_period_and_candidates_exportable():
    assert month_tokens('202412&202605#数字商品及联运')==['2024-12','2026-05']
    data=income(selected_sources('2026-07'),'2026-07')
    assert data['inScopeCount']==53 and data['inScopeAmountCents']==4890756364
    c=next(c for c in data['cases'] if c['amountCents']==214200)
    assert c['reviewType']=='candidate' and '缺少' not in c['reason']
    c=next(c for c in data['cases'] if c['amountCents']==1082454934)
    assert c['status']=='harmony' and '缺少可定位' not in c['reason']
    c=next(c for c in data['cases'] if c['amountCents']==605949563)
    assert c['deltaCents']==0
    c=next(c for c in data['cases'] if c['amountCents']==367679724)
    assert c['deltaCents']==-8350
