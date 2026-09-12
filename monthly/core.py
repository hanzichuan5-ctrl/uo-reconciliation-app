from __future__ import annotations

import collections
import contextlib
import datetime as dt
import hashlib
import io
import json
import os
import re
import sqlite3
import unicodedata
import zipfile
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import openpyxl

APP = Path(__file__).resolve().parents[1]
ROOT = APP.parent
STATE = Path(os.environ.get('UO_WORKBENCH_STATE', APP / 'monthly-state'))
OUTPUT = ROOT / 'outputs' / '01a09463-3c83-7201-9d6d-69bf2672c926-monthly'
ENTITIES = ['芜湖享游', '上海幻电']
ROLES = {'invoice': '实际发票', 'application': '开票申请', 'system': 'UO收入明细', 'bank': '银行原始流水', 'template': '回款底稿'}
ENGINE_VERSION = 'monthly-1.4'


def now():
    return dt.datetime.now(dt.timezone(dt.timedelta(hours=8))).isoformat(timespec='seconds')


def digest(data):
    return hashlib.sha256(data).hexdigest()


def key(*values):
    return digest(json.dumps(values, ensure_ascii=False, sort_keys=True, default=str).encode())[:24]


def norm(value):
    return re.sub(r'\s+', '', unicodedata.normalize('NFKC', str(value or '')))


def cents(value):
    if value in (None, '', 'N/A'):
        return None
    number = Decimal(str(value).replace(',', ''))
    if not number.is_finite():
        raise ValueError('金额不是有限数值')
    return int((number * 100).quantize(Decimal('1'), rounding=ROUND_HALF_UP))


def date(value):
    if isinstance(value, (dt.date, dt.datetime)):
        return value.isoformat()[:10]
    return str(value or '')[:10]


def identifier(value):
    return str(int(value)) if isinstance(value,(int,float)) and float(value).is_integer() else str(value or '')


def month_valid(month):
    if not re.fullmatch(r'20\d{2}-(0[1-9]|1[0-2])', month):
        raise ValueError('月份须为 YYYY-MM')
    return month


def permitted_name(name):
    name = norm(Path(name).name)
    if name.startswith('1.代理游戏联运渠道回款情况'):
        raise ValueError('此业务查看表已被排除，未读取或导入')
    if name.startswith('~$') or not name.lower().endswith('.xlsx'):
        raise ValueError('请选择原始 .xlsx 文件')


def db():
    STATE.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(STATE / 'workbench.sqlite3', timeout=30)
    con.row_factory = sqlite3.Row
    con.execute('PRAGMA foreign_keys=ON')
    con.executescript('''
      CREATE TABLE IF NOT EXISTS sources(
        id TEXT PRIMARY KEY, name TEXT NOT NULL, sha TEXT NOT NULL, role TEXT NOT NULL,
        entity TEXT NOT NULL, origin TEXT NOT NULL, path TEXT NOT NULL, months TEXT NOT NULL,
        rows INTEGER NOT NULL, imported_at TEXT NOT NULL);
      CREATE TABLE IF NOT EXISTS attachments(
        month TEXT NOT NULL, source_id TEXT NOT NULL REFERENCES sources(id),
        PRIMARY KEY(month, source_id));
      CREATE TABLE IF NOT EXISTS templates(
        month TEXT PRIMARY KEY, source_id TEXT NOT NULL REFERENCES sources(id));
      CREATE TABLE IF NOT EXISTS decisions(
        month TEXT NOT NULL, kind TEXT NOT NULL, item_id TEXT NOT NULL, revision TEXT NOT NULL,
        value TEXT NOT NULL, note TEXT NOT NULL, updated_at TEXT NOT NULL,
        PRIMARY KEY(month, kind, item_id));
      CREATE TABLE IF NOT EXISTS runs(
        month TEXT PRIMARY KEY, fingerprint TEXT NOT NULL, built_at TEXT NOT NULL, payload TEXT NOT NULL);
      CREATE TABLE IF NOT EXISTS exports(
        id TEXT PRIMARY KEY, month TEXT NOT NULL, kind TEXT NOT NULL, fingerprint TEXT NOT NULL,
        path TEXT NOT NULL, sha TEXT NOT NULL, created_at TEXT NOT NULL, verification TEXT NOT NULL);
      CREATE TABLE IF NOT EXISTS events(
        id INTEGER PRIMARY KEY AUTOINCREMENT, month TEXT NOT NULL, time TEXT NOT NULL,
        action TEXT NOT NULL, detail TEXT NOT NULL);
    ''')
    return con


def event(con, month, action, detail):
    con.execute('INSERT INTO events(month,time,action,detail) VALUES(?,?,?,?)', (month, now(), action, detail))


@contextlib.contextmanager
def workbook(source):
    content = Path(source['path']).read_bytes()
    if digest(content) != source['sha']:
        raise ValueError(f"来源副本已变化：{source['name']}，请重新导入")
    w = openpyxl.load_workbook(io.BytesIO(content), read_only=True, data_only=True)
    try:
        yield w
    finally:
        w.close()


def sheet_rows(w, name):
    s = w[name]
    s.reset_dimensions()  # Bank exports may incorrectly declare A1 as the used range.
    return list(s.values)


def source_ref(source, sheet, row):
    return {'id': source['id'], 'file': source['name'], 'sheet': sheet, 'row': row, 'sha': source['sha']}


def getv(row, h, name, default=None):
    i = h.get(name)
    return row[i] if i is not None and i < len(row) else default


def register(content, name, role, entity='', origin='', attach_month=None, select_template=False):
    permitted_name(name)
    if role not in ROLES:
        raise ValueError('未知资料类型')
    sha = digest(content)
    sid = key(sha, role, entity)
    folder = STATE / 'sources'
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / f'{sha}.xlsx'
    temporary = {'id': sid, 'name': name, 'sha': sha, 'path': str(path)}
    if not path.exists():
        path.write_bytes(content)
    months, count = set(), 0
    with workbook(temporary) as w:
        if role == 'invoice':
            rows = sheet_rows(w, '发票基础信息')
            h = {v: i for i, v in enumerate(rows[0]) if v}
            if not {'购买方名称', '开票日期', '价税合计', '销方名称'} <= h.keys():
                raise ValueError('实际发票表头不完整')
            sellers = {str(getv(r, h, '销方名称') or '') for r in rows[1:] if getv(r, h, '销方名称')}
            if entity not in ENTITIES or any(entity not in s for s in sellers):
                raise ValueError('选择的主体与发票销方不一致')
            months = {date(getv(r, h, '开票日期'))[:7] for r in rows[1:] if getv(r, h, '开票日期')}
            count = len(rows)-1
        elif role == 'bank':
            rows = sheet_rows(w, '原始流水对账单明细')
            header = next((i for i, r in enumerate(rows) if '收入金额' in r and '交易日期' in r), None)
            if header is None:
                raise ValueError('银行流水缺少标准表头')
            h = {v: i for i, v in enumerate(rows[header]) if v}
            months = {date(getv(r, h, '交易日期'))[:7] for r in rows[header+1:] if getv(r, h, '交易日期')}
            count = sum(bool(getv(r, h, '交易日期')) for r in rows[header+1:])
            sellers = {str(getv(r,h,'本方账号名称') or '') for r in rows[header+1:] if getv(r,h,'交易日期')}
            if entity not in ENTITIES or any(entity not in s for s in sellers):
                raise ValueError('选择的主体与银行本方账户不一致')
        elif role == 'system':
            rows = sheet_rows(w, '数据')
            if rows[0][:3] != ('月份','游戏','渠道名称') or rows[1][11] != '开票金额' or rows[1][15] != '收入同步状态':
                raise ValueError('UO导出结构发生变化，需调整字段读取')
            months = {str(r[0]) for r in rows[2:] if r[0]}
            count = len(rows)-2
        elif role == 'application':
            valid = [s for s in w if s.title.startswith('20') and ('幻电' in s.title or '芜湖' in s.title)]
            if not valid:
                raise ValueError('开票申请缺少按年、主体区分的工作表')
            count = sum(max(s.max_row-1,0) for s in valid)
        else:
            rows = sheet_rows(w, '回款')  # Never examine the business data of other sheets.
            if not any('往来单位名称' in r and '金额' in r for r in rows):
                raise ValueError('回款页结构不支持')
            months = {date(r[3])[:7] for r in rows if len(r)>3 and r[3] and isinstance(r[2], (int,float))}
            count = sum(len(r)>3 and bool(r[3]) and isinstance(r[2],(int,float)) for r in rows)
    months = sorted(m for m in months if re.fullmatch(r'20\d{2}-\d{2}',m))
    with db() as con:
        created = con.execute('SELECT 1 FROM sources WHERE id=?',(sid,)).fetchone() is None
        con.execute('INSERT OR IGNORE INTO sources VALUES(?,?,?,?,?,?,?,?,?,?)',
                    (sid,name,sha,role,entity,origin,str(path),json.dumps(months),count,now()))
        destinations = [month_valid(attach_month)] if attach_month else (months if role in ('invoice','bank') else [])
        if attach_month and role in ('invoice','bank') and attach_month not in months:
            raise ValueError(f'该文件实际日期为 {"、".join(months)}，与所选月份不符')
        for m in destinations:
            con.execute('INSERT OR IGNORE INTO attachments VALUES(?,?)',(m,sid))
            if created:
                event(con,m,'导入资料',f'{ROLES[role]}：{name}')
        if role == 'template' and select_template and attach_month:
            con.execute('INSERT INTO templates VALUES(?,?) ON CONFLICT(month) DO UPDATE SET source_id=excluded.source_id',(attach_month,sid))
            con.execute('INSERT OR IGNORE INTO attachments VALUES(?,?)',(attach_month,sid))
    return sid


def selected_sources(month):
    with db() as con:
        rows = con.execute('SELECT s.* FROM sources s JOIN attachments a ON a.source_id=s.id WHERE a.month=?',(month,)).fetchall()
        template = con.execute('SELECT source_id FROM templates WHERE month=?',(month,)).fetchone()
    result = [dict(r) for r in rows]
    # Template versions must be explicitly selected. Reference data versions are selected by upload attachment replacement.
    return [r for r in result if r['role']!='template' or template and r['id']==template[0]]


def bootstrap():
    """Import only named, allowed original inputs; no historical results or matching recipes."""
    with db() as con:
        if con.execute('SELECT count(*) FROM sources').fetchone()[0]:
            return
    archive = ROOT/'归档.zip'
    if archive.exists():
        with zipfile.ZipFile(archive) as z:
            for entry in z.infolist():
                try:
                    name = entry.filename if entry.flag_bits & 0x800 else entry.filename.encode('cp437').decode('utf8')
                except UnicodeError:
                    continue
                if '/' in name or not name.endswith('.xlsx'):
                    continue
                role = 'invoice' if '-全量发票查询导出结果-' in name else 'application' if name.startswith('开票申请 ') else 'bank' if name in ('上海幻电7月bankflow.xlsx','原始银行流水明细_芜湖享游_2607.xlsx') else None
                if role:
                    entity = '上海幻电' if '上海幻电' in name else '芜湖享游' if '芜湖享游' in name else ''
                    register(z.read(entry),name,role,entity,f'归档.zip / {name}')
    for name, entity in [('上海幻电_bankflow_08(1).xlsx','上海幻电'),('原始银行流水明细 (1)(1).xlsx','芜湖享游')]:
        p = ROOT/'2026-08回款资料'/name
        if p.exists():
            register(p.read_bytes(),name,'bank',entity,str(p.relative_to(ROOT)))
    p = ROOT/'第三方联运渠道分成-2025-07_2026-07 (4).xlsx'
    if p.exists():
        register(p.read_bytes(),p.name,'system',origin=p.name)
    # This exact file was used in the latest user-requested inspection. It is only copied, never overwritten.
    p = ROOT/'代理游戏联运渠道回款情况-2026.8.31-to业务.xlsx'
    if p.exists():
        register(p.read_bytes(),p.name,'template',origin=p.name,attach_month='2026-08',select_template=True)
    with db() as con:
        months = [r[0] for r in con.execute('SELECT DISTINCT month FROM attachments')]
        refs = [r[0] for r in con.execute("SELECT id FROM sources WHERE role IN ('application','system')")]
        for m in months:
            for sid in refs:
                con.execute('INSERT OR IGNORE INTO attachments VALUES(?,?)',(m,sid))
        event(con,'2026-08','建立月度工作台','已载入7月实际发票与7、8月银行原始流水；日期分别管理。')


def all_months():
    with db() as con:
        return [r[0] for r in con.execute('SELECT DISTINCT month FROM attachments ORDER BY month DESC')]


def month_tokens(value):
    text = str(value or '')
    result = set()
    for y, a, b in re.findall(r'(20\d{2})年\s*(\d{1,2})\s*[-—~至]\s*(\d{1,2})月',text):
        if 1 <= int(a) <= int(b) <= 12:
            result.update(f'{y}-{m:02}' for m in range(int(a),int(b)+1))
    cross=re.search(r'(?<!\d)(20\d{2})年(\d{1,2})月\s*[-—~至]\s*(20\d{2})年(\d{1,2})月',text)
    if cross:
        y1,m1,y2,m2=map(int,cross.groups())
        if 1<=m1<=12 and 1<=m2<=12 and 0<=((y2-y1)*12+m2-m1)<=24:
            for n in range(y1*12+m1-1,y2*12+m2):
                result.add(f'{n//12}-{n%12+1:02}')
    for y, m in re.findall(r'(?<!\d)(20\d{2})[年./-](0?[1-9]|1[0-2])(?=月|[^\d]|$)',text):
        result.add(f'{y}-{int(m):02}')
    for y,m in re.findall(r'(?<!\d)(20\d{2})(0[1-9]|1[0-2])(?=[#&＆、月\s]|$)',text):
        result.add(f'{y}-{int(m):02}')
    return sorted(result)


CHANNELS = {'uc':'UC九游','uc九游':'UC九游','应用宝':'应用宝','vivo':'VIVO','oppo':'OPPO','网易mumu':'网易yofun','网易yofun':'网易yofun'}


def channel(value):
    if isinstance(value,float) and value.is_integer():
        value = int(value)
    text = norm(value)
    return CHANNELS.get(text.lower(),text)


GAME_ALIASES = {'邦邦':'BanG Dream!','公主':'公主连结','坎公':'坎特伯雷公主与骑士唤醒冠军之剑的奇幻冒险',
                '三国':'三国：谋定天下','谋定天下':'三国：谋定天下','百将牌':'三国：百将牌','碧蓝':'碧蓝航线',
                '梦王子':'梦王国与沉睡的100王子','依露希尔':'依露希尔：星晓','纳萨力克':'纳萨力克之王','暖雪':'暖雪手游'}


def games(value):
    return {norm(GAME_ALIASES.get(t.strip(),t.strip())) for t in re.split(r'[、,，/\n]',str(value or '')) if t.strip()}


def applications(sources, year):
    result = []
    for source in sources:
        if source['role']!='application':
            continue
        with workbook(source) as w:
            for s in w:
                if not s.title.startswith(year) or not ('幻电' in s.title or '芜湖' in s.title):
                    continue
                entity = '上海幻电' if '幻电' in s.title else '芜湖享游'
                batch_date = ''
                for rn,r in enumerate(sheet_rows(w,s.title),1):
                    if rn==1 or len(r)<6:
                        continue
                    if r[0]:
                        batch_date = date(r[0])
                    if not r[3] or not isinstance(r[4],(int,float)):
                        continue
                    result.append({'entity':entity,'company':str(r[3]),'channel':channel(r[1]),'games':sorted(games(r[2])),
                                   'gamesText':str(r[2] or ''),'amountCents':cents(r[4]),'months':month_tokens(r[5]),
                                   'periodText':str(r[5] or ''),'date':batch_date,'note':' / '.join(str(v) for v in r[6:] if v),
                                   'source':source_ref(source,s.title,rn)})
    return result


def system_rows(sources):
    result=[]
    for source in sources:
        if source['role']!='system':
            continue
        with workbook(source) as w:
            for rn,r in enumerate(sheet_rows(w,'数据')[2:],3):
                if len(r)<19 or not r[0] or not r[1]:
                    continue
                obj={'month':str(r[0]),'game':str(r[1]),'channel':channel(r[2]),'originalChannel':str(r[2]),
                     'amountCents':cents(r[11]),'syncStatus':r[15] or '未提供','invoiceStatus':r[18] or '未提供',
                     'source':source_ref(source,'数据',rn)}
                obj['businessKey']=key(obj['month'],obj['game'],obj['channel'])
                result.append(obj)
    counts=collections.Counter(r['businessKey'] for r in result)
    for r in result:
        r['duplicate']=counts[r['businessKey']]>1
    return result


def income(sources, month):
    apps=applications(sources,month[:4]); systems=system_rows(sources)
    companies={(a['entity'],norm(a['company'])) for a in apps}
    invoices=[]; skipped=collections.Counter(); seen={}; problems=[]; invoice_game_owners=collections.defaultdict(set)
    system_games={s['game'] for s in systems}
    for source in sources:
        if source['role']!='invoice':
            continue
        with workbook(source) as w:
            rows=sheet_rows(w,'发票基础信息'); h={v:i for i,v in enumerate(rows[0]) if v}
            for rn,r in enumerate(rows[1:],2):
                if date(getv(r,h,'开票日期'))[:7]!=month:
                    continue
                no=str(getv(r,h,'数电发票号码') or getv(r,h,'发票号码') or '')
                buyer=str(getv(r,h,'购买方名称') or '')
                remark=str(getv(r,h,'备注') or '')
                for game in system_games:
                    if norm(game) in norm(remark):invoice_game_owners[norm(game)].add(source['entity'])
                obj={'id':key(source['entity'],no),'entity':source['entity'],'no':no,'company':buyer,
                     'date':date(getv(r,h,'开票日期')),'amountCents':cents(getv(r,h,'价税合计')),
                     'invoiceState':str(getv(r,h,'发票状态') or ''),'remark':remark,'source':source_ref(source,'发票基础信息',rn)}
                if not no or obj['amountCents'] is None:
                    problems.append(f'{source["name"]} 第{rn}行缺少票号或金额'); continue
                if obj['id'] in seen:
                    prev=seen[obj['id']]
                    if any(prev[k]!=obj[k] for k in ('company','date','amountCents','invoiceState','remark')):
                        raise ValueError('同一发票号存在冲突内容：'+no)
                    skipped['重复发票']+=1; continue
                seen[obj['id']]=obj
                if (obj['entity'],norm(buyer)) not in companies:
                    # Explicit consumer game-service receipts are outside the channel workflow, even for corporate buyers.
                    if re.search(r'\d{5,}.*游戏服务',remark) or not re.search(r'公司|企业|工作室|合伙',buyer):
                        skipped['玩家及非渠道发票']+=1; continue
                    obj['scopeUnknown']=True
                invoices.append(obj)
    grouped=collections.defaultdict(list)
    for inv in invoices:
        pool=[a for a in apps if a['entity']==inv['entity'] and norm(a['company'])==norm(inv['company'])]
        exact=[a for a in pool if a['amountCents']==inv['amountCents'] and a['date']<=inv['date']]
        # Deduplicate identical business links while retaining every supporting source.
        meanings={(a['channel'],tuple(a['months'])) for a in exact if a['months']}
        periods=month_tokens(inv['remark']); chosen=[]; issue=''
        if len(meanings)==1:
            ch,ms=next(iter(meanings)); chosen=exact
            if periods and set(periods)!=set(ms):
                issue='发票备注账期与申请账期不一致'
            ms=list(ms)
        else:
            chs={a['channel'] for a in pool}
            if len(chs)==1 and periods:
                ch=next(iter(chs)); ms=periods
                chosen=[a for a in pool if a['channel']==ch and a['months']==ms]
                if not chosen:
                    issue='发票有账期，但申请游戏明细尚未对应'
            else:
                ch=next(iter(chs)) if len(chs)==1 else '渠道待核实'; ms=periods
                issue='存在多个申请对应关系' if len(meanings)>1 else '发票与申请的渠道、账期尚未对应'
        if inv.get('scopeUnknown'):
            issue='该购买方未在本年开票申请中出现，需确认是否属于联运渠道范围'
        inv['linkApps']=chosen; inv['issue']=issue
        groupkey=(inv['entity'],norm(inv['company']),ch,tuple(ms),inv['id'] if issue else '')
        grouped[groupkey].append(inv)
    cases=[]
    for (entity,_,ch,ms,_),ivs in grouped.items():
        ivs.sort(key=lambda x:x['no']); claims={key(a['source']):a for i in ivs for a in i['linkApps']}
        game_set={g for a in claims.values() for g in a['games']}
        selected=[s for s in systems if s['channel']==ch and s['month'] in ms and norm(s['game']) in game_set]
        # Source application game ownership must be unambiguous across the two legal entities.
        owner_map=collections.defaultdict(set)
        for a in apps:
            for g in a['games']:
                owner_map[g].add(a['entity'])
        for g,owners in invoice_game_owners.items():owner_map[g].update(owners)
        for s in systems:
            if s['game'].endswith('-大屏版'):
                owner_map[norm(s['game'])].update(owner_map[norm(s['game'].removesuffix('-大屏版'))])
        # An application's abbreviated game list is not a complete monthly system population.
        coverage_unknown=[s for s in systems if s['channel']==ch and s['month'] in ms and not owner_map[norm(s['game'])] and s['amountCents'] is not None]
        complete_pool=[s for s in systems if s['channel']==ch and s['month'] in ms and owner_map[norm(s['game'])]=={entity} and s['amountCents'] is not None]+coverage_unknown
        if complete_pool:selected=complete_pool
        issues=[i['issue'] for i in ivs if i['issue']]
        if coverage_unknown:
            issues.append('已补充系统游戏候选，需确认主体范围：'+'、'.join(sorted({s['game'] for s in coverage_unknown})))
        if any(i['invoiceState']!='正常' or i['amountCents']<=0 for i in ivs):
            issues.append('包含红字、非正常状态或非正金额发票，需单独核实')
        if any(s['duplicate'] for s in selected):
            issues.append('系统存在重复业务键，不能自动合计确认')
        amount=sum(i['amountCents'] for i in ivs)
        total=sum(s['amountCents'] or 0 for s in selected) if selected else None
        delta=amount-total if total is not None else None
        if ch=='鸿蒙':
            if not issues and claims:
                selected=[];total=None;delta=None
            else:
                issues=[x for x in issues if '系统' not in x]
        elif not selected:
            issues.append('缺少可定位的系统游戏收入记录')
        elif any(s['amountCents'] is None for s in selected):
            issues.append('系统开票金额缺失')
        elif delta:
            issues.append(f'发票与已定位系统明细差额 {delta/100:,.2f} 元')
        if selected and any(s['syncStatus'] not in ('已同步','未同步') for s in selected) and ch!='鸿蒙':
            issues.append('系统收入同步状态缺失或无法识别')
        scope_unknown=any(i.get('scopeUnknown') for i in ivs)
        status='scope' if scope_unknown else 'review' if issues else 'harmony' if ch=='鸿蒙' else 'done' if all(s['syncStatus']=='已同步' for s in selected) else 'ready'
        action='核实对应关系或差额后再处理' if status=='review' else '鸿蒙暂不检查收入同步状态' if status=='harmony' else '已同步，无需重复操作' if status=='done' else '核对后在 UO 点击收入同步'
        if status=='ready' and any(s['syncStatus']=='未同步' and s['invoiceStatus']!='已开票' for s in selected):
            action='核实实际发票后，先补开票状态，再同步收入'
        case={'id':key('income',entity,sorted(i['no'] for i in ivs)),'entity':entity,'company':ivs[0]['company'],'channel':ch,
              'months':list(ms),'games':sorted({s['game'] for s in selected}),'amountCents':amount,'systemTotalCents':total,'deltaCents':delta,
              'status':status,'action':action,'reason':'；'.join(dict.fromkeys(issues)),'scopeUnknown':scope_unknown,
              'invoices':[{k:v for k,v in i.items() if k not in ('linkApps','issue')} for i in ivs],
              'applications':list(claims.values()),'systems':selected}
        candidates=[]
        if not ms and ch!='渠道待核实':
            pool_apps=[a for a in apps if a['entity']==entity and norm(a['company'])==norm(ivs[0]['company']) and a['channel']==ch]
            possible_games={g for a in pool_apps for g in a['games']}
            buckets=collections.defaultdict(list)
            for s in systems:
                if s['channel']==ch and norm(s['game']) in possible_games and s['month']<=month:
                    buckets[s['month']].append(s)
            for m,rs in buckets.items():
                if all(s['amountCents'] is not None and not s['duplicate'] for s in rs) and sum(s['amountCents'] for s in rs)==amount:
                    candidates.append({'month':m,'games':sorted({s['game'] for s in rs}),'systems':rs,
                                       'note':'同公司渠道、申请所列游戏的系统月合计与发票一致，账期仍需核实'})
        case['candidates']=candidates
        if coverage_unknown and delta==0 and len(ms)==1:
            candidates.append({'month':ms[0],'games':case['games'],'systems':selected,'note':'补齐系统游戏候选后金额一致；请同时确认游戏主体范围和账期'})
        if candidates:
            case['reviewType']='candidate'
            case['reason']='已找到同公司渠道和游戏的同额系统候选；确认收入账期后再操作'
            case['action']='查看并确认候选账期'
        elif ch=='鸿蒙':
            case['reviewType']='harmony'
            if not issues and claims:
                case.update(status='harmony',reason='实际发票与申请账期、金额已对应；游戏拆分尚未由系统明细验证，暂不检查同步状态',action='查看开票依据；暂不检查同步状态',games=sorted({a['gamesText'] for a in claims.values()}))
        elif not selected and ch!='渠道待核实' and ms:
            case['reviewType']='missing'
        elif delta:
            case['reviewType']='difference'
        else:case['reviewType']='link'
        case['revision']=key(ENGINE_VERSION,case)
        cases.append(case)
    # Propose a company invoice bundle when the full monthly system total matches exactly.
    # Keep it as a candidate; do not infer an unmatched invoice's period merely from its issue date.
    bundled=[]; consumed=set()
    for c in cases:
        if c['id'] in consumed or c['scopeUnknown'] or not c['months'] or c['deltaCents'] is None:continue
        peers=[p for p in cases if p['id']!=c['id'] and not p['months'] and p['entity']==c['entity'] and p['company']==c['company'] and p['channel']==c['channel'] and not p['scopeUnknown'] and p['id'] not in consumed]
        if len(peers)==1 and c['amountCents']+peers[0]['amountCents']==c['systemTotalCents']:
            p=peers[0];merged=dict(c);invs=c['invoices']+p['invoices']
            merged.update(id=key('income',c['entity'],sorted(i['no'] for i in invs)),invoices=invs,amountCents=c['amountCents']+p['amountCents'],months=[],games=[],systems=[],systemTotalCents=None,deltaCents=None,status='review',reviewType='candidate',reason='两张发票合计与同公司渠道的系统月合计一致；请确认组合账期',action='查看并确认组合候选',
                candidates=[{'month':c['months'][0],'games':c['games'],'systems':c['systems'],'note':'两张实际发票合计与该月系统合计一致，组合账期仍需确认'}])
            merged['revision']=key(ENGINE_VERSION,merged);bundled.append(merged);consumed.update([c['id'],p['id']])
    cases=[c for c in cases if c['id'] not in consumed]+bundled
    # A system record cannot be allocated to two independent invoice groups.
    allocations=collections.Counter(s['businessKey'] for c in cases for s in c['systems'])
    for c in cases:
        if any(allocations[s['businessKey']]>1 for s in c['systems']):
            c.update(status='review',action='核实多票与收入明细的对应范围',reason='同一系统明细出现在多个发票组，暂不重复分配')
            c['revision']=key(ENGINE_VERSION,c)
    cases.sort(key=lambda c:({'ready':0,'review':1,'harmony':2,'done':3,'scope':4}[c['status']],c['entity'],c['company'],c['months']))
    return {'cases':cases,'invoiceCount':len(invoices),'amountCents':sum(i['amountCents'] for i in invoices),
            'inScopeCount':sum(not i.get('scopeUnknown',False) for i in invoices),
            'inScopeAmountCents':sum(i['amountCents'] for i in invoices if not i.get('scopeUnknown')),
            'scopeReviewCount':sum(i.get('scopeUnknown',False) for i in invoices),
            'skipped':dict(skipped),'problems':problems,'rawInvoiceCount':len(seen),
            'stats':dict(collections.Counter(c['status'] for c in cases))}


def receipt_template(source):
    if not source:
        return []
    result=[]; entity=None
    with workbook(source) as w:
        for rn,r in enumerate(sheet_rows(w,'回款'),1):
            if r and r[0] in ENTITIES+['东银河']:
                entity=r[0]
            if len(r)>3 and r[1] and r[3] and isinstance(r[2],(int,float)):
                result.append({'entity':entity,'payer':str(r[1]),'date':date(r[3]),'amountCents':cents(r[2]),'row':rn})
    return result


def receipts(sources,month):
    template=next((s for s in sources if s['role']=='template'),None)
    existing=receipt_template(template)
    existing_counts=collections.Counter((r['entity'],r['payer'],r['date'],r['amountCents']) for r in existing if r['date'][:7]==month)
    rows=[]; seen={}; duplicate_count=0
    for source in sources:
        if source['role']!='bank':
            continue
        with workbook(source) as w:
            values=sheet_rows(w,'原始流水对账单明细'); hi=next(i for i,r in enumerate(values) if '收入金额' in r and '交易日期' in r)
            h={v:i for i,v in enumerate(values[hi]) if v}
            for rn,r in enumerate(values[hi+1:],hi+2):
                d=date(getv(r,h,'交易日期'))
                amount=cents(getv(r,h,'收入金额'))
                if d[:7]!=month or amount is None or amount<=0:
                    continue
                transaction=str(getv(r,h,'交易流水号') or '')
                bankid=identifier(getv(r,h,'系统ID'))
                account=str(getv(r,h,'本方账号') or '')
                payer=str(getv(r,h,'对方账号名称') or '')
                currency=str(getv(r,h,'货币') or '')
                summary=str(getv(r,h,'摘要') or '')
                bank_name=str(getv(r,h,'银行名称') or '')
                category=str(getv(r,h,'分析类别') or '')
                stable=transaction or bankid
                rid=key('bank',source['entity'],account,stable) if stable else key('bank-no-id',source['sha'],rn)
                sig=(d,payer,currency,amount,summary)
                if rid in seen:
                    if seen[rid]!=sig:
                        raise ValueError('重复银行交易标识存在不同内容：'+stable)
                    duplicate_count+=1; continue
                seen[rid]=sig
                obj={'id':rid,'entity':source['entity'],'date':d,'payer':payer,'currency':currency,'amountCents':amount,
                     'summary':summary,'bankName':bank_name,'category':category,'transactionId':stable,'source':source_ref(source,'原始流水对账单明细',rn)}
                signature=(obj['entity'],payer,d,amount)
                if currency=='CNY' and existing_counts[signature]>0:
                    obj.update(status='included',reason='已在指定底稿回款页，逐笔银行来源一致',existing=True)
                    existing_counts[signature]-=1
                elif re.search(r'微信公众号注册退款|对公验证|验证码|账号验证',summary):
                    obj.update(status='excluded',reason='银行摘要明确为注册退款或验证款')
                elif not payer and bank_name in ('支付宝','微信支付'):
                    obj.update(status='excluded',reason='支付平台汇总流水，未作为联运渠道回款纳入')
                elif currency!='CNY':
                    obj.update(status='review',reason='外币来款，当前人民币底稿不纳入或折算')
                elif payer=='深圳市腾讯计算机系统有限公司' and summary.startswith('FPP-'):
                    obj.update(status='review',reason='腾讯FPP来款需按本批范围确认')
                elif payer in ('支付宝支付科技有限公司','财付通支付科技有限公司') or re.search(r'理财|利息|下拨|同名|划转|退款|退回:|扩岗补助|社保公积金|人民法院',payer+' '+summary+' '+category):
                    obj.update(status='excluded',reason='归集、划转、理财、利息或退款等非联运摘要')
                elif re.search(r'联运SDK|游戏.*分成|分成.*游戏|游戏联运',summary):
                    obj.update(status='proposed',reason='银行摘要明确游戏分成，待检查本批纳入范围')
                else:
                    obj.update(status='review',reason='仅凭银行摘要尚不能确定是否纳入')
                if not stable:
                    obj.update(status='review',reason='缺少交易标识，需检查是否为重复流水')
                obj['revision']=key(obj['entity'],obj['date'],obj['payer'],obj['currency'],obj['amountCents'],obj['summary'],bank_name,category,stable)
                rows.append(obj)
    # Same economics without a reliable shared identifier must never silently become duplicate confirmed receipts.
    unmatched=sum(existing_counts.values())
    rows.sort(key=lambda r:(ENTITIES.index(r['entity']),r['date'],r['source']['row']))
    return {'rows':rows,'template':{k:template[k] for k in ('id','name','sha')} if template else None,
            'duplicateCount':duplicate_count,'unmatchedTemplateCount':unmatched,
            'existingCount':sum(r.get('existing',False) for r in rows),
            'templateMonthMismatch': bool(existing and any(r['date'][:7]!=month for r in existing))}


def build_month(month):
    month_valid(month); sources=selected_sources(month)
    # Multiple invoice and bank files may be complementary; duplicate IDs are checked during extraction.
    for role in ('system','application'):
        if sum(s['role']==role for s in sources)>1:
            raise ValueError(f'{ROLES[role]}存在多个版本，请先在资料页选择有效版本')
    result={'month':month,'income':income(sources,month),'receipts':receipts(sources,month),
            'sources':[{k:s[k] for k in ('id','name','role','entity','origin','sha','rows','months','imported_at')} for s in sources]}
    with db() as con:
        decisions=[dict(r) for r in con.execute('SELECT * FROM decisions WHERE month=? ORDER BY kind DESC,item_id',(month,))]
        for d in decisions:
            if d['kind']!='income_link':
                continue
            item=next((r for r in result['income']['cases'] if r['id']==d['item_id']),None)
            if not item or d['revision']!=item['revision']:
                if item:item['staleDecision']=True
                continue
            candidate=next((c for c in item['candidates'] if c['month']==d['value']),None)
            if candidate:
                item['months']=[candidate['month']];item['systems']=candidate['systems'];item['games']=candidate['games']
                item['systemTotalCents']=sum(s['amountCents'] for s in candidate['systems']);item['deltaCents']=0
                item['status']='done' if all(s['syncStatus']=='已同步' for s in item['systems']) else 'ready'
                if any(s['syncStatus'] not in ('已同步','未同步') for s in item['systems']):item['status']='review'
                item['reason']='';item['linkDecision']=d
                item['action']='已同步，无需重复操作' if item['status']=='done' else '核对后在 UO 点击收入同步'
                if item['status']=='ready' and any(s['invoiceStatus']!='已开票' and s['syncStatus']=='未同步' for s in item['systems']):
                    item['action']='核实实际发票后，先补开票状态，再同步收入'
                item['revision']=key(item['revision'],candidate,d['updated_at'])
        for d in decisions:
            if d['kind']=='income_link':continue
            pool=result['income']['cases'] if d['kind']=='income' else result['receipts']['rows']
            item=next((r for r in pool if r['id']==d['item_id']),None)
            if not item:
                continue
            if d['revision']!=item['revision']:
                item['staleDecision']=True; continue
            item['decision']=d
            if d['kind']=='income':
                item['localDone']=d['value']=='done'
            else:
                item['status']=d['value']
                item['reason']='本批已人工确认：'+('纳入回款' if d['value']=='included' else '不纳入回款')
        result['receipts']['stats']=dict(collections.Counter(r['status'] for r in result['receipts']['rows']))
        included=[r for r in result['receipts']['rows'] if r['status']=='included']
        result['receipts']['amountCents']=sum(r['amountCents'] for r in included)
        result['receipts']['entityTotals']={e:{'count':sum(r['entity']==e for r in included),'amountCents':sum(r['amountCents'] for r in included if r['entity']==e)} for e in ENTITIES}
        result['income']['pendingCount']=sum(c['status']=='ready' and not c.get('localDone') for c in result['income']['cases'])
        result['income']['stats']=dict(collections.Counter(c['status'] for c in result['income']['cases']))
        fingerprint=key(ENGINE_VERSION,month,sorted((s['id'],s['sha']) for s in sources),decisions)
        previous=con.execute('SELECT fingerprint,built_at FROM runs WHERE month=?',(month,)).fetchone()
        result['fingerprint']=fingerprint
        result['builtAt']=previous['built_at'] if previous and previous['fingerprint']==fingerprint else now()
        if not previous or previous['fingerprint']!=fingerprint:
            event(con,month,'更新月度结果',f"收入 {result['income']['invoiceCount']} 张；回款已纳入 {len(included)} 笔")
        con.execute('INSERT INTO runs VALUES(?,?,?,?) ON CONFLICT(month) DO UPDATE SET fingerprint=excluded.fingerprint,built_at=excluded.built_at,payload=excluded.payload',
                    (month,fingerprint,result['builtAt'],json.dumps(result,ensure_ascii=False)))
    return result


def decide(month,kind,item_id,revision,value,note=''):
    result=build_month(month)
    if kind not in ('income','income_link','receipt'):
        raise ValueError('未知处理类型')
    pool=result['income']['cases'] if kind.startswith('income') else result['receipts']['rows']
    item=next((r for r in pool if r['id']==item_id),None)
    if not item or item['revision']!=revision:
        raise ValueError('该记录或来源已变化，请刷新后再确认')
    if value!='reset':
        if kind=='income_link' and not any(c['month']==value for c in item['candidates']):
            raise ValueError('对应月份不在当前核对候选中')
        if kind=='income_link':
            candidate=next(c for c in item['candidates'] if c['month']==value)
            selected_keys={s['businessKey'] for s in candidate['systems']}
            if any(selected_keys.intersection(s['businessKey'] for s in c['systems']) for c in pool if c['id']!=item_id):
                raise ValueError('此候选包含其他发票组已定位的系统明细，不能重复分配')
        if kind=='income' and (value!='done' or item['status']!='ready'):
            raise ValueError('只有已核对的待同步事项可记录本地完成')
        if kind=='receipt' and value not in ('included','excluded'):
            raise ValueError('无效回款决定')
        if kind=='receipt' and value=='included' and item['currency']!='CNY':
            raise ValueError('当前回款底稿只支持人民币，不能将外币直接写入')
    with db() as con:
        if value=='reset':
            con.execute('DELETE FROM decisions WHERE month=? AND kind=? AND item_id=?',(month,kind,item_id))
            if kind=='income_link':con.execute("DELETE FROM decisions WHERE month=? AND kind='income' AND item_id=?",(month,item_id))
        else:
            con.execute('INSERT INTO decisions VALUES(?,?,?,?,?,?,?) ON CONFLICT(month,kind,item_id) DO UPDATE SET revision=excluded.revision,value=excluded.value,note=excluded.note,updated_at=excluded.updated_at',
                        (month,kind,item_id,revision,value,note,now()))
        label={'reset':'已撤销','included':'纳入回款','excluded':'不纳入回款','done':'本地记录UO已完成'}.get(value,'确认收入账期 '+value)
        event(con,month,'撤销确认' if value=='reset' else '保存本批确认',f"{item.get('company',item.get('payer'))}：{label}")
    return build_month(month)
