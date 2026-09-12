from __future__ import annotations
import json
import os
import threading
from pathlib import Path

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Request
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from .core import APP, ROLES, all_months, bootstrap, build_month, db, decide, digest, event, month_valid, now, register
from .exporter import export_month

app=FastAPI(title='UO 月度工作台',docs_url=None,redoc_url=None)
MUTATION_LOCK=threading.RLock()


@app.middleware('http')
async def local_only(request:Request,call_next):
    if request.method not in ('GET','HEAD','OPTIONS'):
        origin=request.headers.get('origin')
        if origin and origin.rstrip('/')!=str(request.base_url).rstrip('/'):
            return JSONResponse({'detail':'仅接受本机页面发起的操作'},status_code=403)
    response=await call_next(request)
    response.headers['Cache-Control']='no-store'
    response.headers['X-Content-Type-Options']='nosniff'
    return response


@app.exception_handler(ValueError)
async def invalid(request,exc):
    return JSONResponse({'detail':str(exc)},status_code=409)


@app.exception_handler(RuntimeError)
async def failed(request,exc):
    return JSONResponse({'detail':str(exc)},status_code=500)


def enriched(month):
    result=build_month(month)
    with db() as con:
        exports=[dict(r) for r in con.execute('SELECT * FROM exports WHERE month=? ORDER BY created_at DESC',(month,))]
        events=[dict(r) for r in con.execute('SELECT * FROM events WHERE month=? ORDER BY id DESC LIMIT 60',(month,))]
    for row in exports:
        row['name']=Path(row['path']).name
        row['current']=row['fingerprint']==result['fingerprint']
        row['verification']=json.loads(row['verification'])
        row.pop('path')
    result['exports']=exports;result['events']=events
    return result


@app.get('/api/health')
def health():
    return {'app':'uo-monthly-workbench','version':'1.0'}


@app.get('/api/months')
def months():
    bootstrap()
    return {'months':all_months()}


@app.get('/api/months/{month}')
def month_data(month:str):
    return enriched(month)


@app.post('/api/months/{month}/run')
def run(month:str):
    with MUTATION_LOCK:
        return enriched(month)


class Decision(BaseModel):
    kind:str
    item_id:str
    revision:str
    value:str
    note:str=''


@app.post('/api/months/{month}/decision')
def decision(month:str,payload:Decision):
    with MUTATION_LOCK:
        decide(month,**payload.model_dump())
        return enriched(month)


@app.post('/api/months/{month}/export/{kind}')
def export(month:str,kind:str):
    with MUTATION_LOCK:
        row=export_month(month,kind)
        return {'id':row['id'],'name':Path(row['path']).name,'url':'/api/exports/'+row['id']}


@app.get('/api/exports/{export_id}')
def download(export_id:str):
    with db() as con:
        row=con.execute('SELECT * FROM exports WHERE id=?',(export_id,)).fetchone()
    if not row or not Path(row['path']).exists():
        raise HTTPException(404,'输出文件不存在')
    if digest(Path(row['path']).read_bytes())!=row['sha']:
        raise HTTPException(409,'输出文件已在外部修改，请重新生成')
    return FileResponse(row['path'],filename=Path(row['path']).name)


@app.get('/api/sources/{source_id}')
def download_source(source_id:str):
    with db() as con:
        row=con.execute('SELECT * FROM sources WHERE id=?',(source_id,)).fetchone()
    if not row:
        raise HTTPException(404,'资料不存在')
    return FileResponse(row['path'],filename=row['name'])


@app.post('/api/months/{month}/sources')
def upload(month:str,role:str=Form(...),entity:str=Form(''),file:UploadFile=File(...)):
    month_valid(month)
    content=file.file.read(80*1024*1024+1)
    if len(content)>80*1024*1024:
        raise HTTPException(413,'单个文件不能超过80MB')
    with MUTATION_LOCK:
        sid=register(content,file.filename,role,entity,'在工作台上传',month,role=='template')
        if role in ('system','application'):
            with db() as con:
                con.execute('DELETE FROM attachments WHERE month=? AND source_id IN (SELECT id FROM sources WHERE role=?) AND source_id<>?',(month,role,sid))
        return enriched(month)


@app.delete('/api/months/{month}/sources/{source_id}')
def detach_source(month:str,source_id:str):
    with MUTATION_LOCK,db() as con:
        row=con.execute('SELECT * FROM sources WHERE id=?',(source_id,)).fetchone()
        if not row:
            raise HTTPException(404,'资料不存在')
        con.execute('DELETE FROM attachments WHERE month=? AND source_id=?',(month,source_id))
        con.execute('DELETE FROM templates WHERE month=? AND source_id=?',(month,source_id))
        event(con,month,'移出本月资料',row['name'])
    return enriched(month)


@app.get('/api/backup')
def backup():
    # Logical export is portable and contains no active SQLite journal files.
    with db() as con:
        data={t:[dict(r) for r in con.execute(f'SELECT * FROM {t}')] for t in ('sources','attachments','templates','decisions','events')}
    data['exportedAt']=now();data['schemaVersion']=1
    return JSONResponse(data,headers={'Content-Disposition':'attachment; filename="uo-monthly-records.json"'})


app.mount('/',StaticFiles(directory=APP/'monthly/web',html=True),name='web')
