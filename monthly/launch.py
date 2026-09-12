"""Start the local service if needed, then open the workbench."""
import json
import os
import subprocess
import sys
import time
import urllib.request
import webbrowser
from pathlib import Path

APP=Path(__file__).resolve().parents[1]
URL='http://127.0.0.1:8770'

def running():
    try:
        with urllib.request.urlopen(URL+'/api/health',timeout=1) as r:
            return json.load(r).get('app')=='uo-monthly-workbench'
    except Exception:
        return False

if not running():
    state=APP/'monthly-state';state.mkdir(exist_ok=True)
    with (state/'server.log').open('a') as log:
        p=subprocess.Popen([sys.executable,'-m','uvicorn','monthly.server:app','--host','127.0.0.1','--port','8770'],cwd=APP,stdout=log,stderr=log,start_new_session=True)
    for _ in range(35):
        if running():break
        if p.poll() is not None:raise SystemExit('启动失败，请查看 monthly-state/server.log；端口可能被其他程序占用。')
        time.sleep(.2)
    else:raise SystemExit('服务启动超时，请查看 monthly-state/server.log')
webbrowser.open(URL)
print('已打开月度工作台：'+URL)
