"""ACI local review application. Uses only installed LibreOffice + Python stdlib."""
import argparse
import base64
import copy
import hashlib
import json
import logging
import os
import re
import secrets
import shutil
import subprocess
import threading
import uuid
import webbrowser
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from engine import SAE_OPTIONS, analyze, validate_records

VERSION = '0.5'
ROOT = Path(__file__).resolve().parents[1]
OUTPUTS = 'ACI_outputs'
SETTINGS = ROOT / 'aci_settings.json'
WORKSPACE = ROOT
if SETTINGS.exists():
    configured = Path(json.loads(SETTINGS.read_text(encoding='utf-8'))['workspace'])
    if configured.is_dir():
        WORKSPACE = configured
RUNTIME = ROOT / '.runtime'
RUNTIME.mkdir(exist_ok=True)
# 로그에는 파일명·행 번호·오류만 남기고 불만 원문은 남기지 않는다.
LOGS = ROOT / 'logs'
LOGS.mkdir(exist_ok=True)
# ACI와 A2X가 하나의 IAXG 로그 파일을 같이 쓴다 (a2x/main.py와 같은 설정).
LOG_FILE = LOGS / ('iaxg_' + datetime.now().strftime('%Y%m%d') + '.log')
LOG = logging.getLogger('iaxg.aci')
_iaxg_log = logging.getLogger('iaxg')
_iaxg_log.setLevel(logging.INFO)
if not _iaxg_log.handlers:
    _handler = logging.FileHandler(LOG_FILE, encoding='utf-8')
    _handler.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(name)s: %(message)s', '%Y-%m-%d %H:%M:%S'))
    _iaxg_log.addHandler(_handler)
TOKEN = secrets.token_urlsafe(32)
SESSIONS = {}
DOWNLOADS = {}
BUSY = threading.Lock()


def office_job(mode, path, payload=None, output=None, backend='excel'):
    job = RUNTIME / uuid.uuid4().hex
    job.mkdir()
    result_path = job / 'result.json'
    if backend == 'excel':
        bundled_ps = Path.home() / '.cache/codex-runtimes/codex-primary-runtime/dependencies/native/powershell/pwsh.exe'
        powershell = str(bundled_ps) if bundled_ps.exists() else (shutil.which('pwsh') or 'powershell.exe')
        cmd = [powershell, '-NoProfile', '-File', str(ROOT / 'aci' / 'excel.ps1'),
               '-Mode', mode, '-InputPath', str(path), '-JsonOut', str(result_path)]
    else:
        raise ValueError('이 버전은 MS Office Excel로 처리합니다.')
    if payload is not None:
        data = job / 'payload.json'
        data.write_text(json.dumps(payload, ensure_ascii=False), encoding='utf-8')
        cmd += ['-PayloadPath', str(data), '-OutputPath', str(output)]
    flags = subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
    LOG.info('Excel %s 시작: %s (PowerShell: %s)', mode, Path(path).name, Path(powershell).name)
    try:
        try:
            result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                    timeout=150, creationflags=flags)
        except subprocess.TimeoutExpired:
            raise RuntimeError('파일 처리 시간이 초과되었습니다. 대상 파일을 저장 후 닫고 다시 실행하세요.')
        stderr = result.stderr.decode('utf-8', 'replace').strip()
        if stderr:
            LOG.warning('Excel %s 오류 출력 (종료 코드 %s):\n%s', mode, result.returncode, stderr[-4000:])
        if not result_path.exists():
            raise RuntimeError('Excel 파일 처리에 실패했습니다: ' + stderr[-2000:])
        data = json.loads(result_path.read_text(encoding='utf-8'))
        if 'error' in data:
            raise RuntimeError(data['error'])
        if result.returncode:
            raise RuntimeError('Excel 파일 처리가 정상 종료되지 않았습니다.')
        LOG.info('Excel %s 완료: %s', mode, Path(path).name)
        return data
    finally:
        shutil.rmtree(job, ignore_errors=True)  # 중간 파일에 원문이 있으므로 남기지 않는다.


def available_files():
    paths = list(WORKSPACE.glob('IDBI_*.xls')) + list(WORKSPACE.glob('IDBI_*.xlsx'))
    folder = WORKSPACE / 'inputs'
    if folder.exists():
        paths += list(folder.rglob('*.xls')) + list(folder.rglob('*.xlsx'))
    return {str(p.relative_to(WORKSPACE)).replace('\\', '/'): p for p in paths
            if p.is_file() and not p.stem.upper().endswith('_ACI') and not p.name.startswith('~$')}


def load_file(path, backend='excel'):
    workbook = office_job('read', path, backend=backend)
    records = analyze(workbook)
    LOG.info('불러오기: %s · %d행 · AE O %d · 익명화 %d건', Path(path).name, len(records),
             sum(r['ae'] == 'O' for r in records), sum(r['anonymized'] for r in records))
    session = uuid.uuid4().hex
    data = dict(session=session, workbook=workbook, records=records, backend=backend, sae_options=SAE_OPTIONS)
    SESSIONS[session] = copy.deepcopy(data)
    return data


def save_result(data):
    session = SESSIONS.get(data.get('session'))
    if not session:
        raise ValueError('검토 세션이 없습니다. 파일을 다시 불러오세요.')
    records = validate_records(data.get('records', []), session['records'])
    workbook = session['workbook']
    path = Path(workbook['path'])
    # 모든 결과를 작업 폴더의 ACI_outputs 한 곳에 둔다: IDBI_YYYYMMDD_ACI.xls, IDBI_YYYYMMDD_검토기록.json
    folder = WORKSPACE / OUTPUTS
    folder.mkdir(exist_ok=True)
    output = folder / (path.stem + '_ACI' + path.suffix.lower())
    review_log = folder / (path.stem + '_검토기록.json')
    if output.exists() and not data.get('overwrite'):
        raise FileExistsError(f'{OUTPUTS}에 {output.name}이(가) 이미 있습니다.')
    # 검증을 통과한 파일만 기존 결과와 바꾼다.
    temp = folder / f'~{path.stem}_ACI_{uuid.uuid4().hex[:8]}{path.suffix.lower()}'
    payload = dict(source_hash=workbook['source_hash'], records=records)
    try:
        result = office_job('export', path, payload, temp, backend=session['backend'])
        try:
            os.replace(temp, output)
        except PermissionError:
            raise RuntimeError(f'{output.name}이(가) Excel 등에서 열려 있어 바꿀 수 없습니다. 파일을 닫고 다시 저장하세요.')
    finally:
        temp.unlink(missing_ok=True)
    result['output'] = str(output)
    log = dict(version=VERSION, backend=session['backend'], created_at=datetime.now().isoformat(),
               source=workbook['name'], source_sha256=workbook['source_hash'],
               output=output.name, ae_o=sum(r['ae'] == 'O' for r in records),
               ae_x=sum(r['ae'] == 'X' for r in records),
               sae=sum(r['ae'] == 'O' and r.get('sae') not in ('', 'NI') for r in records),
               unconfirmed=sum(not r.get('reviewed') for r in records),
               relatedness_missing=sum(r['ae'] == 'O' and not r.get('relatedness') for r in records),
               verification=result, records=[{k:v for k,v in r.items() if k not in ('values', 'g_text')} for r in records])
    review_log.write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding='utf-8')
    LOG.info('결과 저장: %s · AE O %d / X %d · SAE %d · 검토 완료 전 %d', output, log['ae_o'], log['ae_x'],
             log['sae'], log['unconfirmed'])
    download_id = uuid.uuid4().hex
    DOWNLOADS[download_id] = output
    return dict(output=str(output), review_log=str(review_log), download_id=download_id, name=output.name,
                rows=result['rows'], ae_o=log['ae_o'], ae_x=log['ae_x'], sae=log['sae'],
                unconfirmed=log['unconfirmed'], relatedness_missing=log['relatedness_missing'], verified=True)


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        pass  # Do not put URLs, source text or identifiers into HTTP logs.

    def send(self, status, data, content_type='application/json; charset=utf-8'):
        body = json.dumps(data, ensure_ascii=False).encode('utf-8') if content_type.startswith('application/json') else data
        self.send_response(status)
        self.send_header('Content-Type', content_type)
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-store')
        self.send_header('X-Content-Type-Options', 'nosniff')
        self.send_header('Referrer-Policy', 'no-referrer')
        self.send_header('Content-Security-Policy', "default-src 'self'; script-src 'self' 'unsafe-inline'; style-src 'self' 'unsafe-inline'; connect-src 'self'; img-src 'self' data:; frame-ancestors 'self'")
        self.end_headers()
        self.wfile.write(body)

    def authorized(self):
        if self.headers.get('Host') != '127.0.0.1:' + str(self.server.server_port):
            return False
        origin = self.headers.get('Origin')
        if origin and origin != 'http://127.0.0.1:' + str(self.server.server_port):
            return False
        return secrets.compare_digest(self.headers.get('X-IAXG-Token', ''), TOKEN)

    def do_GET(self):
        route = urlparse(self.path)
        if route.path == '/':
            self.send(200, (ROOT / 'aci' / 'index.html').read_bytes(), 'text/html; charset=utf-8')
            return
        if not self.authorized():
            self.send(403, {'error':'접근 권한이 없습니다. IAXG 실행 파일에서 화면을 다시 여세요.'})
            return
        try:
            if route.path == '/api/files':
                self.send(200, {'files':sorted(available_files()), 'workspace':str(WORKSPACE)})
            elif route.path == '/api/load':
                name = parse_qs(route.query).get('name', [''])[0]
                path = available_files().get(name)
                if not path:
                    raise ValueError('지정한 입력 파일이 없습니다.')
                if not BUSY.acquire(blocking=False):
                    self.send(409, {'error':'다른 파일을 처리 중입니다.'})
                    return
                try:
                    backend = parse_qs(route.query).get('backend', ['excel'])[0]
                    self.send(200, load_file(path, backend))
                finally:
                    BUSY.release()
            elif route.path == '/api/download':
                identity = parse_qs(route.query).get('id', [''])[0]
                path = DOWNLOADS.get(identity)
                if not path:
                    raise ValueError('다운로드할 결과가 없습니다.')
                self.send(200, path.read_bytes(), 'application/vnd.ms-excel')
            elif route.path == '/api/log':
                lines = LOG_FILE.read_text(encoding='utf-8').splitlines() if LOG_FILE.exists() else []
                self.send(200, {'path':str(LOG_FILE), 'lines':lines[-300:], 'version':VERSION})
            elif route.path == '/api/version':
                self.send(200, {'version':VERSION})
            else:
                self.send(404, {'error':'없는 경로입니다.'})
        except Exception as e:
            LOG.exception('요청 오류 %s', route.path)
            self.send(400, {'error':str(e)})

    def do_POST(self):
        global WORKSPACE
        if not self.authorized():
            self.send(403, {'error':'접근 권한이 없습니다.'})
            return
        if self.path == '/api/log':
            # 화면에서 보인 오류 메시지를 기록한다. 처리 중에도 받아야 하므로 BUSY 잠금을 쓰지 않는다.
            try:
                length = int(self.headers.get('Content-Length', '0'))
                data = json.loads(self.rfile.read(min(length, 20000)) or b'{}')
                LOG.error('화면 오류: %s', str(data.get('message', ''))[:2000])
                self.send(200, {'logged':True})
            except Exception as e:
                self.send(400, {'error':str(e)})
            return
        if not BUSY.acquire(blocking=False):
            self.send(409, {'error':'다른 파일을 처리 중입니다.'})
            return
        try:
            length = int(self.headers.get('Content-Length', '0'))
            if not 0 < length <= 30 * 1024 * 1024:
                raise ValueError('전송 크기 제한은 30MB입니다.')
            data = json.loads(self.rfile.read(length))
            if self.path == '/api/upload':
                name = Path(data['name']).name
                if not re.fullmatch(r'IDBI_[0-9]{8}\.(?:xls|xlsx)', name, re.I):
                    raise ValueError('입력 파일명은 IDBI_YYYYMMDD.xls 또는 .xlsx여야 합니다.')
                content = base64.b64decode(data['content'], validate=True)
                if len(content) > 20 * 1024 * 1024:
                    raise ValueError('입력 파일 제한은 20MB입니다.')
                if not (content.startswith(bytes.fromhex('d0cf11e0a1b11ae1')) or content.startswith(b'PK')):
                    raise ValueError('정상적인 XLS/XLSX 파일이 아닙니다.')
                folder = WORKSPACE / 'inputs' / uuid.uuid4().hex
                folder.mkdir(parents=True)
                path = folder / name
                path.write_bytes(content)
                self.send(200, load_file(path, data.get('backend', 'excel')))
            elif self.path == '/api/workspace':
                folder = Path(data['workspace']).expanduser().resolve()
                if not folder.is_dir():
                    raise ValueError('기존 작업 폴더의 전체 경로를 입력하세요.')
                WORKSPACE = folder
                SESSIONS.clear()
                SETTINGS.write_text(json.dumps({'workspace':str(folder)}, ensure_ascii=False), encoding='utf-8')
                LOG.info('작업 폴더 변경: %s', folder)
                self.send(200, {'workspace':str(WORKSPACE), 'files':sorted(available_files())})
            elif self.path == '/api/export':
                self.send(200, save_result(data))
            elif self.path == '/api/stop':
                self.send(200, {'stopped':True})
                threading.Thread(target=self.server.shutdown, daemon=True).start()
            else:
                self.send(404, {'error':'없는 경로입니다.'})
        except FileExistsError as e:
            LOG.info('결과 파일 덮어쓰기 확인 요청: %s', e)
            self.send(409, {'error':str(e), 'exists':True})
        except Exception as e:
            LOG.exception('요청 오류 %s', self.path)
            self.send(400, {'error':str(e)})
        finally:
            BUSY.release()


def main():
    global WORKSPACE
    parser = argparse.ArgumentParser()
    parser.add_argument('--batch')
    parser.add_argument('--backend', choices=['excel'], default='excel')
    parser.add_argument('--folder')
    parser.add_argument('--no-browser', action='store_true')
    parser.add_argument('--connection-file')
    args = parser.parse_args()
    if args.folder:
        WORKSPACE = Path(args.folder).resolve()
        if not WORKSPACE.is_dir():
            raise ValueError('작업 폴더가 없습니다.')
    if args.batch:
        data = load_file(Path(args.batch).resolve(), args.backend)
        result = save_result(data)
        print(json.dumps(result, ensure_ascii=False))
        return
    server = ThreadingHTTPServer(('127.0.0.1', 0), Handler)
    LOG.info('ACI %s 시작 · 작업 폴더: %s', VERSION, WORKSPACE)
    url = 'http://127.0.0.1:' + str(server.server_port) + '/#' + TOKEN
    if args.connection_file:
        Path(args.connection_file).write_text(json.dumps({'url':url,'port':server.server_port,'token':TOKEN}), encoding='utf-8')
    else:
        (RUNTIME / 'current_url.txt').write_text(url, encoding='utf-8')
    if not args.no_browser:
        webbrowser.open(url)
    try:
        server.serve_forever()
    finally:
        server.server_close()


if __name__ == '__main__':
    main()
