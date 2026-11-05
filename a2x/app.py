"""A2X 로컬 화면 (IAXG 통합 화면의 2단계): 작업 폴더의 ACI 결과 엑셀을 골라 E2B(R3) XML을 만든다. ACI app.py와 같은 방식."""
import json
import os
import secrets
import sys
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

import main as a2x

ROOT = a2x.ROOT
SETTINGS = ROOT / 'aci_settings.json'  # 작업 폴더는 ACI와 함께 쓴다.
LOG = a2x.LOG
TOKEN = secrets.token_urlsafe(32)
BUSY = threading.Lock()
STATE = {'files': set()}  # 마지막 실행에서 만든 XML (미리보기 허용 목록)


def workspace():
    if '--folder' in sys.argv:  # 시험용: 설정 파일을 바꾸지 않고 다른 폴더를 쓴다 (ACI app.py와 같음)
        return Path(sys.argv[sys.argv.index('--folder') + 1]).resolve()
    if SETTINGS.exists():
        folder = Path(json.loads(SETTINGS.read_text(encoding='utf-8')).get('workspace', ''))
        if folder.is_dir():
            return folder
    return ROOT


def available_files():
    """ACI_outputs의 ACI 결과. 이전 버전 위치(outputs 하위, 작업 폴더 바로 아래)도 찾는다."""
    base = workspace()
    paths = list((base / 'ACI_outputs').glob('*_ACI.xls*')) + list(base.glob('*_ACI.xls*'))
    if (base / 'outputs').is_dir():
        paths += list((base / 'outputs').rglob('*_ACI.xls*'))
    found = {str(p.relative_to(base)).replace('\\', '/'): p for p in paths
             if p.is_file() and not p.name.startswith(('~', '.'))}
    # 최신 파일이 위로
    return dict(sorted(found.items(), key=lambda kv: kv[1].stat().st_mtime, reverse=True))


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        pass  # URL·사례 내용을 HTTP 로그에 남기지 않는다.

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

    def strip_prefix(self):
        # IAXG 화면에서는 A2X API를 /api/a2x/…로 부른다.
        if self.path.startswith('/api/a2x/'):
            self.path = '/api/' + self.path[len('/api/a2x/'):]

    def do_GET(self):
        self.strip_prefix()
        route = urlparse(self.path)
        if route.path == '/':
            self.send(200, (ROOT / 'a2x' / 'index.html').read_bytes(), 'text/html; charset=utf-8')
            return
        if not self.authorized():
            self.send(403, {'error': '접근 권한이 없습니다. IAXG 실행 파일에서 화면을 다시 여세요.'})
            return
        try:
            if route.path == '/api/files':
                self.send(200, {'workspace': str(workspace()), 'files': list(available_files()),
                                'output': str(workspace() / a2x.OUTPUTS), 'version': a2x.VERSION,
                                'sender': a2x.CONFIG['sender_id'], 'receiver': a2x.CONFIG['receiver_id']})
            elif route.path == '/api/xml':
                name = Path(parse_qs(route.query).get('file', [''])[0]).name
                if name not in STATE['files']:
                    raise ValueError('미리볼 수 있는 XML이 아닙니다. 먼저 XML을 만드세요.')
                self.send(200, (workspace() / a2x.OUTPUTS / name).read_bytes(), 'text/plain; charset=utf-8')
            elif route.path == '/api/log':
                log_file = a2x.LOG_FILE
                lines = log_file.read_text(encoding='utf-8').splitlines() if log_file.exists() else []
                self.send(200, {'path': str(log_file), 'lines': lines[-300:]})
            else:
                self.send(404, {'error': '없는 경로입니다.'})
        except Exception as e:
            LOG.exception('요청 오류 %s', route.path)
            self.send(400, {'error': str(e)})

    def do_POST(self):
        self.strip_prefix()
        if not self.authorized():
            self.send(403, {'error': '접근 권한이 없습니다.'})
            return
        try:
            length = int(self.headers.get('Content-Length', '0'))
            if not 0 <= length <= 1024 * 1024:
                raise ValueError('요청이 너무 큽니다.')
            data = json.loads(self.rfile.read(length) or b'{}')
        except Exception as e:
            self.send(400, {'error': str(e)})
            return
        if self.path == '/api/log':  # 화면 오류 기록. 처리 중에도 받는다.
            LOG.error('화면 오류: %s', str(data.get('message', ''))[:2000])
            self.send(200, {'logged': True})
            return
        if self.path == '/api/stop':
            LOG.info('IAXG 종료')
            self.send(200, {'stopped': True})
            threading.Thread(target=self.server.shutdown, daemon=True).start()
            return
        if not BUSY.acquire(blocking=False):
            self.send(409, {'error': '다른 파일을 처리 중입니다.'})
            return
        try:
            if self.path == '/api/workspace':
                folder = Path(data['workspace']).expanduser().resolve()
                if not folder.is_dir():
                    raise ValueError('기존 작업 폴더의 전체 경로를 입력하세요.')
                SETTINGS.write_text(json.dumps({'workspace': str(folder)}, ensure_ascii=False), encoding='utf-8')
                LOG.info('작업 폴더 변경: %s', folder)
                self.send(200, {'workspace': str(folder), 'files': list(available_files())})
            elif self.path == '/api/run':
                source = available_files().get(data.get('name', ''))
                if not source:
                    raise ValueError('선택한 ACI 결과 파일이 없습니다. 목록을 새로고침하세요.')
                summary = a2x.run(source, workspace() / a2x.OUTPUTS)
                STATE['files'] = {c['file'] for c in summary['cases']}
                self.send(200, summary)
            elif self.path == '/api/open':
                folder = workspace() / a2x.OUTPUTS
                folder.mkdir(exist_ok=True)
                os.startfile(folder)  # Windows 탐색기로 저장 폴더를 연다.
                self.send(200, {'opened': str(folder)})
            else:
                self.send(404, {'error': '없는 경로입니다.'})
        except Exception as e:
            LOG.exception('요청 오류 %s', self.path)
            self.send(400, {'error': str(e)})
        finally:
            BUSY.release()


def start():
    server = ThreadingHTTPServer(('127.0.0.1', 0), Handler)
    url = 'http://127.0.0.1:' + str(server.server_port) + '/#' + TOKEN
    runtime = ROOT / '.runtime'
    runtime.mkdir(exist_ok=True)
    (runtime / 'a2x_url.txt').write_text(url, encoding='utf-8')
    LOG.info('A2X %s 화면 시작 · 작업 폴더: %s', a2x.VERSION, workspace())
    if '--no-browser' not in sys.argv:
        webbrowser.open(url)
    try:
        server.serve_forever()
    finally:
        server.server_close()


if __name__ == '__main__':
    start()
