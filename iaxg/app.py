"""IAXG (IDBI AE XML Generator): ACI(AE 검토)와 A2X(AE → XML)를 한 화면에서 쓰는 로컬 서버.

화면 / 는 1·2단계 탭, /aci 와 /a2x 는 각 단계 화면이다. API는 ACI가 /api/…, A2X가 /api/a2x/… 를 쓴다.
두 모듈의 처리 코드는 그대로 두고 요청만 나눠 준다.
"""
import argparse
import importlib.util
import sys
import webbrowser
from http.server import ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

VERSION = '0.5'
ROOT = Path(__file__).resolve().parents[1]
sys.path[:0] = [str(ROOT / 'aci'), str(ROOT / 'a2x')]


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


aci = load('aci_app', ROOT / 'aci' / 'app.py')
a2x = load('a2x_app', ROOT / 'a2x' / 'app.py')
a2x.TOKEN = aci.TOKEN  # 한 화면이므로 같은 토큰을 쓴다.
PAGES = {'/': ROOT / 'iaxg' / 'index.html', '/aci': ROOT / 'aci' / 'index.html', '/a2x': ROOT / 'a2x' / 'index.html'}


class Handler(aci.Handler, a2x.Handler):
    """ACI 요청은 그대로, /api/a2x/… 는 A2X 처리기로 넘긴다.

    send·authorized는 ACI 것(앞쪽 부모)을 같이 쓰고, A2X 전용 메서드(strip_prefix)는 A2X에서 받는다.
    """

    def do_GET(self):
        path = urlparse(self.path).path
        if path in PAGES:
            self.send(200, PAGES[path].read_bytes(), 'text/html; charset=utf-8')
        elif path == '/api/version':
            self.send(200, {'iaxg': VERSION, 'aci': aci.VERSION, 'a2x': a2x.a2x.VERSION})
        elif path.startswith('/api/a2x/'):
            a2x.Handler.do_GET(self)
        else:
            aci.Handler.do_GET(self)

    def do_POST(self):
        if self.path.startswith('/api/a2x/'):
            a2x.Handler.do_POST(self)
        else:
            aci.Handler.do_POST(self)


def main():
    parser = argparse.ArgumentParser(description='IAXG 화면 실행')
    parser.add_argument('--no-browser', action='store_true')
    parser.add_argument('--folder', help='시험용 작업 폴더 (설정 파일은 바꾸지 않음)')
    args = parser.parse_args()
    if args.folder:
        aci.WORKSPACE = Path(args.folder).resolve()  # A2X는 같은 --folder 인자를 직접 읽는다.
    server = ThreadingHTTPServer(('127.0.0.1', 0), Handler)
    url = 'http://127.0.0.1:' + str(server.server_port) + '/#' + aci.TOKEN
    (ROOT / '.runtime').mkdir(exist_ok=True)
    (ROOT / '.runtime' / 'iaxg_url.txt').write_text(url, encoding='utf-8')
    aci.LOG.info('IAXG %s 시작 (ACI %s · A2X %s) · 작업 폴더: %s', VERSION, aci.VERSION, a2x.a2x.VERSION, aci.WORKSPACE)
    if not args.no_browser:
        webbrowser.open(url)
    try:
        server.serve_forever()
    finally:
        server.server_close()


if __name__ == '__main__':
    main()
