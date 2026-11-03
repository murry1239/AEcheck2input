"""AAI (AE Auto Input): ACI 결과 엑셀의 H열 O 행 → 사례별 ICH E2B(R3) XML."""
import argparse
import html
import json
import logging
import os
import re
import shutil
import subprocess
import sys
import uuid
import webbrowser
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'aci'))  # 익명화 규칙은 ACI와 같은 것을 쓴다.
from anonymize import anonymize  # noqa: E402
from e2b import build, is_ae, parse_case  # noqa: E402

VERSION = '0.1'
CONFIG = json.loads((ROOT / 'aai' / 'aai_config.json').read_text(encoding='utf-8'))
_local = ROOT / 'aai' / 'aai_config.local.json'  # 담당자 이름·이메일 등 PC별 값 (저장소에 올리지 않음)
if _local.exists():
    CONFIG.update(json.loads(_local.read_text(encoding='utf-8')))
LOGS = ROOT / 'logs'
LOGS.mkdir(exist_ok=True)
LOG = logging.getLogger('aai')
LOG.setLevel(logging.INFO)
_handler = logging.FileHandler(LOGS / ('aai_' + datetime.now().strftime('%Y%m%d') + '.log'), encoding='utf-8')
_handler.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(message)s', '%Y-%m-%d %H:%M:%S'))
LOG.addHandler(_handler)


def read_workbook(path):
    """Excel로 ACI 결과 A~M열을 읽는다. ACI와 같은 PowerShell을 쓴다."""
    job = ROOT / '.runtime' / ('aai_' + uuid.uuid4().hex)
    job.mkdir(parents=True)
    result_path = job / 'result.json'
    bundled = Path.home() / '.cache/codex-runtimes/codex-primary-runtime/dependencies/native/powershell/pwsh.exe'
    powershell = str(bundled) if bundled.exists() else (shutil.which('pwsh') or 'powershell.exe')
    cmd = [powershell, '-NoProfile', '-File', str(ROOT / 'aai' / 'read_aci.ps1'),
           '-InputPath', str(path), '-JsonOut', str(result_path)]
    flags = subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
    try:
        result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=180, creationflags=flags)
        stderr = result.stderr.decode('utf-8', 'replace').strip()
        if stderr:
            LOG.warning('Excel 읽기 오류 출력 (종료 코드 %s):\n%s', result.returncode, stderr[-4000:])
        if not result_path.exists():
            raise RuntimeError('Excel 파일 읽기에 실패했습니다: ' + stderr[-1000:])
        data = json.loads(result_path.read_text(encoding='utf-8'))
        if 'error' in data:
            raise RuntimeError(data['error'])
        return data
    except subprocess.TimeoutExpired:
        raise RuntimeError('Excel 읽기 시간이 초과되었습니다. 대상 파일을 닫고 다시 실행하세요.')
    finally:
        shutil.rmtree(job, ignore_errors=True)  # 중간 파일에 원문이 있으므로 남기지 않는다.


def load_schema():
    path = ROOT / CONFIG['schema_path']
    try:
        from lxml import etree
    except ImportError:
        return None, 'lxml이 없어 XSD 검증을 건너뛰었습니다.'
    if not path.exists():
        return None, f'XSD 파일이 없어 검증을 건너뛰었습니다: {path}'
    return etree.XMLSchema(etree.parse(str(path))), ''


def convert(workbook, out_dir, meddra, now=None):
    """workbook: read_aci.ps1 결과. H열 O 행만 XML로 만든다."""
    now = now or datetime.now()
    m = re.search(r'IDBI_(\d{8})', workbook['name'])
    file_date = m.group(1) if m else ''
    schema, schema_note = load_schema()
    cases, skipped = [], 0
    out_dir.mkdir(parents=True, exist_ok=True)
    for sheet in workbook['sheets']:
        for record in sheet['records']:
            values = record['values']
            if not is_ae(values[7] if len(values) > 7 else ''):
                skipped += 1
                continue
            case = parse_case(values, int(record['row']), file_date, CONFIG, anonymize)
            xml, issues = build(case, CONFIG, meddra, now)
            # cubeSAFETY 출력 파일명처럼: ICSR_R3(KR)_사례번호_제품명.xml
            product = re.sub(r'[\\/:*?"<>|\s]+', '', case['product'])[:40]
            name = f"ICSR_R3({CONFIG['country']})_{case['case_id']}" + (f'_{product}' if product else '') + '.xml'
            (out_dir / name).write_bytes(xml)
            status, xsd_error = '검토 필요' if issues else '완료', ''
            if schema is not None:
                from lxml import etree
                if not schema.validate(etree.fromstring(xml)):
                    status, xsd_error = '오류', str(schema.error_log.last_error)
            cases.append(dict(row=case['row'], case_id=case['case_id'], file=name, product=case['product'],
                              terms=case['terms'], relatedness=case['relatedness'], status=status,
                              issues=issues, xsd_error=xsd_error))
            LOG.info('%s · %s · %s', name, status, '; '.join(issues) or '-')
    summary = dict(version=VERSION, source=workbook['name'], source_sha256=workbook.get('source_hash', ''),
                   created_at=now.isoformat(timespec='seconds'), schema=schema_note or '검증함', output=str(out_dir),
                   ae_cases=len(cases), skipped_rows=skipped,
                   meddra_version=meddra.get('meddra_version', ''), config=CONFIG, cases=cases)
    (out_dir / 'AAI_결과.json').write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding='utf-8')
    (out_dir / 'AAI_결과.html').write_text(report(summary), encoding='utf-8')
    return summary


def report(s):
    esc = lambda x: html.escape(str(x))
    counts = {k: sum(c['status'] == k for c in s['cases']) for k in ('완료', '검토 필요', '오류')}
    rows = ''.join(
        f"<tr class='{ {'완료': 'ok', '검토 필요': 'warn', '오류': 'err'}[c['status']] }'><td>{c['row']}</td>"
        f"<td><a href='{esc(c['file'])}'>{esc(c['case_id'])}</a></td><td>{esc(c['product'])}</td>"
        f"<td>{esc(', '.join(c['terms']))}</td><td>{esc(c['relatedness'] or '—')}</td><td>{c['status']}</td>"
        f"<td>{'<br>'.join(esc(i) for i in c['issues'] + ([c['xsd_error']] if c['xsd_error'] else [])) or '—'}</td></tr>"
        for c in s['cases'])
    note = (f"<p>발송자 {esc(s['config']['sender_id'])} → 수신자 {esc(s['config']['receiver_id'])} · "
            f"저장 폴더: <code>{esc(s['output'])}</code></p>")
    return f"""<!doctype html><html lang="ko"><head><meta charset="utf-8"><title>AAI {s['version']} · {esc(s['source'])}</title>
<style>body{{font:14px "Malgun Gothic",sans-serif;margin:24px;color:#172132;background:#f5f6f8}}h1{{font-size:22px}}
table{{border-collapse:collapse;width:100%;background:#fff}}th,td{{border:1px solid #d8dfe7;padding:8px;text-align:left;vertical-align:top}}
th{{background:#eef2f7}}tr.ok td:nth-child(6){{color:#1f7a3a}}tr.warn td:nth-child(6){{color:#9a6400}}tr.err td:nth-child(6){{color:#ae2020;font-weight:700}}
.stats span{{display:inline-block;background:#fff;border:1px solid #d8dfe7;border-radius:6px;padding:8px 14px;margin:0 8px 12px 0}}
p.warn{{background:#fff9e8;border-left:3px solid #d0a039;padding:10px 14px}}code{{background:#eef2f7;padding:1px 4px}}</style></head><body>
<h1>AAI {s['version']} · E2B(R3) XML 생성 결과</h1>
<p>원본: {esc(s['source'])} · 생성: {esc(s['created_at'])} · XSD: {esc(s['schema'])} · MedDRA 버전: {esc(s['meddra_version'] or '미설정')}</p>
<div class="stats"><span>AE 사례 {s['ae_cases']}</span><span>완료 {counts['완료']}</span><span>검토 필요 {counts['검토 필요']}</span><span>오류 {counts['오류']}</span><span>건너뛴 행(H열 O 아님) {s['skipped_rows']}</span></div>
{note}<p>‘검토 필요’ XML은 구조는 맞지만 MedDRA 코드 등 담당자가 채울 항목이 있습니다. ‘오류’ XML은 안전성 DB로 가져오지 마세요.</p>
<table><thead><tr><th>행</th><th>사례번호 (XML)</th><th>제품</th><th>이상사례</th><th>인과관계</th><th>상태</th><th>확인 사항</th></tr></thead>
<tbody>{rows or '<tr><td colspan=7>H열이 O인 행이 없습니다.</td></tr>'}</tbody></table></body></html>"""


def workspace(source):
    """ACI와 같은 작업 폴더(aci_settings.json). 없으면 입력 파일 폴더."""
    settings = ROOT / 'aci_settings.json'
    if settings.exists():
        folder = Path(json.loads(settings.read_text(encoding='utf-8')).get('workspace', ''))
        if folder.is_dir():
            return folder
    return source.parent


def main():
    parser = argparse.ArgumentParser(description='ACI 결과 엑셀 → E2B(R3) XML')
    parser.add_argument('input', help='ACI 결과 엑셀 (IDBI_YYYYMMDD_ACI.xls)')
    parser.add_argument('--out', help='저장 폴더 (기본: ACI 작업 폴더의 aai_outputs/파일명_시각)')
    parser.add_argument('--no-open', action='store_true', help='결과 보고서를 브라우저로 열지 않음')
    args = parser.parse_args()
    source = Path(args.input).resolve()
    LOG.info('AAI %s 시작: %s', VERSION, source.name)
    try:
        meddra = json.loads((ROOT / 'aai' / 'meddra_map.json').read_text(encoding='utf-8'))
        workbook = read_workbook(source)
        stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        out_dir = Path(args.out) if args.out else workspace(source) / CONFIG['output_folder'] / f'{source.stem}_{stamp}'
        summary = convert(workbook, out_dir, meddra)
    except Exception:
        LOG.exception('AAI 실패: %s', source.name)
        raise
    counts = {k: sum(c['status'] == k for c in summary['cases']) for k in ('완료', '검토 필요', '오류')}
    LOG.info('AAI 완료: AE %d건 (완료 %d / 검토 필요 %d / 오류 %d), 건너뜀 %d행 → %s', summary['ae_cases'],
             counts['완료'], counts['검토 필요'], counts['오류'], summary['skipped_rows'], out_dir)
    print(json.dumps(dict(output=str(out_dir), ae_cases=summary['ae_cases'], skipped_rows=summary['skipped_rows'],
                          **counts), ensure_ascii=False))
    if not args.no_open:
        webbrowser.open((out_dir / 'AAI_결과.html').as_uri())


if __name__ == '__main__':
    main()
