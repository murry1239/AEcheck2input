"""Native XLS read/write through an isolated, installed LibreOffice instance."""
import argparse
import hashlib
import json
import os
import subprocess
import sys
import time
import uuid
from pathlib import Path

import uno

from engine import fingerprint, spans_for

ROOT = Path(__file__).resolve().parents[1]
HEADERS = ['구분', '일자', '성별', '나이', '품목명', '문의종류', '불만사항']
EXTRA = ['AE', 'Cube', 'SAE', 'relatedness', 'QC', 'comment']


def prop(name, value):
    p = uno.createUnoStruct('com.sun.star.beans.PropertyValue')
    p.Name, p.Value = name, value
    return p


def filehash(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


class Office:
    def __enter__(self):
        self.doc = None
        self.desktop = None
        self.profile = ROOT / '.runtime' / ('lo_' + uuid.uuid4().hex)
        self.profile.mkdir(parents=True)
        name = 'aci_' + uuid.uuid4().hex
        binary = Path(sys.executable).parent / 'soffice.com'
        if not binary.exists():
            binary = binary.with_suffix('.exe')
        command = [str(binary), '-env:UserInstallation=' + self.profile.as_uri(),
                   '--headless', '--nologo', '--nodefault', '--nofirststartwizard',
                   '--accept=pipe,name=' + name + ';urp;StarOffice.ComponentContext']
        flags = subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
        self.process = subprocess.Popen(command, stdout=subprocess.DEVNULL,
                                        stderr=subprocess.DEVNULL, creationflags=flags)
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
        deadline = time.monotonic() + 45
        while time.monotonic() < deadline:
            try:
                ctx = resolver.resolve('uno:pipe,name=' + name + ';urp;StarOffice.ComponentContext')
                self.desktop = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
                return self
            except Exception:
                time.sleep(.2)
        self.__exit__(None, None, None)
        raise RuntimeError('LibreOffice 작업 인스턴스 연결 시간 초과. 실행 환경을 확인하세요.')

    def load(self, path, readonly=True):
        if self.doc:
            self.doc.close(True)
        self.doc = self.desktop.loadComponentFromURL(Path(path).resolve().as_uri(), '_blank', 0,
            (prop('Hidden', True), prop('ReadOnly', readonly), prop('MacroExecutionMode', uno.getConstantByName('com.sun.star.document.MacroExecMode.NEVER_EXECUTE')), prop('UpdateDocMode', 0)))
        if not self.doc or not self.doc.supportsService('com.sun.star.sheet.SpreadsheetDocument'):
            raise ValueError('지원하는 엑셀 통합문서를 열 수 없습니다.')
        # A hidden Calc frame has no viewport height; set its virtual size so
        # freezing row 1 is not clamped to zero by the view controller.
        self.doc.CurrentController.Frame.ContainerWindow.setPosSize(0, 0, 1280, 800,
            uno.getConstantByName('com.sun.star.awt.PosSize.SIZE'))
        return self.doc

    def __exit__(self, *args):
        try:
            if self.doc:
                self.doc.close(True)
        finally:
            try:
                if self.desktop:
                    self.desktop.terminate()
            finally:
                if getattr(self, 'process', None):
                    try:
                        self.process.wait(timeout=8)
                    except subprocess.TimeoutExpired:
                        self.process.terminate()  # Only the private process started above.


def used_size(sheet):
    cursor = sheet.createCursor()
    cursor.gotoEndOfUsedArea(False)
    a = cursor.RangeAddress
    return a.EndRow + 1, a.EndColumn + 1


def snapshot(doc):
    result = []
    for name in doc.Sheets.ElementNames:
        s = doc.Sheets.getByName(name)
        nrows, ncols = used_size(s)
        if nrows > 65536 or ncols > 256:
            raise ValueError('XLS 행/열 제한을 초과하는 파일입니다.')
        area = s.getCellRangeByPosition(0, 0, ncols - 1, nrows - 1)
        result.append(dict(name=name, rows=nrows, columns=ncols,
                           values=area.getDataArray(), formulas=area.getFormulaArray()))
    return result


def describe(path, doc):
    sheets = []
    for name in doc.Sheets.ElementNames:
        s = doc.Sheets.getByName(name)
        nrows, ncols = used_size(s)
        headers = list(s.getCellRangeByPosition(0, 0, 6, 0).getDataArray()[0])
        if [str(x).replace(' ', '').strip() for x in headers] != HEADERS:
            continue
        records = []
        for row in range(1, nrows):
            values = list(s.getCellRangeByPosition(0, row, 6, row).getDataArray()[0])
            if not any(v != '' for v in values):
                continue
            records.append(dict(row=row + 1, values=values))
        sheets.append(dict(name=name, records=records))
    if not sheets:
        raise ValueError('첫 행 A~G 제목이 지정된 IDBI 양식과 일치하는 시트가 없습니다.')
    return dict(path=str(Path(path).resolve()), name=Path(path).name,
                source_hash=filehash(path), sheets=sheets)


def bold_spans(cell):
    result = []
    pos = 0
    paragraphs = cell.Text.createEnumeration()
    while paragraphs.hasMoreElements():
        paragraph = paragraphs.nextElement()
        portions = paragraph.createEnumeration()
        while portions.hasMoreElements():
            portion = portions.nextElement()
            text = portion.String
            if portion.CharWeight >= 150:
                result.append([pos, pos + len(text)])
            pos += len(text)
        pos += 1  # Paragraph separator is one character in Calc cell strings.
    return result


def export(path, target, payload, office):
    if filehash(path) != payload['source_hash']:
        raise ValueError('검토 이후 원본이 변경되었습니다. 파일을 다시 불러오세요.')
    target = Path(target).resolve()
    if target == Path(path).resolve() or target.exists():
        raise ValueError('원본 또는 기존 결과를 덮어쓸 수 없습니다.')
    doc = office.load(path, readonly=False)
    before = snapshot(doc)
    bysheet = {}
    for record in payload['records']:
        bysheet.setdefault(record['sheet'], []).append(record)
    for name, records in bysheet.items():
        sheet = doc.Sheets.getByName(name)
        for c, title in enumerate(EXTRA, 7):
            cell = sheet.getCellByPosition(c, 0)
            if cell.String and cell.String != title:
                raise ValueError('H~M 영역에 다른 제목이 있습니다. 기존 열을 덮어쓰지 않습니다.')
            cell.String = title
            cell.CellStyle = sheet.getCellByPosition(6, 0).CellStyle
            cell.CharWeight = 150
            cell.CharHeight = sheet.getCellByPosition(6, 0).CharHeight
            cell.CellBackColor = sheet.getCellByPosition(6, 0).CellBackColor
            cell.CharColor = sheet.getCellByPosition(6, 0).CharColor
        sheet.Columns.getByIndex(6).Width = 16500
        sheet.Columns.getByIndex(7).Width = 5800
        sheet.Columns.getByIndex(10).Width = 3400
        for c in [8, 9, 11, 12]:
            if sheet.Columns.getByIndex(c).Width < 2200:
                sheet.Columns.getByIndex(c).Width = 2200
        for r in records:
            row = r['row'] - 1
            g = sheet.getCellByPosition(6, row)
            values = list(sheet.getCellRangeByPosition(0, row, 6, row).getDataArray()[0])
            if fingerprint(values) != r['key']:
                raise ValueError('원본 행 내용과 검토 데이터가 다릅니다.')
            g.IsTextWrapped = True
            g.VertJustify = 1  # TOP
            if r['ae'] == 'O':
                g.CellBackColor = 0xFFFF00
                for start, end in spans_for(g.String, r['snippets']):
                    cursor = g.Text.createTextCursor()
                    cursor.gotoStart(False)
                    cursor.goRight(start, False)
                    cursor.goRight(end - start, True)
                    cursor.CharWeight = 150
            h = sheet.getCellByPosition(7, row)
            h.String = 'O — ' + r['summary'].strip() if r['ae'] == 'O' else 'X'
            h.IsTextWrapped = True
            h.VertJustify = 1
            k = sheet.getCellByPosition(10, row)
            k.String = r.get('relatedness', '')
            k.IsTextWrapped = True
            k.VertJustify = 1
            sheet.Rows.getByIndex(row).OptimalHeight = True
        doc.CurrentController.setActiveSheet(sheet)
        doc.CurrentController.freezeAtPosition(0, 1)
    doc.calculateAll()
    target.parent.mkdir(parents=True, exist_ok=True)
    filter_name = 'MS Excel 97' if target.suffix.lower() == '.xls' else 'Calc MS Excel 2007 XML'
    doc.storeAsURL(target.as_uri(), (prop('FilterName', filter_name), prop('Overwrite', False)))
    doc = office.load(target)
    after = snapshot(doc)
    if [s['name'] for s in before] != [s['name'] for s in after]:
        raise RuntimeError('저장 후 시트 구성이 바뀌었습니다.')
    for b, a in zip(before, after):
        modified_rows = {r['row'] - 1 for r in bysheet.get(b['name'], [])}
        for row in range(b['rows']):
            for col in range(b['columns']):
                changed = b['name'] in bysheet and ((row == 0 and 7 <= col <= 12) or
                                                    (row in modified_rows and col in (7, 10)))
                if changed:
                    continue
                if b['values'][row][col] != a['values'][row][col] or b['formulas'][row][col] != a['formulas'][row][col]:
                    raise RuntimeError('원본 값/수식 보존 검증 실패: ' + b['name'] + ' ' + str((row + 1, col + 1)))
    details = []
    for name, records in bysheet.items():
        s = doc.Sheets.getByName(name)
        if [s.getCellByPosition(c, 0).String for c in range(7, 13)] != EXTRA:
            raise RuntimeError('추가 제목 저장 검증 실패')
        doc.CurrentController.setActiveSheet(s)
        if not doc.CurrentController.hasFrozenPanes():
            raise RuntimeError('첫 행 고정 저장 검증 실패')
        for r in records:
            row = r['row'] - 1
            g, h, k = (s.getCellByPosition(c, row) for c in [6, 7, 10])
            expected_h = 'O — ' + r['summary'].strip() if r['ae'] == 'O' else 'X'
            if h.String != expected_h or k.String != r.get('relatedness', '') or not g.IsTextWrapped:
                raise RuntimeError('AE/인과관계/줄 바꿈 저장 검증 실패: ' + str(r['row']))
            spans = bold_spans(g)
            if r['ae'] == 'O':
                if g.CellBackColor != 0xFFFF00:
                    raise RuntimeError('AE 노란색 저장 검증 실패')
                for start, end in spans_for(g.String, r['snippets']):
                    if not any(a <= start and end <= b for a, b in spans):
                        # Adjacent formatted portions may be split by paragraph boundaries.
                        if any(not any(a <= i < b for a, b in spans) for i in range(start, end) if g.String[i] != '\n'):
                            raise RuntimeError('AE 부분 굵게 저장 검증 실패: ' + str(r['row']))
            details.append(dict(sheet=name, row=r['row'], h=h.String, k=k.String, bold_spans=spans))
    if filehash(path) != payload['source_hash']:
        raise RuntimeError('원본 해시 보존 검증 실패')
    return dict(output=str(target), verified=True, rows=len(details), details=details)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('mode', choices=['read', 'export'])
    parser.add_argument('path')
    parser.add_argument('--payload')
    parser.add_argument('--output')
    parser.add_argument('--json-out', required=True)
    args = parser.parse_args()
    try:
        with Office() as office:
            if args.mode == 'read':
                result = describe(args.path, office.load(args.path))
            else:
                payload = json.loads(Path(args.payload).read_text(encoding='utf-8'))
                result = export(args.path, args.output, payload, office)
        Path(args.json_out).write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    except Exception as e:
        Path(args.json_out).write_text(json.dumps({'error': str(e)}, ensure_ascii=False), encoding='utf-8')
        raise


if __name__ == '__main__':
    main()
