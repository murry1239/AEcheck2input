import json
from pathlib import Path
import openpyxl
from openpyxl.cell.rich_text import CellRichText

ROOT=Path(__file__).resolve().parents[1]
source=json.loads((ROOT/'work/native_input.json').read_text(encoding='utf-8'))
expected=json.loads((ROOT/'work/expected_records.json').read_text(encoding='utf-8'))
book=openpyxl.load_workbook(ROOT/'work/ACI_final_preview.xlsx',rich_text=True)
assert book.sheetnames==[s['name'] for s in source['sheets']]
for sheet in source['sheets']:
    s=book[sheet['name']]
    assert s.freeze_panes=='A2'
    assert [s.cell(1,c).value for c in range(8,14)]==['AE','Cube','SAE','relatedness','QC','comment']
    for row in sheet['records']:
        for c,v in enumerate(row['values'],1):
            actual=s.cell(row['row'],c).value
            if isinstance(actual,CellRichText):actual=str(actual)
            assert ('' if actual is None else actual)==v,(row['row'],c)
for row in expected:
    s=book[row['sheet']];i=row['row'];g=s.cell(i,7);h=s.cell(i,8);k=s.cell(i,11)
    assert h.value==('O — '+row['summary'] if row['ae']=='O' else 'X')
    assert (k.value or '')==row['relatedness']
    assert g.alignment.wrap_text
    assert all(s.cell(i,c).value is None for c in [9,10,12,13])
    if row['ae']=='O':
        assert g.fill.fgColor.rgb[-6:]=='FFFF00'
        assert isinstance(g.value,CellRichText)
        all_text=str(g.value);start=0;bold=set()
        for portion in g.value:
            text=str(portion)
            if getattr(getattr(portion,'font',None),'b',False):bold.update(range(start,start+len(text)))
            start+=len(text)
        wanted=set()
        for snippet in row['snippets']:
            pos=0
            while (pos:=all_text.find(snippet,pos))>=0:
                wanted.update(range(pos,pos+len(snippet)));pos+=len(snippet)
        assert bold==wanted,('partial bold scope',i)
print(json.dumps({'rows':len(expected),'ae_o':sum(r['ae']=='O' for r in expected),
    'source_preserved':True,'rich_text_exact':True,'freeze':'A2','G_width':book.active.column_dimensions['G'].width,
    'capped_rows':[i for i,d in book.active.row_dimensions.items() if d.height and d.height>=409.5]}))
