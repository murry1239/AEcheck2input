import json
from pathlib import Path
from urllib.error import HTTPError
from urllib.request import Request,urlopen

root=Path(__file__).resolve().parents[1]
c=json.loads((root/'work/connection.json').read_text(encoding='utf-8'))
base='http://127.0.0.1:'+str(c['port'])
def request(path,body=None,auth=True):
    h={'X-ACI-Token':c['token']} if auth else {}
    if body is not None:h['Content-Type']='application/json'
    r=Request(base+path,data=json.dumps(body).encode() if body is not None else None,headers=h)
    try:
        with urlopen(r,timeout=10) as response:return response.status,json.load(response)
    except HTTPError as e:return e.code,json.load(e)
assert request('/api/files',auth=False)[0]==403
assert request('/api/load?name=../../IDBI_20261001.xls')[0]==400
original=request('/api/files')[1]['workspace']
test=root/'tests/workspace_check';test.mkdir(exist_ok=True)
try:
    status,data=request('/api/workspace',{'workspace':str(test)})
    assert status==200 and data['workspace']==str(test) and not data['files']
    assert request('/api/files')[1]['workspace']==str(test)
    assert request('/api/workspace',{'workspace':str(test/'does_not_exist')})[0]==400
finally:
    assert request('/api/workspace',{'workspace':original})[0]==200
assert 'IDBI_20261001.xls' in request('/api/files')[1]['files']
print('HTTP: unauthorized access, traversal, workspace selection/persistence, invalid directory and restore verified')
