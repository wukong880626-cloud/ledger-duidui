from pathlib import Path
import json
import os
import re
import select
import shutil
import subprocess
import tempfile
import urllib.request
import zipfile

root = Path(__file__).resolve().parent.parent
version=json.loads((root/'package.json').read_text())['version']
name=f'台账对对看_v{version}'
archive_path = root / 'release' / f'{name}.zip'
with zipfile.ZipFile(archive_path) as archive, tempfile.TemporaryDirectory(prefix='ledger-package-check-') as temp:
    assert archive.testzip() is None
    archive.extractall(temp)
    project = Path(temp) / name
    assert not (project / 'node_modules').exists()
    assert (project / 'dist/app.js').read_bytes() == (root / 'dist/app.js').read_bytes()
    assert (project / 'dist/worker.js').read_bytes() == (root / 'dist/worker.js').read_bytes()
    node = shutil.which('node')
    assert node, '需要 Node.js 验证启动'
    process = subprocess.Popen([node, 'scripts/serve.mjs'], cwd=project, env={**os.environ, 'PORT': '0'}, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
    try:
        assert select.select([process.stdout], [], [], 10)[0], '启动超时'
        line = process.stdout.readline()
        match = re.search(r'http://127\.0\.0\.1:\d+', line)
        assert match, line
        base = match.group()
        sizes = {}
        for path in ['/', '/style.css', '/app.js', '/worker.js']:
            with urllib.request.urlopen(base + path, timeout=5) as response:
                assert response.status == 200
                sizes[path] = len(response.read())
        assert sizes['/app.js'] > 100000
        evidence = {'version':version,'zipCRC': 'passed', 'standaloneWithoutNodeModules': True, 'staticFilesHTTP200': sizes, 'bundleMatchesBuild': True}
        (root / 'output/package-validation.json').write_text(json.dumps(evidence, ensure_ascii=False, indent=2), encoding='utf-8')
        print(json.dumps(evidence, ensure_ascii=False))
    finally:
        process.terminate()
        process.wait(timeout=5)
