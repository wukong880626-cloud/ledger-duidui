from pathlib import Path
import hashlib
import zipfile
import json
import re

root = Path(__file__).resolve().parent.parent
release = root / 'release'
release.mkdir(exist_ok=True)
version=json.loads((root/'package.json').read_text())['version']
name = f'台账对对看_v{version}'
target = release / f'{name}.zip'
entries = ['dist', 'src', 'samples', 'test', 'scripts', 'output', 'README.md', '测试结果.md', '第二轮验收_v0.1.1.md', '更新说明_v0.1.1.md', 'package.json', 'package-lock.json', '.gitignore', '打开台账对对看.html', '启动台账对对看.command']
files = []
for entry in entries:
    path = root / entry
    if not path.exists():
        raise RuntimeError(f'缺少交付文件：{entry}')
    candidates=[p for p in path.rglob('*') if p.is_file() and p.name != '.DS_Store'] if path.is_dir() else [path]
    for p in candidates:
        rel=p.relative_to(root).as_posix()
        if rel.startswith('output/round2/fixtures/') or re.search(r'_(?:csv|xlsx|xlsxnum)_\d+\.xlsx$',rel):
            continue
        files.append(p)
with zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as archive:
    for path in sorted(files):
        archive.write(path, Path(name) / path.relative_to(root))
with zipfile.ZipFile(target) as archive:
    bad = archive.testzip()
    if bad:
        raise RuntimeError(f'ZIP CRC 校验失败：{bad}')
digest = hashlib.sha256(target.read_bytes()).hexdigest()
(release / 'SHA256SUMS.txt').write_text(f'{digest}  {target.name}\n', encoding='utf-8')
print(f'{target}\n{len(files)} files / {target.stat().st_size:,} bytes\nSHA256 {digest}')
