from pathlib import Path
import base64, hashlib, shutil, zipfile

root = Path('.')
staging = root / '.repo_upload_v21'
parts = sorted(staging.glob('mention_v21.b64.part*'))
if len(parts) != 3:
    raise SystemExit(f'expected 3 upload parts, found {len(parts)}')

encoded = ''.join(p.read_text(encoding='utf-8') for p in parts)
archive_bytes = base64.b64decode(encoded.encode('ascii'))
digest = hashlib.sha256(archive_bytes).hexdigest()
expected = '3f08602b4e497bf87c878eb29454be22d70abe7cd7605342e8c34db12cacecca'
if digest != expected:
    raise SystemExit(f'SHA256 mismatch: {digest}')

archive = root / '.mention_v21.zip'
archive.write_bytes(archive_bytes)

latest = root / 'メンション依頼フォーム_HTA_最新版'
if latest.exists():
    shutil.rmtree(latest)
latest.mkdir(parents=True)

with zipfile.ZipFile(archive, 'r') as z:
    z.extractall(latest)

archive.unlink(missing_ok=True)

# 一時アップロード素材を削除
for p in staging.glob('*'):
    if p.is_file():
        p.unlink()
try:
    staging.rmdir()
except OSError:
    pass
