from pathlib import Path
import base64, hashlib, shutil, zipfile

root = Path('.')
staging = root / '.repo_upload'
parts = sorted(staging.glob('mention_latest_v13.b64.part*'))
if len(parts) != 6:
    raise SystemExit(f'expected 6 upload parts, found {len(parts)}')

encoded = ''.join(p.read_text(encoding='utf-8') for p in parts)
archive_bytes = base64.b64decode(encoded.encode('ascii'))
digest = hashlib.sha256(archive_bytes).hexdigest()
expected = 'b76aefce3bd0680e84c8891e8f2e4e00701e809e9e6d19b41d2a553188d2cd3d'
if digest != expected:
    raise SystemExit(f'SHA256 mismatch: {digest}')

archive = root / '.mention_v13.zip'
archive.write_bytes(archive_bytes)

latest = root / 'メンション依頼フォーム_HTA_最新版'
demo = root / 'メンション依頼フォーム_HTA_デモ版'

if latest.exists():
    shutil.rmtree(latest)
latest.mkdir(parents=True)
with zipfile.ZipFile(archive, 'r') as z:
    z.extractall(latest)

if demo.exists():
    shutil.rmtree(demo)
shutil.copytree(latest, demo)

worker = demo / 'mention-request-worker.ps1'
if worker.exists():
    worker.unlink()

hta = demo / 'メンション依頼フォーム.hta'
hta_text = hta.read_text(encoding='utf-8-sig')
hta_text = hta_text.replace(
    '<script type="text/javascript" src="mention-form.js"></script>',
    '<script type="text/javascript" src="mention-form.js"></script>\n<script type="text/javascript" src="mention-form-demo.js"></script>',
    1
)
hta_text = hta_text.replace(
    '<title>メンション依頼フォーム</title>',
    '<title>メンション依頼フォーム｜仕様書用デモ</title>',
    1
)
hta.write_text(hta_text, encoding='utf-8-sig')

demo_js = '''/* 仕様書用デモ：本番v13の画面を使い、実ファイル操作だけ無効化 */
var DEMO_PENDING_COUNT=1;
var DEMO_PENDING_STATE="idle";

function initApp(){
    populateTypes();
    setRequesterName();
    setStatus("");
    lastAvailWidth=screen.availWidth;
    lastAvailHeight=screen.availHeight;
    resizeApp(true);
    startScreenWatcher();
    bindSameCASync();
    updatePendingRetryUI();
}
function getPendingFilePaths(){return [];}
function getPendingCount(){return DEMO_PENDING_COUNT;}
function updatePendingRetryUI(){
    var area=$("pendingRetryArea"),text=$("pendingRetryText"),button=$("retryPendingButton");
    if(!area||!text||!button){return;}
    setSendButtonVisible(true);
    if(DEMO_PENDING_COUNT<=0){area.className="pending-retry-area hidden";text.innerText="";button.style.display="none";return;}
    area.className="pending-retry-area";
    if(DEMO_PENDING_STATE==="manual"){
        text.innerText="再送中です… 完了後、Windows通知でお知らせします。";
        button.style.display="none";
    }else{
        text.innerText="⚠ 送信できませんでした｜未送信 "+DEMO_PENDING_COUNT+"件";
        button.style.display="inline-block";
        button.disabled=false;
    }
}
function startPendingWatcher(){updatePendingRetryUI();}
function recoverPendingPackagesOnStartup(){updatePendingRetryUI();}
function launchAllPendingWorkers(mode){DEMO_PENDING_STATE="manual";updatePendingRetryUI();return DEMO_PENDING_COUNT;}
function launchBackgroundWorker(){return false;}
function retryPendingPackages(){
    DEMO_PENDING_STATE="manual";
    updatePendingRetryUI();
    window.setTimeout(function(){
        alert("【仕様書用デモ】\\n\\n再送ボタン押下時の表示確認です。実際のCSV書き込み・Pending再送は行いません。");
        DEMO_PENDING_STATE="idle";
        updatePendingRetryUI();
    },700);
}
function sendRequest(){
    if(!validateForm()){return;}
    alert("【仕様書用デモ】\\n\\n実際のCSV書き込み・Pending作成・Windows通知は行いません。");
}
'''
(demo / 'mention-form-demo.js').write_text(demo_js, encoding='utf-8-sig')

readme = demo / 'README.txt'
readme_text = readme.read_text(encoding='utf-8-sig')
readme_text += '\n\n[仕様書用デモ版 v13]\n本番v13レイアウトを使用しています。\nCSV書き込み・Pending作成/参照・Background Worker・Windows通知は無効です。\n起動時は「⚠ 送信できませんでした｜未送信 1件」＋「未送信を再送」＋常時表示の「送信」を確認できます。\n'
readme.write_text(readme_text, encoding='utf-8-sig')

archive.unlink(missing_ok=True)

for p in staging.glob('*'):
    if p.is_file():
        p.unlink()
try:
    staging.rmdir()
except OSError:
    pass
