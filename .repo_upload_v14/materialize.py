from pathlib import Path
import re

root = Path('メンション依頼フォーム_HTA_最新版')
js_path = root / 'mention-form.js'
readme_path = root / 'README.txt'

js = js_path.read_text(encoding='utf-8-sig')
readme = readme_path.read_text(encoding='utf-8-sig')

pattern = re.compile(r'var PRODUCTION_CSV_FOLDER\s*=\s*"[^"]*";')
m = pattern.search(js)
if not m:
    raise SystemExit('PRODUCTION_CSV_FOLDER definition not found')
js = js[:m.start()] + 'var CSV_SUBFOLDER_NAME = "書き込み用";' + js[m.end():]

start = js.find('function getProductionCsvFolder(required){')
if start < 0:
    raise SystemExit('getProductionCsvFolder not found')
end = js.find('\nfunction ', start + 10)
if end < 0:
    raise SystemExit('end of getProductionCsvFolder not found')

new_func = r'''function getProductionCsvFolder(required){
    var fso=new ActiveXObject("Scripting.FileSystemObject");
    var baseFolder=getCurrentFolderPath();
    var csvFolder=fso.BuildPath(baseFolder,CSV_SUBFOLDER_NAME);

    if(!fso.FolderExists(csvFolder)){
        if(required){
            throw new Error(
                "CSV保存先フォルダが見つかりません。\n\n"+
                "メンション依頼フォームと同じ場所に「"+CSV_SUBFOLDER_NAME+"」フォルダを作成してください。\n\n"+
                "確認先：\n"+csvFolder
            );
        }
        return "";
    }

    return csvFolder;
}
'''
js = js[:start] + new_func + js[end:]
js = js.replace(
    'setStatus("本番CSV保存先が未設定です");',
    'setStatus("書き込み用フォルダが見つかりません");'
)

readme += r'''

[v14 CSV保存先：HTAと同じ階層の「書き込み用」フォルダ]
-----------------------------------------------------
CSV保存先の固定パス設定を廃止しました。

■ 配置例
メンション依頼フォーム/
├─ メンション依頼フォーム.hta
├─ mention-form.js
├─ mention-form.css
├─ mention-request-worker.ps1
└─ 書き込み用/
   └─ 拠点_YYYYMMDD.csv

■ 動作
・HTA自身の配置場所を取得
・同じ階層の「書き込み用」フォルダをCSV保存先として使用
・PCごとの C:\Users\... 設定は不要
・mention-form.config.ini も不要
・「書き込み用」フォルダが存在しない場合は送信せずエラー表示

■ 注意
共有ディスク上で使用する場合、
「書き込み用」フォルダに利用者の書き込み権限が必要です。

Pendingは従来どおり各PCのデスクトップへ退避します。
Background Worker / 300秒再送 / CSV lock / 重複防止仕様は維持しています。
'''

js_path.write_text(js, encoding='utf-8-sig')
readme_path.write_text(readme, encoding='utf-8-sig')

csv_dir = root / '書き込み用'
csv_dir.mkdir(exist_ok=True)
(csv_dir / 'このフォルダにCSVが作成されます.txt').write_text(
    'このフォルダはメンション依頼フォームのCSV保存先です。\n'
    '削除・名称変更しないでください。\n',
    encoding='utf-8-sig'
)
