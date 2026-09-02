メンション依頼フォーム｜本番実装候補 v1
========================================

■ 本番化した3本柱

① ローカル退避（データ消失対策）
送信ボタン押下後、最初にDesktop\MentionRequest_Pendingへ
可読形式の.pendingを保存します。

Pending保存後にだけBackground Workerを起動するため、
共有フォルダ障害やCSV競合が発生しても元データを保持します。

② Background Worker（処理中断対策）
CSV書き込みはHTAとは別のPowerShell workerで実行します。

受付完了後はHTAを×で閉じても処理継続。
成功後にWindows右下通知を表示します。

HTA次回起動時には残存Pendingを自動検出して再送します。

③ lock＋ランダム再試行（同時送信対策）
CSVごとに.lockを排他取得してから処理します。

lock方式はCreateNewではなく
OpenOrCreate + FileShare.None
へ変更しました。

このため、workerが異常終了して.lockファイル自体が残っても、
Windowsのhandleが解放されていれば再度lock取得可能です。
「15秒経過したlockを強制削除」のような危険な処理はありません。

再試行は450～1250msのランダム間隔です。


■ RequestID重複・途中書き込み対策

RequestIDだけで「送信済み」と判断せず、
RequestID＋依頼番号（依頼1/2/3）単位で確認します。

例：
依頼1・2・3を送信中にPC/workerが停止
↓
CSVには依頼1だけ反映
↓
次回Pending自動再送
↓
依頼1は既存としてスキップ
依頼2・3だけ追記
↓
3件すべて揃ったことを確認
↓
Pending削除

これにより、
「一部だけ書けた状態でRequestIDを重複扱いして残りが消える」
リスクを抑えています。


■ CSV書き込み確認

同じlockを保持したまま、
書き込み後にRequestID＋依頼番号を再読込し、
Pending内の全依頼番号がCSVに存在することを確認します。

不足があればPendingを削除せず再試行します。


■ CSVファイルの日付

復旧が翌日になっても、
Pendingを作成した日付のCSVへ書き込みます。

例：
9/3 23:59に送信 → 9/4に復旧
→ 呉服_20260903.csv


■ 送信先設定

mention-form.config.ini
のCSV_FOLDERへ本番共有フォルダを設定してください。

推奨：
\\サーバー名\共有フォルダ\メンション依頼

UNCパスを推奨します。

V3 Pendingには送信時点のCSV_FOLDERも保存するため、
送信後にconfig設定が変更されても、
既存Pendingは元の送信先へ復旧します。


■ Windows通知

通常成功：
メンション依頼｜送信完了

既に全件送信済み：
メンション依頼｜送信済み確認

途中書き込みから復旧：
メンション依頼｜送信完了
「途中まで反映済みの依頼を補完して…」

長時間待機：
メンション依頼｜自動再送中


■ ローカル診断ログ

%LOCALAPPDATA%\MentionRequest\Logs\worker_YYYYMMDD.log

START / RETRY / SUCCESS / DUPLICATE_SKIP /
RECOVERED_PARTIAL / FATAL 等を記録します。

ログ失敗は送信処理へ影響しません。


■ 本番前に残っている確認

・mention-form.config.iniへ実際の共有フォルダを設定
・職場PC 2台以上＋同じ共有フォルダで負荷テスト
・PowerShell実行可否
・Windows通知表示
・共有フォルダの作成/追記権限

1台PCでは250workerの競合負荷試験までクリア済みの設計を
本番フォーム側へ反映したバージョンです。


■ ファイル構成

README.txt
mention-form.config.ini
mention-form.css
mention-form.js
mention-request-worker.ps1
メンション依頼フォーム.hta


[v2 設定パス判定修正]
--------------------
v1ではCSV_FOLDERの値に「【」が含まれているだけで
未設定と誤判定する不具合がありました。

例：
C:\Users\tagut\OneDrive\デスクトップ\【拡張機能】\HTA\書き込み用

上記のような正常なパスも未設定扱いになっていました。

v2では、
CSV_FOLDER=【ここに共有フォルダのUNCパスを設定】
という初期プレースホルダーの場合だけ未設定扱いにします。

日本語・【】を含むフォルダ名は正常に使用できます。


[v3 送信完了メッセージ変更]
--------------------------
フォーム上の受付メッセージを以下へ変更しました。

タイトル：
送信完了

本文：
メンション依頼を送信しました。

受け付け完了後、右下のWindows通知でお知らせいたします。

Background Worker / Pending / lock / 再試行 / 重複防止などの処理仕様は変更ありません。
