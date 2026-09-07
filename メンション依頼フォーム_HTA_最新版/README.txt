メンション依頼フォーム HTA 本番対策版 v28.3
============================================================

■ v28.3 修正内容

v28.2で書き込み完了後のポップアップが表示されなかったため、
Worker起動時の wscript.exe オプションを修正しました。

旧：
wscript.exe //B //NoLogo //E:JScript

新：
wscript.exe //NoLogo //E:JScript

//B はバッチモードで、画面表示を抑制する挙動があるため、
完了ポップアップとの相性を考えて廃止しています。

■ 完了通知

次のすべてが成功した場合だけ表示します。

1. 共有CSVへの作成／追記
2. 書き込み後の再読込確認
3. 退避CSV削除

表示：
「メンション依頼の送信が完了しました。」
＋ 拠点名

約5秒で自動的に閉じます。

フォームを閉じたあとでも、
別プロセスのWorkerから表示される設計です。

■ 維持している仕様

・Pure JScript Worker
・Workerは完全ASCII／BOMなし
・PowerShellなし
・.ps1なし
・EncodedCommandなし
・ExecutionPolicy Bypassなし
・ログなし
・退避形式はCSVのみ
・フォームを閉じてもWorker継続
・.jlockによる排他制御
・RequestID＋依頼番号の重複防止
・部分復旧
・約450～1250msのランダム再試行
・最大300秒再試行
・書き込み後確認
・成功時のみ退避CSV削除

■ フォルダ構成

メンション依頼フォーム/
├─ README.txt
├─ mention-form.css
├─ mention-form.js
├─ mention-request-worker.js
├─ メンション依頼フォーム.hta
└─ 書き込み用/
   └─ このフォルダにCSVが作成されます.txt
