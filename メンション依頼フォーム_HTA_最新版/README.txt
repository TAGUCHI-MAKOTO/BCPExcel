メンション依頼フォーム HTA 最新版 v28.16 Layout Fix
============================================================

■ この版について

v28.15 本体一式をベースに、
MentionRequestForm_v28.16_Layout_Fix.zip のレイアウトCSSを適用した最新版です。

■ v28.16 Layout Fix

・メールメモ／処理日時／期日または面接日程を同じ幅・間隔に統一
・追加／取り消し／クリアボタンを、上の入力欄と同じ列に配置
・期日の入力欄、タイプのリスト、クリアボタン、送信ボタンの右端を統一
・本体幅を1480pxに調整し、右側の余白を圧縮

■ v28.15までの主な仕様

・「代理CA組織」入力欄を廃止
  ※CSV互換のため列自体は空欄で維持
・「代理CA名」は任意入力
・各依頼カード左下に入力時の注意事項を表示
  「組織・CA名はCanvasからコピーしてください」
・依頼2／3で「同一CA」にチェックした場合は注意事項を非表示
・拠点／タイプの選択肢マスタは mention-form.js に集約
・送信ボタン押下後、Pending CSV保存 → Worker起動 → 共有CSVへ即時書き込み
・受付完了はメインHTA内モーダルで表示
・受付OKでフォームを最小化
・フォームを閉じてもWorkerは送信処理を継続
・共有CSVの書き込み／再読込確認成功時のみPending CSVを削除
・送信完了はWindows標準Popupで通知
・完了Popupは前面＋最前面表示、OKを押すまで表示

■ 安全性／送信処理

・Pure JScript Worker
・PowerShellなし
・.ps1なし
・ログなし
・PendingはCSVのみ
・.jlock排他制御
・RequestID＋依頼番号で重複防止
・最大300秒再試行
・書き込み後再読込確認
・成功時のみPending削除

■ ファイル構成

メンション依頼フォーム_HTA_最新版/
├─ README.txt
├─ mention-form.css
├─ mention-form.js
├─ mention-request-worker.js
├─ メンション依頼フォーム.hta
└─ 書き込み用/
   └─ このフォルダにCSVが作成されます.txt

■ 確認状況

・v28.15までの送信処理は従来仕様を維持
・v28.16 Layout Fix はCSS差し替えのみ
・Layout Fix作成時点ではWindows HTA実機での表示確認は未実施
