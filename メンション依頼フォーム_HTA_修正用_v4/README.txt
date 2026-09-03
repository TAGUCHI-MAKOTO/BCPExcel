メンション依頼フォーム_HTA 修正用 v4
=====================================

本番環境で「mention-form.config.ini が見つかりません」と表示される事象の修正版です。

■ 修正内容
HTA自体をUNC共有フォルダから起動した場合、
window.location.pathnameだけでは欠落するサーバー名を
window.location.hostnameと組み合わせて復元します。

例：
\\server\share\HTA\メンション依頼フォーム.hta

■ 診断強化
iniが見つからない場合は、
実際に探しに行ったフルパスを「確認先」として表示します。

■ 構成
本番候補v3のファイルをベースに、
unc-fix.jsを追加し、HTAから最後に読み込む修正確認用構成です。

修正確認後、問題なければ最新版のmention-form.jsへ統合します。
