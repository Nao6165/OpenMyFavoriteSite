# OpenMyFavoriteSite
曜日ごとに登録したお気に入りサイトを開くプログラムです。
WindowsPCのスタートアップに登録し、忙しい中でお気に入りサイトのチェックし忘れを防ぐために作ってみました。

動作環境：
　対象OS：Windows（Windows10にて動作確認済み）
　必要なアプリ：Excel (Microsoft Office 365 ProPlus Excelにて動作確認済み)

構成：
・MyFavoritePageList.xlsm
　⇒起動する曜日とサイトを登録します。
・OpenMyFavoritPage.bat
　⇒runExcelMacro.vbsを起動します。
・runExcelMacro.vbs
　⇒MyFavoritePageList.xlsmを起動します。

手順：
１．同一フォルダ内に、上記の3ファイルを格納してください。
２．MyFavoritePageList.xlsmを開いてください。
３．曜日欄に登録したいサイトを立ち上げる曜日を設定してください。
４．サイト欄にタイトルを記入し、リンクを張ってください。
５．MyFavoritePageList.xlsmを保存して閉じてください。
６．同フォルダ内のOpenMyFavoritPage.batを実行してください。
