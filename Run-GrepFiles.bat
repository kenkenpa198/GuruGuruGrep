:: 標準出力へコマンドを表示しない
@echo off
:: カレントディレクトリを再設定
cd /d %~dp0
:: 環境変数をローカル化
setlocal
:: 文字コードを設定
chcp 65001

echo GrepFiles.py を実行します。

python GrepFiles.py

echo 終了するには何らかのキーを押してください。
pause
