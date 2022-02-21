<!-- omit in toc -->
# Grep Office

![イメージ図](images/kv.png)

指定ディレクトリより下層に存在するファイルに対して、指定された文字列で Grep 検索を行うツールです。  
プレーンテキストファイルに加えて、Word 、 Excel 、 PowerPoint のファイルの検索にも対応しているのが特徴です。

試作段階のため、対応済みの環境やファイルであっても実行に不備がある可能性があります。ご容赦ください。  
不具合などあった場合はご連絡いただけると大変ありがたいです。

<!-- omit in toc -->
## 目次

- [1. 対応ファイル](#1-対応ファイル)
  - [1.1. デフォルトで対応済みのファイル](#11-デフォルトで対応済みのファイル)
  - [1.2. 追加モジュールをインストールすると対応するファイル](#12-追加モジュールをインストールすると対応するファイル)
  - [1.3. 現行版では未対応のファイル](#13-現行版では未対応のファイル)
- [2. 動作環境](#2-動作環境)
- [3. 環境構築](#3-環境構築)
  - [3.1. Python のインストール](#31-python-のインストール)
  - [3.2. 当ツールのダウンロード・展開](#32-当ツールのダウンロード展開)
- [4. 使い方](#4-使い方)
  - [4.1. bat ファイルから実行する場合](#41-bat-ファイルから実行する場合)
  - [4.2. CLI から実行する場合](#42-cli-から実行する場合)
- [5. 補足・便利な使い方](#5-補足便利な使い方)
  - [5.1. bat ファイルをショートカットから実行](#51-bat-ファイルをショートカットから実行)
  - [5.2. 行・文字番号の見方](#52-行文字番号の見方)
  - [5.3. 正規表現や検索ファイル・除外ファイルの設定](#53-正規表現や検索ファイル除外ファイルの設定)
  - [5.4. PDF ファイルを検索できるようにする](#54-pdf-ファイルを検索できるようにする)
- [6. 留意点](#6-留意点)
  - [6.1. 【⚠️重要】検索するディレクトリはなるべく深い階層を設定する](#61-️重要検索するディレクトリはなるべく深い階層を設定する)
  - [6.2. 正規表現検索を Word と PowerPoint のファイルへ実行する場合の仕様について](#62-正規表現検索を-word-と-powerpoint-のファイルへ実行する場合の仕様について)
  - [6.3. bat ファイルを終了する際の出力について](#63-bat-ファイルを終了する際の出力について)
- [7. ライセンス](#7-ライセンス)
- [8. 参考サイト様](#8-参考サイト様)

## 1. 対応ファイル

### 1.1. デフォルトで対応済みのファイル

- プレーンテキストファイル（文字コードが UTF-8 のみ）
- Word ファイル（拡張子が `.docx` のみ）
- Excel ファイル（拡張子が `.xlsx` のみ） ※1
- Power Point ファイル（拡張子が `.pptx` のみ）

※1 Excel ファイルの場合、セルに直接記述されたテキストのみが取得できます。  
関数の出力結果やシェイプ中の文字は取得できません。

### 1.2. 追加モジュールをインストールすると対応するファイル

- PDF ファイル

導入方法は [5. 補足・便利な使い方 > 5.4. PDF ファイルを検索できるようにする](#54-pdf-ファイルを検索できるようにする) をご確認ください。

### 1.3. 現行版では未対応のファイル

- 文字コードが UTF-8 以外のプレーンテキストファイル（今後対応予定）
- 拡張子に `x` が付かない古い Office 系ファイル（`.xls` , `.ppt` , `.doc`）
  - 検索の際に利用している xml ファイルの仕様が違うらしいため。
- 隠しファイルや隠しディレクトリ配下の検索
  - 不要そうなので現在は未対応。今後対応させるかも。

## 2. 動作環境

- 動作確認済みの環境
  - Windows 10, 11
  - WSL2（Ubuntu）
  - macOS Monterey
- ソフトウェア
  - Python 3.6 以上（利用モジュールの推奨要件が 3.6 以上のため）

特殊な実装は特に行っていないので、たいていの環境で動くと思います。

## 3. 環境構築

環境構築手順を掲載しています。所要時間は 10 分ほどです。

### 3.1. Python のインストール

Python を導入していない方は、下記のサイトを参考に Python のインストールを行ってください。  
[Python のインストール方法 - Windows - Python の準備 - やさしい Python 入門](https://python.softmoco.com/devenv/how-to-install-python-windows.php)

### 3.2. 当ツールのダウンロード・展開

1. Python を導入済みの PC で [Releases · kenkenpa198/GrepOffice](https://github.com/kenkenpa198/GrepOffice/releases) にアクセスする。
2. 最新のリリース > `Assets` > `Source code (zip)` をクリックしてダウンロード後、好みのディレクトリに展開する。

以上で完了です。  
[4. 使い方](#4-使い方) を参考に、ツールが問題なく動くか確認してください。

ツールを削除したい場合は展開したディレクトリごと削除すれば OK です。

## 4. 使い方

### 4.1. bat ファイルから実行する場合

実行が手軽に行えるように実行用の bat ファイルを準備しています。  
Windows 環境の場合はこちらがオススメです。

1. 展開したディレクトリ配下の `Run-GrepOffice.bat` を起動する。
   1. 初回起動時のみ警告が出ると思いますので許可してください。
2. 検索対象のディレクトリパスを入力して `Enter` キーで送信する。
   1. 入力の際はコピペでも OK です。エクスプローラーのアドレスバーからコピペしてくるのがオススメです。
   2. **⚠️ ディレクトリパスはなるべく深い階層を指定してください。** 詳細は [6. 留意点 > 6.1. 【⚠️重要】検索するディレクトリはなるべく深い階層を設定する](#61-【⚠️重要】検索するディレクトリはなるべく深い階層を設定する) をご参照ください。
3. 検索条件とする文字列を入力して `Enter` キーで送信する。
4. 検索条件の一覧を確認したら `Enter` キーを押す。
5. 検索が開始され、検索結果が下記のイメージで表示される。

    ```text
    ▼ 検索結果
    ----------------------------------------------------------

    C:\Users\aaa\bbb\テキストファイル.txt (3, 1) : さしすせそ
    C:\Users\aaa\bbb\エクセルファイル.xlsx (3, 1) : さしすせそ
    C:\Users\aaa\bbb\ccc\パワポファイル.pptx (2, 1) : さしすせそたちつてとあいうえお

    ----------------------------------------------------------
    3 件ヒットしました。

    Grep Office を終了します。
    ```

基本的な使い方は以上です。

終了する際は画面の指示に従ってください。  
ウィンドウを直接閉じても OK です。

### 4.2. CLI から実行する場合

CLI から Python を直接実行しても OK です。  
WSL2 や MacOS などを利用している方はこちらでご利用ください。

1. 展開したディレクトリを PowerShell や Terminal などで開く。
2. `python3 GrepOffice.py` を実行する。
3. 以降の手順は bat ファイルでの手順と同様です。

## 5. 補足・便利な使い方

### 5.1. bat ファイルをショートカットから実行

bat ファイルはショートカットからの実行も可能です。  
ショートカットを作成して、デスクトップやタスクバーなど好きな場所に配置すると簡単に実行できて便利です。

参考: [【Windows10】バッチファイルをタスクバーにピン留めする方法 – Plane Note](https://note.z0i.net/2020/09/pin-batch-file.html)

### 5.2. 行・文字番号の見方

検索結果に表示される `(X, Y)` は、そのファイルの中でマッチした行（もしくはスライド）と文字の位置を表しています。

- プレーンテキストの場合
  - X: 上から○行目
  - Y: 左から○文字目
- Word ファイルの場合
  - X: 上から○行目
  - Y: 左から○文字目
- Excel ファイルの場合
  - X: 読み込めたテキストをセルごとに行として並べた状態での○行目（改善の余地あり。。）
  - Y: セルごとの左から○文字目
- Power Point ファイルの場合
  - X: スライド○枚目
  - Y: スライドごとの○文字目
- PDF ファイルの場合
  - X: ページをすべて連結した状態での○行目
  - Y: 左から○文字目

### 5.3. 正規表現や検索ファイル・除外ファイルの設定

ディレクトリ直下に存在する `setup.py` を編集することで、以下のような設定が行えます。

- 正規表現を使用して検索する
  - デフォルトは使用しません。
- 検索を限定したいファイル名を指定する
  - デフォルトは指定なしです。
- 検索から除外したいファイル名を設定する
  - デフォルトは `.exe` 等いくつかのファイルを設定済みです。
- PDF ファイルを検索する
  - 外部モジュールのインストールが必要なため、デフォルトでは検索しません。次項の手順を参考に有効化してください。

詳しくは `setup.py` のコメントを参考に設定を行ってください。  
除外ファイルのデフォルト値は [サクラエディタ](https://sakura-editor.github.io/) の初期設定値から引用させていただいています。

### 5.4. PDF ファイルを検索できるようにする

デフォルトでは PDF ファイルの検索に対応していませんが、  
外部モジュールをインストールすることで検索が可能になります。

Windows 環境の場合は下記の手順でインストールと設定の変更を行ってください。  
所要時間は 10 分ほどです。

MacOS 等でも同様の手順で可能です。

1. ツールを展開したディレクトリパスをエクスプローラーで開く。
2. `Shift` キーを押しながら右クリック > `PowerShell ウィンドウをここで開く(S)` を選択する。
3. PowerShell が開き、パスの部分に展開したディレクトリパスが表示されていることを確認する。
4. `pip3 -V` を送信し、pip コマンドが使えることを確認する。
5. `pip3 install --upgrade pip` を送信し、pip をアップデートする。
6. `pip3 install -r requirements.txt` を送信し、同ディレクトリの `requirements.txt` に記載されたモジュールをインストールする。
7. インストールが完了し、エラーなどが出ていないことを確認する。
8. `pip3 list` を送信し、いろいろ表示された中から下記のように `pdfminer.six` が表示されていることを確認する。

    ```text
    Package                   Version
    ------------------------- --------
    ...
    pdfminer.six              20211012
    ...
    ```

9. ディレクトリ上の `setup.py` を開く。
10. `■ PDF の検索設定` とコメントで記述されている箇所に存在する定数 `DETECT_PDF_FILE` を下記のように `True` へ書き換えて保存する。

    ```python
    DETECT_PDF_FILE = True
    ```

11. ツールを bat ファイル等から実行して下記を確認する。
    1. エラーなどが出力されないこと。
    2. 検索条件の入力後に表示される検索情報に `PDF ファイル     : 検索する` と表示されていること。
    3. 検索を実行してみて、PDF ファイルの内容が読み取れていること。
12. 完了です！

## 6. 留意点

### 6.1. 【⚠️重要】検索するディレクトリはなるべく深い階層を設定する

このツールは、指定ディレクトリ配下に存在するファイルを順番に開いて検索するという処理を繰り返す仕様です。

ファイルサイズやファイル数などでの制限は現行だと特に設けていないため、  
**指定した階層が浅すぎると、処理に時間がかかったり PC や検索先のサーバー等に負荷をかけてしまう可能性があります。**

そのため、検索対象にするディレクトリはなるべく深い階層を指定するようにしてください。

### 6.2. 正規表現検索を Word と PowerPoint のファイルへ実行する場合の仕様について

Word と PowerPoint のファイルに対して正規表現での検索を使用する場合、行頭 `^` や行末 `$` の指定が正しく検知できない可能性があります。  
これらのファイルの場合は読み込んだ文字列を結合してから検索するので、ソースのテキストの行頭や行末が行中になってしまうことがあるためです。

そのためこれらの正規表現を用いる際はご留意ください。

### 6.3. bat ファイルを終了する際の出力について

bat ファイルでツールを起動中に `Ctrl + C` を送信して終了する際、 `Terminate batch job (Y/N)?` （訳: バッチ ジョブを終了しますか？）とメッセージが出力されます。  
[これは cmd.exe の仕様によるものだそう](https://mattn.kaoriya.net/software/windows/20120920154016.htm) で、動作には特に影響ありません。

表示された場合は `N` を入力して送信するか、直接ウィンドウを閉じて OK です。

## 7. ライセンス

このアプリケーションは MIT ライセンスの下でリリースされています。  
ライセンス全文はディレクトリ直下の `LICENSE` ファイルをご確認ください。

## 8. 参考サイト様

主に以下のサイト様の情報を参考にさせていただきました🙇‍♂️  
ソースコードのコメント内でも引用させていただいております。

- [Windows10 - ExcelやWord内の写真を画像として取り出し保存する方法 - Win10ラボ](https://win10labo.info/win10-excel-photo/)
- [PythonでExcelファイルを操作する!(1) /実はzip圧縮ファイルだって知ってました？│YUUKOU's 経験値](https://yuukou-exp.plus/handle-xlsx-with-python-intro/)
- [Pythonを使ってzipファイルからファイルを取り出す - Qiita](https://qiita.com/mriho/items/f82c66e7a232b6b37206)
- [パワーポイント内のテキストをgrepする - Qiita](https://qiita.com/kaityo256/items/2977d53e70bbffd4d601)
- [PythonでのXMLファイル操作例 - Qiita](https://qiita.com/sino20023/items/0314438d397240e56576)
- [【Python】pdfminerでPDFからテキストを抽出する | ジコログ](https://self-development.info/%E3%80%90python%E3%80%91pdfminer%E3%81%A7pdf%E3%81%8B%E3%82%89%E3%83%86%E3%82%AD%E3%82%B9%E3%83%88%E3%82%92%E6%8A%BD%E5%87%BA%E3%81%99%E3%82%8B/)
- [基本的な正規表現一覧 | murashun.jp](https://murashun.jp/article/programming/regular-expression.html)
