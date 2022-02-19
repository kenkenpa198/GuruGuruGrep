<!-- omit in toc -->
# Grep Office

![イメージ図](images/kv.png)

指定ディレクトリの配下に存在するファイルをもとに、指定された文字列で Grep 検索するツールです。  
Word 、 Excel 、 PowerPoint のファイル（拡張子が `.docx` , `.xlsx` , `.pptx` のもののみ）の検索にも対応しているのが特徴です。

試作段階のため、対応済みの環境やファイルであっても実行に不備がある可能性があります。ご容赦ください。  
不具合などあった場合はご連絡いただけると大変ありがたいです。

## 1. 対応ファイル

### 1.1. 対応済みのファイル

- プレーンテキストファイル（UTF-8 のみ）
- Word ファイル（`.docx` のみ）
- Excel ファイル（`.xlsx` のみ）
- Power Point ファイル（`.pptx` のみ）

※ Excel ファイルの場合、セルに直接記述されたテキストのみが取得できます。関数の出力結果やシェイプ中の文字は取得できません。

### 1.2. 今後対応させたいファイル

- UTF-8 以外のプレーンテキストファイル（Shift-JIS など）
- PDF ファイル

### 1.3. 未対応のファイル

- 拡張子に `x` が付かない古い Office 系ファイル（`.xls` , `.ppt` , `.doc`）
  - 検索の際に利用している xml ファイルの仕様が違うらしいため。
- 隠しファイルや隠しディレクトリ配下の検索
  - モジュールのオプション指定で簡単に対応できそうですが、不要そうなので現在は未対応。今後対応させるかも。

## 2. 動作環境

- 動作確認済みの環境
  - Windows 10, 11
  - WSL2（Ubuntu）
  - macOS Monterey
- ソフトウェア
  - Python 3.5 以上（glob モジュールのオプション指定に必要なため、3.5 以上を要件としています）

特殊な実装は特に行っていないので、たいていの環境で動くと思います。

コーディングと動作確認は Windows 及び WSL2（Ubuntu）環境上で行っています。  
MacOS は簡単な動作確認のみ行っています。

## 3. 環境構築

動作環境の要件を満たしている PC へ `Code > Download ZIP` からダウンロードして、好みのディレクトリに展開してください。

Python そのものの構築手順はもう少し詳しく準備予定。  

## 4. 使い方

### 4.1. bat ファイルから実行する場合

実行が手軽に行えるように実行用の bat ファイルを準備しています。  
Windows 環境の場合はこちらがオススメです。

1. 展開したディレクトリ配下の `Run-GrepOffice.bat` を起動する。
   1. 初回起動時のみ警告が出ると思いますので許可してください。
2. 画面の指示に従って、検索対象のディレクトリパスと検索条件の文字列を入力して送信する。
3. 検索条件の一覧を確認したら `Enter` キーを押す。
4. 検索が開始する。
5. 検索結果が下記のイメージで表示される。

    ```text
    ▼ 検索結果
    ----------------------------------------------------------

    C:\Users\aaa\bbb\エクセルファイル.xlsx (3, 1) : さしすせそ
    C:\Users\aaa\bbb\テキストファイル.txt (3, 1) : さしすせそ
    C:\Users\aaa\bbb\パワポファイル.pptx (2, 1) : さしすせそたちつてとあいうえお

    ----------------------------------------------------------
    3 件ヒットしました。

    Grep Office を終了します。
    ```

基本的な使い方は以上です。

終了する際は画面の指示に従ってください。  
ウィンドウを直接閉じても OK です。

### 4.2. CLI から実行する場合

CLI から Python を直接実行しても OK です。  
WSL2 や MacOS を利用している方はこちらでご利用ください。

1. 展開したディレクトリを PowerShell や Terminal などで開く。
2. `python3 GrepOffice.py` を実行する。
3. 以降の手順は bat ファイルでの手順と同様です。

## 5. 補足・便利な使い方

### 5.1. 行・文字番号の見方

検索結果に表示される `(X, Y)` は、そのファイルの中でマッチした行（もしくはスライド）と文字の位置を表しています。  
確認する際の目安にお使いください。

- プレーンテキストの場合
  - X: ○行目
  - Y: ○文字目
- Word ファイルの場合
  - X: 必ず `1` が出力されます。
  - Y: 読み込めたすべての文字情報を連結させた状態での○文字目
- Excel ファイルの場合
  - X: 読み込めたテキストをセルごとに行として並べた状態での○行目
  - Y: ○文字目
- Power Point ファイルの場合
  - X: スライド○枚目
  - Y: スライドごとの○文字目

Word 、Excel の出力はもう少し改善したいなあと思っています。。

### 5.2. 検索ファイル・除外ファイルの設定

展開したディレクトリ直下に存在する `setup.py` を編集することで、検索するファイルを絞り込んだり検索から除外するファイルを指定したりできます。  
詳しくは `setup.py` のコメントを参照してください。

### 5.3. ショートカットから実行

bat ファイルはショートカットからの実行も可能です。  
ショートカットを作成して、デスクトップやスタートメニュー、タスクバーなど好きな場所に配置すると簡単に実行できて便利です。  

参考: [【Windows10】バッチファイルをタスクバーにピン留めする方法 – Plane Note](https://note.z0i.net/2020/09/pin-batch-file.html)

## 6. 留意点

### 6.1. 【⚠️重要】検索するディレクトリはなるべく深めに設定する

このツールは、指定ディレクトリ配下に存在する展開可能なファイルを順番に開いて検索するという処理を繰り返す仕様です。

ファイルサイズやファイル数指定などでの制限は現行だと特に設けていないため、  
**指定した階層が浅すぎると PC や検索先のサーバー等に負荷をかけてしまう可能性があります。**

そのため、検索対象にするディレクトリはなるべく深い階層を指定するようにしてください。

### 6.2. bat を終了する際の出力について

bat でツールを起動中に `Ctrl + C` を送信して終了する際、 `Terminate batch job (Y/N)?` （訳: バッチ ジョブを終了しますか？）とメッセージが出力されます。  
これは cmd.exe の仕様によるものだそうで、動作には特に影響ありません。

表示された場合は `N` を入力して送信するか、直接ウィンドウを閉じて OK です。

## 7. ライセンス

このアプリケーションは MIT ライセンスの下でリリースされています。  
ライセンス全文はディレクトリ直下の `LICENSE` ファイルをご確認ください。

## 8. 参考サイト様

以下のサイト様の情報を主に参考にさせていただきました。

- [Windows10 - ExcelやWord内の写真を画像として取り出し保存する方法 - Win10ラボ](https://win10labo.info/win10-excel-photo/)
- [PythonでExcelファイルを操作する!(1) /実はzip圧縮ファイルだって知ってました？│YUUKOU's 経験値](https://yuukou-exp.plus/handle-xlsx-with-python-intro/)
- [Pythonを使ってzipファイルからファイルを取り出す - Qiita](https://qiita.com/mriho/items/f82c66e7a232b6b37206)
- [パワーポイント内のテキストをgrepする - Qiita](https://qiita.com/kaityo256/items/2977d53e70bbffd4d601)
- [PythonでのXMLファイル操作例 - Qiita](https://qiita.com/sino20023/items/0314438d397240e56576)
