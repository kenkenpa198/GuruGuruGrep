import csv
import datetime
import os
import re
import xml.etree.ElementTree as ET
import zipfile

import pdfminer.high_level as pdfm_hl
import pandas as pd
import tqdm

import setup


'''
■ 入力情報をテキストで出力する関数

入力情報をテキスト形式で保存する。
'''
def export_input_data(export_dir_path, search_dir, keyword, regexp_flag=False, filter_path=None, exclude_path=None, excel_search_setting=None):
    dt_now = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
    filename = f'{dt_now}_input.txt'

    os.makedirs(export_dir_path, exist_ok=True)
    with open(os.path.join(export_dir_path, filename), 'w', newline='') as f:
        f.write('▼検索情報\n----------------------------------------------------------\n\n')

        f.write(f'検索ディレクトリ : {search_dir}\n')
        f.write(f'検索キーワード   : {keyword}\n')

        if regexp_flag:
            f.write(f'正規表現で検索   : 使用する\n')
        else:
            f.write(f'正規表現で検索   : 使用しない\n')

        if filter_path:
            f.write(f'検索の絞り込み   : {setup.FILTER_PATH}\n')
        else:
            f.write(f'検索の絞り込み   : 設定なし\n')

        if exclude_path:
            f.write(f'検索から除外     : {setup.EXCLUDE_PATH}\n')
        else:
            f.write(f'検索から除外     : 設定なし\n')

        if excel_search_setting:
            f.write(f'Excel の検索設定 : {setup.EXCEL_SEARCH_SETTING}\n')
        else:
            f.write(f'Excel の検索設定 : 設定値が None のようです。setup.py をご確認ください。\n')

        f.write('\n----------------------------------------------------------\n')

    return os.path.join(export_dir_path, filename)



'''
■ CSV 生成用の多次元リストへ追加する関数

検索結果をCSV 出力用の多次元リストに追加する。
与えられたファイルパスはディレクトリパスとファイル名に分割して格納する。
'''
result_multi_list = []

def append_result_multi_list(file_path, line_num, char_num, match_text='-', page_name='-'):

    filename = os.path.basename(file_path)
    dir_path = os.path.dirname(file_path)

    result_list = [dir_path, filename, page_name, line_num, char_num, match_text]
    result_multi_list.append(result_list)
    return result_multi_list



'''
■ 結果を CSV で出力する関数

渡された多次元リストを CSV として出力する。
'''

def export_result_csv(export_dir_path, multi_list):
    dt_now = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
    filename = f'{dt_now}_result.csv'

    os.makedirs(export_dir_path, exist_ok=True)
    with open(os.path.join(export_dir_path, filename), 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['dir_path', 'filename', 'page_num', 'line_num', 'char_num', 'match_text'])

        # リストの NULL 判定
        if multi_list:
            writer.writerows(multi_list)

    return os.path.join(export_dir_path, filename)


'''
■ 与えられた文字列を検索して結果テキストを返す関数

正規表現検索設定のオンオフによって検索方法を切り替える。
マッチした場合は結果テキストの文字列を返し、 CSV 出力用の多次元リストへ追加する。
マッチしなかった場合は何も返さない。
'''
def search_keyword(target_text, keyword, regexp_flag, file_path, line_num, page_name=None,):

    match_num = None
    result_text = None

    # 正規表現検索フラグが True の場合は正規表現マッチを使用する
    if regexp_flag:
        m = re.search(keyword, target_text)
        if m:
            match_num = m.start() + 1
            match_text = target_text.rstrip()

    # 正規表現検索フラグが True でない場合は find() を使用する
    else:
        s = target_text.find(keyword)
        if s >= 0:
            match_num = s + 1
            match_text = target_text.rstrip()

    # ページ名を表示しない場合: ディレクトリパス (行番号, 列番号) : 改行を削除したテキスト
    if match_num and page_name:
        result_text = '%s (%s, %d, %d) : %s' % (file_path, page_name, line_num, match_num, match_text.replace('\n',''))
        append_result_multi_list(file_path, line_num, match_num, match_text, page_name)

    # ページ名を表示しない場合: ディレクトリパス (ページ名, 行番号, 列番号) : 改行を削除したテキスト
    if match_num and not page_name:
        result_text = '%s (%d, %d) : %s' % (file_path, line_num, match_num, match_text.replace('\n',''))
        append_result_multi_list(file_path, line_num, match_num, match_text)

    return result_text


'''
■ 与えられた文字列をリストから検索して結果をプリントする関数
引数で指定されたリストに対して「与えられた文字列を検索して結果を返す関数」を繰り返し処理し、
結果が返ってくれば結果テキストをプリントする。

ヒット件数はヒットごとにインクリメントし、最終的に合算するため戻り値として返す。
'''
def search_to_print_from_list(target_text_list, keyword, regexp_flag, file_path, page_name=None):
    line_num = 0
    hit_num = 0

    for target_text in target_text_list:
        line_num += 1

        search_result = search_keyword(target_text, keyword, regexp_flag, file_path, line_num, page_name)
        if search_result:
            hit_num += 1
            tqdm.tqdm.write(search_result)

    return hit_num


'''
■ Office ファイルからテキスト情報を取り出せるか確認するチェック用関数

.xlsx .pptx .docx が空のファイルの場合、zip アーカイブ化できなかったり
テキスト情報が詰まった xml ファイルが存在しなかったりしてエラーを吐いていたので
テキスト情報を取り出せるか否かの事前チェックを行う関数。

ファイルによって以下のように違いがあるようなので、どの場合でも検出できるようにしている。
    - 空の .xlsx の場合: zip ファイルとして扱えるが、xml ファイルが存在しない
    - 空の .pptx の場合: zip ファイルとしてそもそも扱えない
    - 空の .docx の場合: zip ファイルとしてそもそも扱えない
'''
def is_gettable_text(check_file_path, xml_file_path):

    # 指定したファイルが zip アーカイブファイルとして扱えない場合は False を返す
    if not zipfile.is_zipfile(check_file_path):
        return False

    # zip ファイルとして扱えてもテキスト情報が保管された xml ファイルが存在しない場合は False を返す
    zip = zipfile.ZipFile(check_file_path)
    if not xml_file_path in zip.namelist():
        return False

    # チェックを通ったもののみ True を返す
    return True


'''
■ .xlsx ファイルに書き込まれているテキストをリストにして返す関数

.xlsx や .docx .pptx などの Office 系ファイルの実体は zip ファイルだそう。
なので zip ファイルとして展開ができ、また展開されたディレクトリ内にはテキストや画像データが一式詰まっている。

この特性を利用して、xlsx ファイルを zip アーカイブファイルとして読み込み、
テキスト情報が詰まった xml ファイルからテキスト情報を抜き出してリスト形式で返すのがこの関数の役目。
パワーポイント用、ワード用の関数も同じような感じ。

テキスト情報が詰まった xml ファイルはツリー状のファイル形式でそのままだと扱いにくいので、
組み込みの xml モジュールを使って要素を抜き出している。

参考:
https://yuukou-exp.plus/handle-xlsx-with-python-intro/
https://qiita.com/sino20023/items/0314438d397240e56576
https://qiita.com/mriho/items/f82c66e7a232b6b37206
'''
def make_xlsx_text_list(src_file_path):

    # テキスト情報が詰まった xml ファイルの指定
    xml_file_path = 'xl/sharedStrings.xml'

    # テキスト情報を取り出せない場合は空文字を返す
    if not is_gettable_text(src_file_path, xml_file_path):
        return ''

    # zip アーカイブ化
    zip = zipfile.ZipFile(src_file_path)

    # xlsx ファイルの中に存在する xml ファイルの読み込み
    with zip.open(xml_file_path, 'r') as f:
        tree = ET.parse(f)                           # zip ファイルの中に存在する xml ファイルをパースされたオブジェクトとして読み込む
    root = tree.getroot()                            # 親階層の要素を取り出す

    # 親階層から見て子階層に存在するテキストをリストへ格納する
    text_list = []
    for child in root:
        for child_2 in child:
            if child_2.text:
                text_list.append(child_2.text)
    text_list_fmt = [''.join(s.splitlines()) for s in text_list] # 改行が入っていると見づらいので改行を削除する

    return text_list_fmt


'''
■ .pptx ファイルに書き込まれているテキストをリストにして返す関数

前述の関数のパワーポイント版。
展開まではエクセル用と同じだが、パワーポイントの場合はスライドごとに xml ファイルが分かれているので
先にスライドごとのファイルパスを格納したリストを生成し、それをベースにしてテキストを格納したリストを作成している。

パワーポイント版関数のテキストの検知はエクセル版と違い正規表現のパターンマッチを利用している。
たぶんエクセル版と同じように子階層を掘っていくやり方でもできると思うが、
階層がエクセルよりも深いようで記述が冗長になりそうだったのでこちらを採用している。

パワーポイントの xml に記述されたテキストは空白や文字の種類を境に
テキストが細切れになっているため、ページごとに文字を結合している。

参考:
https://qiita.com/kaityo256/items/2977d53e70bbffd4d601

TODO: パラグラフの文字列ごとに切り分けられそうなのでがんばる
'''
def search_to_print_from_pptx(src_file_path, keyword, regexp_flag, ):

    # テキスト情報が詰まった xml ファイルの指定
    xml_file_path_left = 'ppt/slides/slide'
    xml_file_path = 'ppt/slides/slide1.xml'

    # テキスト情報を取り出せない場合は 0 を返す
    if not is_gettable_text(src_file_path, xml_file_path):
        return 0

    # zip アーカイブ化
    zip = zipfile.ZipFile(src_file_path)

    # ppt/slides/slideXX.xml をリストへ格納する
    slide_path_list = [
        zfp.filename for zfp in zip.filelist
        if zfp.filename.startswith(xml_file_path_left)
    ]

    # インクリメント用変数を定義
    slide_num = 0
    line_num = 0
    hit_num = 0

    # slideXX.xml ごとにテキスト情報のリスト化⇒結合を繰り返してテキストリストを作成
    for slide_path in slide_path_list:
        slide_num += 1
        line_num += 1

        # xml ファイルをデコードして読み込む
        with zip.open(slide_path, 'r') as f:
            body = f.read().decode('utf-8')

        # パターンマッチでテキストが格納されている箇所を検知してリスト化する
        find_text_list = re.findall(r'\<a\:t\>.+?\<\/a\:t\>', body)
        find_text_list_fmt = [s.lstrip('<a:t>').rstrip('</a:t>') for s in find_text_list] # 左右のタグを削除する

        # テキストを結合してリストへ再格納する
        target_text = ''.join(find_text_list_fmt)

        # 検索 & 出力
        search_result = search_keyword(target_text, keyword, regexp_flag, src_file_path, line_num, page_name='slide' + str(slide_num))
        if search_result:
            tqdm.tqdm.write(search_result)
            hit_num += 1

    return hit_num


'''
■ .docx ファイルに書き込まれているテキストをリストにして返す関数

前述の関数のワード版。処理の流れは前述の関数と同様。

ワードの xml ファイルの場合は、ひとつの xml ファイルの中に細切れになったテキストが格納されているようなイメージ。
なのでパワーポイントと同様に正規表現のパターンマッチ + 文字の結合を行っている。

パワーポイント版のようにページごとにファイルが分けられているのではないため、ページごとの行番号の出力は不可能そう。
そういう情報が保存されていればできそうなので調べてみてもいいかも？
'''
def make_docx_text_list(src_file_path):

    # テキスト情報が詰まった xml ファイルの指定
    xml_file_path = 'word/document.xml'

    # テキスト情報を取り出せない場合は空文字を返す
    if not is_gettable_text(src_file_path, xml_file_path):
        return ''

    # zip アーカイブ化
    zip = zipfile.ZipFile(src_file_path)

    # xml ファイルをデコードして読み込む
    with zip.open(xml_file_path, 'r') as f:
        body = f.read().decode('utf-8')

    # テキスト要素を行ごとに一旦リスト化する
    find_line_list = re.findall(r'\<w\:p.+?\<\/w\:p\>', body)

    # [[1行目 要素1, 1行目 要素2], [2行目 要素1, 2行目 要素2], ...] という形の多次元リストに変換する
    for index, value in enumerate(find_line_list):
        find_line_list[index] = re.findall(r'\<w\:t\>.+?\<\/w\:t\>', value)                         # 細切れの要素ごとにリスト化して格納する
        find_line_list[index] = [s.lstrip('<w:t>').rstrip('</w:t>') for s in find_line_list[index]] # 細切れの要素ごとに付いている左右のタグを削除する

    # 多次元リスト中のテキストを結合して出力用のリストへ格納する
    text_list = []
    for i in find_line_list:
        text_list.append(''.join(i))

    return text_list


'''
■ PDF ファイルに書き込まれているテキストをリストにして返す関数

PDF ファイルを扱える外部モジュール pdfminer.sys を利用してテキストを取り出します。
ページごとにリストを生成したかったけど自分には難しかったので行ごとに取り出すことにしました。

参考:
https://self-development.info/%E3%80%90python%E3%80%91pdfminer%E3%81%A7pdf%E3%81%8B%E3%82%89%E3%83%86%E3%82%AD%E3%82%B9%E3%83%88%E3%82%92%E6%8A%BD%E5%87%BA%E3%81%99%E3%82%8B/
'''
def make_pdf_text_list(src_file_path):

    text = pdfm_hl.extract_text(src_file_path)
    text_list = text.splitlines()

    return text_list


'''
■ Excel 用新関数
Pandas モジュールを使って読み込みを行う関数。
pandas.DataFrame から生成した多次元リストを基に検索を行う。

事前設定によって出力の仕様を変更する。
'''

setup.EXCEL_SEARCH_SETTING

def search_to_print_from_xlsx(src_file_path, keyword, regexp_flag, ):

    # Excel ファイルを pandas.DataFrame の辞書として読み込み
    df_dict = pd.read_excel(src_file_path, sheet_name=None, header=None, index_col=None, dtype=str)

    # JOIN_ROW : 行ごとの多次元リストを結合して検索する
    if setup.EXCEL_SEARCH_SETTING == 'JOIN_ROW':

        # インクリメント用変数を定義
        hit_num = 0

        # シートごとに多次元リストを生成して検索する繰り返し処理
        for key in df_dict:
            line_num = 0

            # 行ごとの多次元リストを生成
            row_multi_list = df_dict[key].fillna('').to_numpy().tolist()

            # 多次元リスト中の行を結合して検索とプリント
            for row_list in row_multi_list:
                line_num += 1
                search_result = search_keyword(' '.join(row_list), keyword, regexp_flag, src_file_path, line_num, key)
                if search_result:
                    tqdm.tqdm.write(search_result)
                    hit_num += 1

    # JOIN_COLUMN : 列ごとの多次元リストを結合して検索する
    if setup.EXCEL_SEARCH_SETTING == 'JOIN_COLUMN':

        # インクリメント用変数を定義
        hit_num = 0

        # シートごとに多次元リストを生成して検索する繰り返し処理
        for key in df_dict:
            line_num = 0

            # 行ごとの多次元リストを生成
            row_multi_list = df_dict[key].fillna('').to_numpy().tolist()

            # 列ごとの多次元リストへ変換
            col_multi_list = [list(x) for x in zip(*row_multi_list)]

            # 多次元リスト中の行を結合して検索とプリント
            for col_list in col_multi_list:
                line_num += 1
                search_result = search_keyword(' '.join(col_list), keyword, regexp_flag, src_file_path, line_num, key)
                if search_result:
                    tqdm.tqdm.write(search_result)
                    hit_num += 1


        # 行ごとの多次元リストを分割したまま検索する
        # TODO: 作成する！
        if setup.EXCEL_SEARCH_SETTING == 'SPLIT_ROW':
            # 行ごとの多次元リストをそのまま使う
            tqdm.tqdm.write('Excel 検索設定 SPLIT_ROW は未実装です。他の設定を指定してください。')
            return


        # 列ごとの多次元リストを分割したまま検索する
        # TODO: 作成する！
        if setup.EXCEL_SEARCH_SETTING == 'SPLIT_COLUMN':
            # 列ごとの多次元リストへ変換
            tqdm.tqdm.write('Excel 検索設定 SPLIT_COLUMN は未実装です。他の設定を指定してください。')
            return

    return hit_num
