import re
import xml.etree.ElementTree as ET
import zipfile

import setup

# PDF の検索設定が True の場合、pdfminer.sys を import する
if setup.DETECT_PDF_FILE:
    import pdfminer.high_level as pdfm_hl

import pandas as pd
from itertools import chain as it_ch

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
'''
def make_pptx_text_list(src_file_path):

    # テキスト情報が詰まった xml ファイルの指定
    xml_file_path_left = 'ppt/slides/slide'
    xml_file_path = 'ppt/slides/slide1.xml'

    # テキスト情報を取り出せない場合は空文字を返す
    if not is_gettable_text(src_file_path, xml_file_path):
        return ''

    # zip アーカイブ化
    zip = zipfile.ZipFile(src_file_path)

    # ppt/slides/slideXX.xml をリストへ格納する
    slide_path_list = [
        zfp.filename for zfp in zip.filelist
        if zfp.filename.startswith(xml_file_path_left)
    ]

    # slideXX.xml ごとにテキスト情報のリスト化⇒結合を繰り返してテキストリストを作成
    text_list_fmt = []
    for slide_path in slide_path_list:

        # xml ファイルをデコードして読み込む
        with zip.open(slide_path, 'r') as f:
            body = f.read().decode('utf-8')

        # パターンマッチでテキストが格納されている箇所を検知してリスト化する
        find_text_list = re.findall(r'\<a\:t\>.+?\<\/a\:t\>', body)
        find_text_list_fmt = [s.lstrip('<a:t>').rstrip('</a:t>') for s in find_text_list] # 左右のタグを削除する

        # テキストを結合してリストへ再格納する
        text_list_fmt.append(''.join(find_text_list_fmt))

    return text_list_fmt


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
■ 与えられた文字列を検索して結果テキストを返す関数

正規表現検索設定のオンオフによって検索方法を切り替える。
マッチした場合、結果テキストの文字列を返す。
マッチしなかった場合は何も返さない。
'''
def search_keyword(target_text, keyword, file_path, line_num):

    match_num = None
    result_text = None

    # 正規表現検索が True の場合は正規表現マッチを使用する
    if setup.USE_REGEXP:
        m = re.search(keyword, target_text)
        if m:
            match_num = m.start() + 1
            match_text = target_text.rstrip()

    # 正規表現検索が True でない場合は find() を使用する
    else:
        s = target_text.find(keyword)
        if s >= 0:
            match_num = s + 1
            match_text = target_text.rstrip()

    if match_num:
        result_text = '%s (%d, %d) : %s' % (file_path, line_num, match_num, match_text)

    return result_text


'''
■ 与えられた文字列をリストから検索して結果をプリントする関数
引数で指定されたリストに対して「与えられた文字列を検索して結果を返す関数」を繰り返し処理し、
結果が返ってくれば結果テキストをプリントする。

ヒット件数はヒットごとにインクリメントし、最終的に合算するため戻り値として返却する。
'''
def search_to_print_from_list(target_text_list, keyword, file_path):
    line_num = 0
    hit_num = 0

    for target_text in target_text_list:
        line_num += 1

        search_result = search_keyword(target_text, keyword, file_path, line_num)
        if search_result:
            print(search_result)
            hit_num += 1

    return hit_num


'''
■ Excel 新関数
TODO: 作成中
'''
def search_to_print_xlsx(src_file_path, search_text):

    EXCEL_SETTING = 'JOIN_ROW'
    text_list = []

    match_num_xlsx = 0
    # Excel ファイルを pandas.DataFrame の辞書として読み込み
    df_dict = pd.read_excel(src_file_path, sheet_name=None, header=None, index_col=None)

    # シートごとに多次元リストを生成する繰り返し処理
    n = 0
    for key in df_dict:

        # 行ごとの多次元リストを生成
        row_multi_list = df_dict[key].fillna('').to_numpy().tolist()

        # 行ごとの多次元リストを分割したまま検索する
        if EXCEL_SETTING == 'SPLIT_ROW':
            # 行ごとの多次元リストをそのまま使う
            print(row_multi_list)
            return

        # 行ごとの多次元リストを結合して検索する
        if EXCEL_SETTING == 'JOIN_ROW':
            # 行ごとの多次元リストを結合する
            join_row_text = ''.join(map(str, list(it_ch.from_iterable(row_multi_list))))
            text_list.append(join_row_text)


        # 列ごとの多次元リストを分割したまま検索する
        if EXCEL_SETTING == 'SPLIT_COLUMN':
            # 列ごとの多次元リストへ変換
            col_multi_list = [list(x) for x in zip(*row_multi_list)]
            print(col_multi_list)
            return

        # 列ごとの多次元リストを結合して検索する
        if EXCEL_SETTING == 'JOIN_COLUMN':
            # 列ごとの多次元リストへ変換
            col_multi_list = [list(x) for x in zip(*row_multi_list)]
            print(col_multi_list)
            return

        n += 1
    print(text_list)
    match_num_xlsx += search_to_print_result(text_list, search_text, src_file_path)
    return match_num_xlsx
