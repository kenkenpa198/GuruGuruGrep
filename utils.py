import glob
import os
import re
import xml.etree.ElementTree as ET
from zipfile import ZipFile


'''
■ .xlsx ファイルに書き込まれているテキストをリストにして返す関数

.xlsx や .docx .pptx などの Office 系ファイルの実体は zip ファイルだそう。
なので zip ファイルとして展開ができ、また展開されたディレクトリ内にはテキストや画像データが一式詰まっています。

この特性を利用して、xlsx ファイルを zip アーカイブファイルとして読み込み、
テキスト情報が詰まった xml ファイルからテキスト情報を抜き出してリスト形式で返すのがこの関数の役目です。
pptx 用の関数も同じような感じです。

テキスト情報が詰まった xml ファイルはツリー上のファイル形式なので、組み込みのモジュールを使って要素を抜き出しています。

参考:
https://yuukou-exp.plus/handle-xlsx-with-python-intro/
https://qiita.com/sino20023/items/0314438d397240e56576
https://qiita.com/mriho/items/f82c66e7a232b6b37206
https://qiita.com/kaityo256/items/2977d53e70bbffd4d601
'''
def make_xlsx_text_list(src_file_path):

    # zip アーカイブのオブジェクトを生成
    zip = ZipFile(src_file_path)

    # xlsx ファイルの中に存在する xml ファイルの読み込み
    with zip.open('xl/sharedStrings.xml', 'r') as f:
        tree = ET.parse(f)                           # zip ファイルの中に存在する xml ファイルをパースされたオブジェクトとして読み込む
    root = tree.getroot()                            # 親階層の要素を取り出す

    # 親階層から見て子階層に存在するテキストをリストへ格納する
    text_list = []
    for child in root:
        for child_2 in child:
            if child_2.text:
                text_list.append(child_2.text)
    text_list_fmt = ['(改行)'.join(s.splitlines()) for s in text_list] # 改行が入っていると見づらいので改行を適当な文字列へ置換する
    return text_list_fmt


'''
■ .pptx ファイルに書き込まれているテキストを多次元リストにして返す関数（作成中）

前述の関数のパワーポイント版。
展開まではエクセル用と同じですが、パワーポイントの場合はスライドごとに xml ファイルが分かれているので
スライドごとにテキストを格納した多次元リストとして返します。

パワーポイント版関数のテキストの検知はエクセル版と違い正規表現のパターンマッチを利用しています。
たぶんエクセル版と同じように子階層を掘っていくやり方でもできると思いますが、
階層がエクセルよりも深いようで記述が冗長になりそうだったのでこちらを採用しました。
'''
def make_pptx_text_list(src_file_path):

    # zip アーカイブのオブジェクトを生成
    zip = ZipFile(src_file_path)

    # ppt/slides/slideXX.xml をリストへ格納する
    slide_path_list = [
        zfp.filename for zfp in zip.filelist
        if zfp.filename.startswith('ppt/slides/slide')
    ]

    print(slide_path_list)

    # slideXX.xml ごとにリスト化を繰り返して多次元リストを作成
    slide_text_list = []
    for slide_path in slide_path_list:

        # xml ファイルをデコードして読み込む
        with zip.open(slide_path, 'r') as f:
            body = f.read().decode('utf-8')

        # パターンマッチでテキストが格納されている箇所を検知してリスト化する
        text_list = re.findall(r'\<a\:t\>.+?\<\/a\:t\>', body)
        text_list_fmt = [s.lstrip('<a:t>').rstrip('</a:t>') for s in text_list] # 左右のタグを削除する

        slide_text_list.append(text_list_fmt)

    print(slide_text_list)

    return text_list_fmt


'''
■ 与えられた文字列をリスト上から検索して結果をプリントする関数
'''
def search_text(src_list, search_text, file_path, hit_num):
    line_num = 1
    for line in src_list:
        m = re.search(search_text, line)
        if m:
            # ファイルパス.txt (X行目, Y文字目) : 行のテキスト
            out = '%s (%d, %d) : %s' % (file_path, line_num, m.start() + 1, line.rstrip())
            print(out)
            hit_num = hit_num + 1
        line_num = line_num + 1
    return hit_num

'''
■ 渡された内容をプリントする関数（例外用）
'''
def print_result_error(file_path, txt, error):
    return '%s : %s <%s>' % (file_path, txt, error)