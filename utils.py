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
パワーポイント用、ワード用の関数も同じような感じです。

テキスト情報が詰まった xml ファイルはツリー状のファイル形式でそのままだと扱いにくいので、
組み込みの xml モジュールを使って要素を抜き出しています。

参考:
https://yuukou-exp.plus/handle-xlsx-with-python-intro/
https://qiita.com/sino20023/items/0314438d397240e56576
https://qiita.com/mriho/items/f82c66e7a232b6b37206
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
    text_list_fmt = [''.join(s.splitlines()) for s in text_list] # 改行が入っていると見づらいので改行を削除する

    return text_list_fmt


'''
■ .pptx ファイルに書き込まれているテキストをリストにして返す関数

前述の関数のパワーポイント版。
展開まではエクセル用と同じですが、パワーポイントの場合はスライドごとに xml ファイルが分かれているので
先にスライドごとのファイルパスを格納したリストを生成し、それをベースにしてテキストを格納したリストを作成しています。

パワーポイント版関数のテキストの検知はエクセル版と違い正規表現のパターンマッチを利用しています。
たぶんエクセル版と同じように子階層を掘っていくやり方でもできると思いますが、
階層がエクセルよりも深いようで記述が冗長になりそうだったのでこちらを採用しています。

パワーポイントの xml に記述されたテキストは空白や文字の種類を境に
テキストが細切れになっているため、ページごとに文字を結合しています。

参考:
https://qiita.com/kaityo256/items/2977d53e70bbffd4d601
'''
def make_pptx_text_list(src_file_path):

    # zip アーカイブのオブジェクトを生成
    zip = ZipFile(src_file_path)

    # ppt/slides/slideXX.xml をリストへ格納する
    slide_path_list = [
        zfp.filename for zfp in zip.filelist
        if zfp.filename.startswith('ppt/slides/slide')
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

前述の関数のワード版。処理の流れは前述の関数と同様です。

ワードの xml ファイルの場合は、ひとつの xml ファイルの中に細切れになったテキストが格納されているようなイメージ。
なのでパワーポイントと同様に正規表現のパターンマッチ + 文字の結合を行っています。

パワーポイント版のようにページごとにファイルが分けられているのではないため、ページごとの行番号の出力は不可能そう。
そういう情報が保存されていればできそうなので調べてみてもいいかも？
'''
def make_docx_text_list(src_file_path):

    # zip アーカイブのオブジェクトを生成
    zip = ZipFile(src_file_path)

    # xml ファイルをデコードして読み込む
    with zip.open('word/document.xml', 'r') as f:
        body = f.read().decode('utf-8')

    # パターンマッチでテキストが格納されている箇所を検知してリスト化する
    find_text_list = re.findall(r'\<w\:t\>.+?\<\/w\:t\>', body)
    find_text_list_fmt = [s.lstrip('<w:t>').rstrip('</w:t>') for s in find_text_list] # 左右のタグを削除する

    # テキストを結合してリストへ格納する
    text_list_fmt = []
    text_list_fmt.append(''.join(find_text_list_fmt))

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
            hit_num += 1
        line_num += 1
    return hit_num

'''
■ 渡された情報をそのままプリントする関数（例外用）
'''
def print_result_error(file_path, txt, error):
    return '%s : %s <%s>' % (file_path, txt, error)