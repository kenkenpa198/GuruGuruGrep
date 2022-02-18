'''
参考
https://yuukou-exp.plus/handle-xlsx-with-python-intro/
https://qiita.com/sino20023/items/0314438d397240e56576
'''

import re
import xml.etree.ElementTree as ET
from zipfile import ZipFile


# xlsx ファイルに書き込まれているテキストをリストにして返す関数
def read_xlsx_text(src_file_path):

    # xlsx ファイルの中に存在する xml ファイルの読み込み
    zip = ZipFile(src_file_path)                     # ファイルを zip ファイルとして読み込む
    xml_path = zip.open('xl/sharedStrings.xml', 'r') # zip ファイルの中に存在する xml ファイルを読み込む
    tree = ET.parse(xml_path)                        # xml ファイルをパースされたオブジェクトとして読み込む
    root = tree.getroot()                            # 親階層の要素を取り出す

    # 親階層の子階層に存在するテキストをリストへ格納する
    text_list = []
    for child in root:
        for child_2 in child:
            if child_2.text:
                text_list.append(child_2.text)
    return text_list


# pptx ファイルに書き込まれているテキストをリストにして返す関数
# TODO: つくりかけ。
def read_pptx_text(src_file_path):

    # pptx ファイルの中に存在する xml ファイルの読み込み
    zip = ZipFile(src_file_path)                       # ファイルを zip ファイルとして読み込む
    xml_path = zip.open('ppt/slides/slide2.xml', 'r') # zip ファイルの中に存在する xml ファイルを読み込む
    body = xml_path.read()
    body = body.decode('utf-8')
    print(body)

    # パターンマッチでテキストが格納されている箇所を検知してリスト化する
    text_list = re.findall(r'\<a\:t\>.+?\<\/a\:t\>', body)
    print(text_list)
    return text_list



# 検索して結果をプリントする関数
def print_result(src_list, search_txt, file_path):
    line_num = 1
    for line in src_list:
        m = re.search(search_txt, line)
        if m:
            # ファイルパス.txt (X行目, Y文字目) : 行のテキスト
            out = '%s (%d, %d) : %s' % (file_path, line_num, m.start() + 1, line.rstrip())
            print(out)
        line_num = line_num + 1


# 渡された内容をプリントする関数（例外用）
def print_result_error(file_path, txt, error):
    return '%s : %s <%s>' % (file_path, txt, error)