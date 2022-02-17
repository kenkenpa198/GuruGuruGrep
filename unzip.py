'''
参考
https://yuukou-exp.plus/handle-xlsx-with-python-intro/
https://qiita.com/sino20023/items/0314438d397240e56576
'''

import os
import xml.etree.ElementTree as ET
from zipfile import ZipFile, ZIP_DEFLATED, BadZipfile

# ソースファイルの読み込みとファイルパスの生成
src_dir = 'workbench'
src_file = 'Book1.xlsx'
src_file_path = os.path.join(src_dir, src_file)

# xlsx ファイルの中に存在する xml ファイルの読み込み
zip = ZipFile(src_file_path)                     # xlsx ファイルを zip ファイルとして読み込む
xml_path = zip.open('xl/sharedStrings.xml', 'r') # zip ファイルの中に存在する xml ファイルを読み込む
tree = ET.parse(xml_path)                        # xml ファイルをパースされたオブジェクトとして読み込む
root = tree.getroot()                            # 親階層の要素を取り出す

# 親階層の子階層に存在するテキストをリストへ格納する
text_list = []
for child in root:
    for child_2 in child:
        if child_2.text:
            text_list.append(child_2.text)

print(text_list)