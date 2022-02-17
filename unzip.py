'''
参考
https://qiita.com/sino20023/items/0314438d397240e56576
'''

import os
import re
import shutil
import xml.etree.ElementTree as ET


# ソースファイルの読み込み
src_dir = 'workbench'
src_file = 'Book1.xlsx'
src_file_path = os.path.join(src_dir, src_file)               # ソースファイルのパスを生成
src_file_root, src_file_ext = os.path.splitext(src_file)      # ソースファイルをファイル名と拡張子に分割
out_dir = os.path.join(src_dir, src_file_root)                # ソースの同ディレクトリ上にファイル名と同じディレクトリを生成
out_file_path = os.path.join(out_dir, src_file_root + '.zip') # 出力先のパスを生成

print(src_dir)
print(src_file)
print(src_file_path)
print(src_file_root)
print(src_file_ext)
print(out_dir)
print(out_file_path)

# zip ファイル化・展開
os.makedirs(out_dir, exist_ok=True)           # 出力先ディレクトリの作成
shutil.copy(src_file_path, out_file_path)     # ファイルのコピーと zip 化
shutil.unpack_archive(out_file_path, out_dir) # zip ファイルの展開


# 展開した XML ファイルの読み込み
xml_path = os.path.join(out_dir, 'xl/sharedStrings.xml')
tree = ET.parse(xml_path) # XML ファイルの読み込み
root = tree.getroot()     # 一番上の階層の要素を取り出す

# 子階層に保管されているテキストをリストへ格納
text_list = []
for child in root:
    for child_2 in child:
        if child_2.text:
            text_list.append(child_2.text)

print(text_list)