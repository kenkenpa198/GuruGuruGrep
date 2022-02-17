'''
Python3.5 以降

参考サイト
https://note.nkmk.me/python-glob-usage/
https://blog.honjala.net/entry/2016/06/24/005852
'''

import glob
import os
import re

# 検索対象のディレクトリとテキストを指定
grep_dir = 'workbench/**'
grep_txt = 'aaa'

# 指定ディレクトリのファイルのみをリスト化
file_path_list = [p for p in glob.glob(grep_dir, recursive=True) if os.path.isfile(p)]
print(file_path_list)

# 検索処理
for file_path in file_path_list:
    root, ext = os.path.splitext(file_path)

    # 拡張子が Office 系の場合の条件分岐
    # TODO: この辺に Office 系ソフトの場合の条件分岐とか unzip 処理を追加する
    if ext in ('.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx'):
        continue

    # ファイルを開いて行ごとに検索
    with open(file_path, encoding='utf-8') as f:
        line_num = 1
        for line in f:
            m = re.search(grep_txt, line)
            if m:
                # 結果をプリントする
                # ファイルパス.txt (X行目, Y文字目) : 行のテキスト
                out = '%s (%d, %d) : %s' % (file_path, line_num, m.start() + 1, line.rstrip())
                print(out)
            line_num = line_num + 1