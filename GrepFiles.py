'''
Python3.5 以降

参考サイト
https://note.nkmk.me/python-glob-usage/
https://blog.honjala.net/entry/2016/06/24/005852
'''

import glob
import os
import sys

import utils

print('\n========================')
print('       Grep Files       ')
print('========================')

try:

    # 除外拡張子リストを指定
    black_file_list = ['.doc', '.docx', '.xls', '.ppt', '.pptx']

    # 検索対象のディレクトリと検索するテキストを指定
    search_dir_input = input('\n検索対象のディレクトリパスを入力してください: ')
    search_dir = os.path.join(search_dir_input, '**')

    if not os.path.exists(search_dir_input):
        print('\n入力されたディレクトリが見つかりませんでした。パスが正しいか確認してください。')
        print('\nGrep Files を終了します。\n')
        sys.exit()

    search_txt = input('検索対象の文字列を入力してください: ')

    print('\n入力された情報で検索を実行します。')
    print('問題なければ Enter キーを押してください。')
    input('キャンセルする場合は Ctrl + C を押してください。')

    # 指定ディレクトリのファイルをリスト化
    file_path_list = [p for p in glob.glob(search_dir, recursive=True) if os.path.isfile(p)]

    print('\n▼ 検索結果')
    print('----------------------------------------------------------\n')
    # 検索処理
    for file_path in file_path_list:

        # 拡張子によって条件分岐するための前処理
        root, ext = os.path.splitext(file_path)

        try:
            # 除外拡張子リストに存在する場合は処理をスキップする
            if ext in black_file_list:
                continue

            # xlsx ファイルの場合の処理
            if ext == '.xlsx':
                f_list = utils.read_xlsx_text(file_path)
                utils.print_result(f_list, search_txt, file_path)
                continue

            # ファイルを開いて行ごとに検索
            with open(file_path, encoding='utf-8') as f:
                f_list = f.readlines()
                utils.print_result(f_list, search_txt, file_path)
                continue
        except UnicodeDecodeError as e:
            continue
            # print(utils.print_result_error(file_path, 'ファイルを読み込めませんでした。', e))

    print('\n----------------------------------------------------------')

except KeyboardInterrupt as e:
    print('\nキーボード入力により処理を中断しました。')

print('\nGrep Files を終了します。\n')