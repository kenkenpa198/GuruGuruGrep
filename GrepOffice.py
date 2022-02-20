import glob
import os
import re
import sys

import setup
import utils

print('\n')
print('=========================')
print('     🔍 Grep Office     ')
print('=========================')

try:
    # 検索対象のディレクトリと検索するテキストを指定
    search_dir_input = input('\n検索対象のディレクトリパスを入力してください: ')

    while not os.path.exists(search_dir_input):
        print('\n入力されたディレクトリが見つかりませんでした。パスが正しいか確認してください。')
        print('中断する場合は Ctrl + C を押してください。')
        search_dir_input = input('\n検索対象のディレクトリパスを入力してください: ')

    search_dir = os.path.join(search_dir_input, '**')

    search_txt = input('検索条件とする文字列を入力してください: ')

    print('\n下記の検索条件で検索を実行します。')
    print('----------------------------------------------------------\n')

    print('検索ディレクトリ : ' + search_dir)
    print('検索条件         : ' + search_txt)

    if setup.USE_REGEXP:
        print('正規表現で検索   : 使用する')
    else:
        print('正規表現で検索   : 使用しない')

    if setup.DETECT_PATH:
        print('検索ファイル指定 : ' + setup.DETECT_PATH)
    else:
        print('検索ファイル指定 : 設定なし')

    if setup.EXCLUDE_PATH:
        print('除外ファイル指定 : ' + setup.EXCLUDE_PATH)
    else:
        print('除外ファイル指定 : 設定なし')

    if setup.DETECT_PDF_FILE:
        print('PDF ファイル     : 検索する')
    else:
        print('PDF ファイル     : 検索しない')

    print('\n----------------------------------------------------------\n')

    print('問題なければ Enter キーを押してください。')
    input('中断する場合は Ctrl + C を押してください。')

    # 指定ディレクトリ以下に存在するファイルをリストへ格納する
    file_path_list = [
        p for p in glob.glob(search_dir, recursive=True) # 指定ディレクトリ以下に存在するファイルパスを再帰的に格納する
        if re.search(setup.DETECT_PATH, p)               # 検索対象のファイルのみを格納する
        if not re.search(setup.EXCLUDE_PATH, p)          # 除外対象のファイルは格納しない
        if os.path.isfile(p)                             # 存在するファイルのみを格納する
    ]

    # for i in file_path_list:
    #     print(i)

    print('\n▼ 検索結果')
    print('----------------------------------------------------------\n')

    # 検索処理
    hit_num = 0
    for file_path in file_path_list:

        # 拡張子によって条件分岐するための前処理
        root, ext = os.path.splitext(file_path)

        try:
            # .xlsx ファイルの場合の処理
            if ext == '.xlsx':
                text_list = utils.make_xlsx_text_list(file_path)
                hit_num = utils.search_to_print_result(text_list, search_txt, file_path, hit_num)
                continue

            # .pptx ファイルの場合の処理
            if ext == '.pptx':
                text_list = utils.make_pptx_text_list(file_path)
                hit_num = utils.search_to_print_result(text_list, search_txt, file_path, hit_num)
                continue

            # .docx ファイルの場合の処理
            if ext == '.docx':
                text_list = utils.make_docx_text_list(file_path)
                hit_num = utils.search_to_print_result(text_list, search_txt, file_path, hit_num)
                continue

            # .pdf ファイルの場合の処理（PDF の検索設定が True の場合のみ処理する）
            if setup.DETECT_PDF_FILE:
                if ext == '.pdf':
                    text_list = utils.make_pdf_text_list(file_path)
                    hit_num = utils.search_to_print_result(text_list, search_txt, file_path, hit_num)
                    continue

            # Office 系でなかった場合はファイルを開いて検索する
            with open(file_path, encoding='utf-8') as f:
                text_list = f.readlines()
                hit_num = utils.search_to_print_result(text_list, search_txt, file_path, hit_num)
                continue

        # ファイルが文字コードエラーで開けなかった場合はスキップする
        except UnicodeDecodeError as e:
            # utils.print_result_error(file_path, 'ファイルを読み込めませんでした。', e)
            continue

    print('\n----------------------------------------------------------')

    print(f'{hit_num} 件ヒットしました。')

except KeyboardInterrupt as e:
    print('\nキーボード入力により処理を中断しました。')

print('\nGrep Office を終了します。\n')
