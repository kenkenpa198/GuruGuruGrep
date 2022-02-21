import glob
import os
import re

import setup
import utils

print('\n')
print('=========================')
print('       Grep Office       ')
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

    print('Enter キーを押すと検索を開始します。')
    print('\n※検索候補となるファイル数が多すぎると完了まで時間がかかる場合があります。')
    input('  中断する場合は Ctrl + C を押してください。')

    print('\n▼ 検索結果')
    print('----------------------------------------------------------\n')

    # インクリメント用の初期値を定義
    hit_num = 0                # 検索条件がヒットした件数
    search_file_num = 0        # 検索したファイルの総点数
    UnicodeDecodeError_num = 0 # 文字コードエラーで開けなかったファイルの総点数
    PermissionError_num = 0    # アクセス権限のエラーで開けなかったファイルの総点数

    # 検索処理
    # イテレータを利用してファイルパスの検知の度に検索を実行する
    for file_path in glob.iglob(search_dir, recursive=True):

        # ファイルパスの判定
        if (
            not os.path.isfile(file_path)                  # ファイルでなければ（ディレクトリだったら）処理をスキップ
            or not re.search(setup.DETECT_PATH, file_path) # 検索対象のファイルでなかったら処理をスキップ
            or re.search(setup.EXCLUDE_PATH, file_path)    # 除外対象のファイルだったら処理をスキップ
        ):
            continue

        search_file_num += 1

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

            # ここまでの処理で判定されなかった場合はファイルを開いて検索する
            with open(file_path, encoding='utf-8') as f:
                text_list = f.readlines()
                hit_num = utils.search_to_print_result(text_list, search_txt, file_path, hit_num)
                continue

        # ファイルが文字コードエラーで開けなかった場合はスキップする
        except UnicodeDecodeError as e:
            # utils.print_result_error(file_path, 'ファイルを読み込めませんでした。', e)
            UnicodeDecodeError_num += 1
            continue

        # アクセス権限のエラーで開けなかった場合はスキップする
        except PermissionError as e:
            PermissionError_num += 1
            print(f'ぱーみっしょんえらー!今 {PermissionError_num} 件だよ')
            continue

        # Ctrl + C で処理を中断した場合は for 文を終了する
        except KeyboardInterrupt as e:
            print('\nキーボード入力により処理を中断しました。\n中断した時点でのヒット件数を出力します。')
            break

    print('\n----------------------------------------------------------')

    print(f'{search_file_num} 点のファイル中 {hit_num} 件ヒットしました。')

    if UnicodeDecodeError_num:
        print(f'\n{search_file_num} 点のファイル中、文字コードのエラーにより開けなかったファイルが {UnicodeDecodeError_num} 点ありました。\n画像や動画などのバイナリファイルが存在した場合でもこのエラーが出ることがあります。')

    if PermissionError_num:
        print(f'\n{search_file_num} 点のファイル中、アクセス権限のエラーにより開けなかったファイルが {PermissionError_num} 点ありました。\n対象のファイルを開いていたなどの理由でもこのエラーが出ることがあります。')

except KeyboardInterrupt as e:
    print('\nキーボード入力により処理を中断しました。')

print('\nGrep Office を終了します。\n')
