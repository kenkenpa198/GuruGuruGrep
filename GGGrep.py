import argparse
import glob
import os
import re

import setup
import utils

import tqdm

# コマンドラインオプション
parser = argparse.ArgumentParser(description='a program is search for any files with grep.')
parser.add_argument('-d', '--directory_path', help='assign a search directory path')
parser.add_argument('-k', '--keyword', help='assign a keyword for search')
parser.add_argument('-r', '--regexp', action='store_true', help='enable search with regular expression')

args = parser.parse_args()

'''
■ コマンドライン引数の受け取り処理
'''

# -r を受け取った場合は正規表現検索をオンにする
if args.regexp:
    use_regexp_option = True
    print('正規表現検索を有効にするコマンドライン引数を受け取りました。')
else:
    use_regexp_option = False

# -d を受け取った場合は指定されたディレクトリを変数へ格納しておく
search_dir_input = None
if args.directory_path:
    print('検索対象のディレクトリパスをコマンドライン引数で受け取りました。')
    search_dir_input = args.directory_path

# -k を受け取った場合は検索キーワードを変数へ格納しておく
keyword = None
if args.keyword:
    print('検索キーワードをコマンドライン引数で受け取りました。')
    keyword = args.keyword

'''
■ GGGrep を起動 ～ 検索に利用する情報をインプットする処理
'''

print('\n========================')
print('      GuruGuruGrep       ')
print('========================\n')


try:
    # 検索対象ディレクトリ引数による指定がなかった場合は入力をリクエストする
    if not search_dir_input:
        search_dir_input = input('検索対象のディレクトリパスを入力してください: ')

    while not os.path.exists(search_dir_input):
        print('\n入力されたディレクトリが見つかりませんでした。パスが正しいか確認してください。')
        print('中断する場合は Ctrl + C を押してください。')
        search_dir_input = input('\n検索対象のディレクトリパスを入力してください: ')

    # 検索対象ディレクトリへサブディレクトリの指定を追加する
    search_dir = os.path.join(search_dir_input, '**')

    # 検索キーワード引数による指定がなかった場合は入力をリクエストする
    if not keyword:
        keyword = input('検索キーワードを入力してください: ')

    print('\n下記の検索条件で検索を実行します。')
    print('----------------------------------------------------------\n')

    print(f'検索ディレクトリ : {search_dir}')
    print(f'検索キーワード   : {keyword}')

    if setup.USE_REGEXP or use_regexp_option:
        print(f'正規表現で検索   : 使用する')
        use_regexp_flag = True
    else:
        print(f'正規表現で検索   : 使用しない')
        use_regexp_flag = False

    if setup.FILTER_PATH:
        print(f'検索の絞り込み   : {setup.FILTER_PATH}')
    else:
        print(f'検索の絞り込み   : 設定なし')

    if setup.EXCLUDE_PATH:
        print(f'検索から除外     : {setup.EXCLUDE_PATH}')
    else:
        print(f'検索から除外     : 設定なし')

    print('\n----------------------------------------------------------\n')

    print('Enter キーを押すと検索を開始します。')
    print('\n[!] 検索候補となるファイル数が多すぎると完了まで時間がかかる場合があります。')
    input('    中断する場合は Ctrl + C を押してください。')

    print('\n▼ 検索結果')
    print('----------------------------------------------------------\n')


    '''
    ■ メインとなる検索処理
    '''


    # 検索条件がヒットした件数を格納する合算用リストを定義
    hit_num_list = []

    # インクリメント用変数の初期値を定義
    search_file_num        = 0 # 検索対象のファイル総件数
    UnicodeDecodeError_num = 0 # 文字コードエラーで開けなかったファイルの総件数
    PermissionError_num    = 0 # アクセス権限のエラーで開けなかったファイルの総件数

    # Ctrl + C を受け取った時のフラグ用変数の初期値を定義
    KeyboardInterrupt_flag = False

    # イテレータを利用してファイルパスの検知の度に検索を実行する
    for file_path in tqdm.tqdm(glob.iglob(search_dir, recursive=True), desc='検索中……'):

        # ファイルパスの判定
        if (
            not os.path.isfile(file_path)                  # ファイルでなければ（ディレクトリだったら）処理をスキップ
            or not re.search(setup.FILTER_PATH, file_path) # 検索対象のファイルでなかったら処理をスキップ
            or re.search(setup.EXCLUDE_PATH, file_path)    # 除外対象のファイルだったら処理をスキップ
        ):
            continue

        search_file_num += 1

        # 拡張子によって条件分岐するための前処理
        root, ext = os.path.splitext(file_path)

        try:
            # .xlsx ファイルの場合の処理
            if ext == '.xlsxw':
                target_text_list = utils.make_xlsx_text_list(file_path)
                hit_num = utils.search_to_print_from_list(target_text_list, keyword, file_path)
                hit_num_list.append(hit_num)
                continue

            if ext == '.xlsx':
                hit_num = utils.search_to_print_from_xlsx(file_path, keyword)
                hit_num_list.append(hit_num)
                continue

            # .pptx ファイルの場合の処理
            if ext == '.pptx':
                hit_num = utils.search_to_print_from_pptx(file_path, keyword)
                hit_num_list.append(hit_num)
                continue

            # .docx ファイルの場合の処理
            if ext == '.docx':
                target_text_list = utils.make_docx_text_list(file_path)
                hit_num = utils.search_to_print_from_list(target_text_list, keyword, file_path)
                hit_num_list.append(hit_num)
                continue

            # .pdf ファイルの場合の処理
            if ext == '.pdf':
                target_text_list = utils.make_pdf_text_list(file_path)
                hit_num = utils.search_to_print_from_list(target_text_list, keyword, file_path)
                hit_num_list.append(hit_num)
                continue

            # ここまでの処理で判定されなかった場合はファイルを開いて検索する
            with open(file_path, encoding='utf-8') as f:
                target_text_list = f.readlines()
                hit_num = utils.search_to_print_from_list(target_text_list, keyword, file_path)
                hit_num_list.append(hit_num)
                continue

        # ファイルが文字コードエラーで開けなかった場合はスキップする
        except UnicodeDecodeError as e:
            # utils.print_result_error(file_path, 'ファイルを読み込めませんでした。', e)
            UnicodeDecodeError_num += 1
            continue

        # アクセス権限のエラーで開けなかった場合はスキップする
        except PermissionError as e:
            PermissionError_num += 1
            continue

        # Ctrl + C で処理を中断した場合は for 文を終了する
        except KeyboardInterrupt as e:
            KeyboardInterrupt_flag = True
            break

    if KeyboardInterrupt_flag:
        print('\nキーボード入力により処理を中断しました。\n中断した時点での結果を出力します。')
    else:
        print('検索を完了しました。')

    print('\n----------------------------------------------------------')

    print(f'{sum(hit_num_list)} 件ヒットしました。')

    print(f'\n検索対象のファイル総件数                               : {search_file_num} 件')

    if UnicodeDecodeError_num:
        print(f'総件数のうち開けなかったファイル（文字コードエラー）   : {UnicodeDecodeError_num} 件')

    if PermissionError_num:
        print(f'総件数のうち開けなかったファイル（アクセス権限エラー） : {PermissionError_num} 件')

    export_flag = input('\n終了する場合は Enter キーを押してください。\n検索結果を CSV ファイルとして出力する場合は C を送信してください: ')

    if export_flag in ('c', 'C', 'ｃ', 'Ｃ'):
        export_dir = os.path.join(os.getcwd(), 'export')

        export_input_path  = utils.export_input_data(export_dir, search_dir, keyword, use_regexp_flag, setup.FILTER_PATH, setup.EXCLUDE_PATH)
        export_result_path = utils.export_result_csv(export_dir, utils.result_multi_list)

        print(f'\n検索結果を以下のファイルへ出力しました。')
        print(f'入力情報: {export_input_path}')
        print(f'検索結果: {export_result_path}')

except KeyboardInterrupt as e:
    print('\nキーボード入力により処理を中断しました。')

print('\nGuruGuruGrep を終了します。\n')
