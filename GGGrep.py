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
■ コマンドライン引数と設定ファイルからの受け取り処理

コマンドライン引数による指定を優先する。
指定がなかった場合は設定ファイルを読み込んだり後続の処理で入力をリクエストする。
'''

# -d を受け取った場合は指定されたディレクトリを変数へ格納しておく
if args.directory_path:
    print('検索対象のディレクトリパスを受け取りました。')
    search_dir_input = args.directory_path
else:
    search_dir_input = None

# -k を受け取った場合は検索キーワードを変数へ格納しておく
if args.keyword:
    print('検索キーワードを受け取りました。')
    keyword = args.keyword
else:
    keyword = None

# -r を受け取った場合は正規表現検索をオンにする
if args.regexp:
    print('正規表現での検索を有効にしました。')
    regexp_flag = True
# -r 引数による正規表現を有効にする操作がなかった場合は設定ファイルから読み込む
else:
    regexp_flag = setup.USE_REGEXP


'''
■ GGGrep を起動 ～ 検索に利用する情報をリクエスト・表示する処理
'''

print('\n========================')
print('      GuruGuruGrep       ')
print('========================\n')


try:
    run_search_flag = True
    while run_search_flag:

        # 検索ディレクトリパス指定がなかった場合は入力をリクエストする
        if not search_dir_input:
            search_dir_input = input('検索対象のディレクトリパスを入力してください: ')

        # ディレクトリパスが存在するかチェック
        while not os.path.exists(search_dir_input):
            print('\n入力されたディレクトリが見つかりませんでした。パスが正しいか確認してください。')
            print('中断する場合は Ctrl + C キー を押してください。')
            search_dir_input = input('\n検索対象のディレクトリパスを入力してください: ')

        # 検索対象ディレクトリへサブディレクトリの指定を追加する
        search_dir = os.path.join(search_dir_input, '**')


        # 検索キーワード指定がなかった場合は入力をリクエストする
        if not keyword:
            keyword = input('検索キーワードを入力してください: ')


        # 検索条件の表示
        print('\n▼検索条件')
        print('----------------------------------------------------------\n')

        print(f'検索ディレクトリ : {search_dir}')
        print(f'検索キーワード   : {keyword}')

        if regexp_flag:
            print(f'正規表現で検索   : 使用する')
        else:
            print(f'正規表現で検索   : 使用しない')

        if setup.FILTER_PATH:
            print(f'検索の絞り込み   : {setup.FILTER_PATH}')
        else:
            print(f'検索の絞り込み   : 設定なし')

        if setup.EXCLUDE_PATH:
            print(f'検索から除外     : {setup.EXCLUDE_PATH}')
        else:
            print(f'検索から除外     : 設定なし')

        print('\n----------------------------------------------------------')

        print('上記の情報で検索を行います。')
        print('\n問題なければ何も入力せずに Enter キーを押してください。')
        print('設定を変更したい場合は下記のコマンドを送信してください。\n')
        print('  d : 検索対象のディレクトリパスを再入力する')
        print('  k : 検索キーワードを再入力する')
        print('  r : 正規表現の検索設定を切り替える')

        reset_cmd = input('\n: ')

        if reset_cmd in ('r', 'R', 'ｒ', 'Ｒ'):
            if regexp_flag:
                regexp_flag = False
                input('\n正規表現検索を「使用しない」に変更しました。\nEnter キーを押すともう一度検索条件を表示します。')
            else:
                regexp_flag = True
                input('\n正規表現検索を「使用する」に変更しました。\nEnter キーを押すともう一度検索条件を表示します。')
            continue

        if reset_cmd in ('d', 'D', 'ｄ', 'Ｄ'):
            search_dir_input = None
            print('')
            continue

        if reset_cmd in ('k', 'K', 'ｋ', 'Ｋ'):
            keyword = None
            print('')
            continue

        if reset_cmd:
            input('\n指定外のコマンドが入力されました。\nEnter キーを押すともう一度検索条件を表示します。')
            continue

        if not reset_cmd:
            break

    print('\n検索条件を確定しました。\nEnter キーを押すと検索を開始します。')
    print('\n[!] 検索候補となるファイル数が多すぎると完了まで時間がかかる場合があります。')
    input('    検索の途中で中断する場合は Ctrl + C キー を押してください。\n: ')


    '''
    ■ メインとなる検索処理
    '''

    print('\n▼ 検索結果')
    print('----------------------------------------------------------\n')

    try:
        # 検索条件がヒットした件数を格納する合算用リストを定義
        hit_num_list = []

        # インクリメント用変数の初期値を定義
        search_file_num        = 0 # 検索対象のファイル総件数
        UnicodeDecodeError_num = 0 # 文字コードエラーで開けなかったファイルの総件数
        PermissionError_num    = 0 # アクセス権限のエラーで開けなかったファイルの総件数

        # Ctrl + C キー を受け取った時のフラグ用変数の初期値を定義
        KeyboardInterrupt_flag = False

        # イテレータを利用してファイルパスの検知の度に検索を実行する
        for file_path in tqdm.tqdm(glob.iglob(search_dir, recursive=True), desc='検索中……'):

            # ファイルパスの判定
            if (
                not os.path.isfile(file_path)                  # ファイルでなければ（ディレクトリだったら）処理をスキップする
                or not re.search(setup.FILTER_PATH, file_path) # 絞り込み対象のファイルでなかったら処理をスキップする 空の場合は偽と評価され判定を通る
                or re.search(setup.EXCLUDE_PATH, file_path)    # 除外対象のファイルだったら処理をスキップする
            ):
                continue

            # ファイルパスの判定を通ったら検索対象のファイル総件数をインクリメント
            search_file_num += 1

            # 拡張子によって条件分岐するための前処理
            root, ext = os.path.splitext(file_path)

            try:
                # .xlsx ファイルの場合の処理（旧）
                # TODO: 検証問題なかったら削除する
                if ext == '.xlsxw':
                    target_text_list = utils.make_xlsx_text_list(file_path)
                    hit_num = utils.search_to_print_from_list(target_text_list, keyword, regexp_flag, file_path)
                    hit_num_list.append(hit_num)
                    continue

                # .xlsx ファイルの場合の処理
                if ext == '.xlsx':
                    hit_num = utils.search_to_print_from_xlsx(file_path, keyword, regexp_flag)
                    hit_num_list.append(hit_num)
                    continue

                # .pptx ファイルの場合の処理
                if ext == '.pptx':
                    hit_num = utils.search_to_print_from_pptx(file_path, keyword, regexp_flag)
                    hit_num_list.append(hit_num)
                    continue

                # .docx ファイルの場合の処理
                if ext == '.docx':
                    target_text_list = utils.make_docx_text_list(file_path)
                    hit_num = utils.search_to_print_from_list(target_text_list, keyword, regexp_flag, file_path)
                    hit_num_list.append(hit_num)
                    continue

                # .pdf ファイルの場合の処理
                if ext == '.pdf':
                    target_text_list = utils.make_pdf_text_list(file_path)
                    hit_num = utils.search_to_print_from_list(target_text_list, keyword, regexp_flag, file_path)
                    hit_num_list.append(hit_num)
                    continue

                # ここまでの処理で判定されなかった場合はファイルを open() で開いて検索する
                with open(file_path, encoding='utf-8') as f:
                    target_text_list = f.readlines()
                    hit_num = utils.search_to_print_from_list(target_text_list, keyword, regexp_flag, file_path)
                    hit_num_list.append(hit_num)
                    continue

            # ファイルが文字コードエラーで開けなかった場合はスキップする
            except UnicodeDecodeError as e:
                UnicodeDecodeError_num += 1
                continue

            # アクセス権限のエラーで開けなかった場合はスキップする
            except PermissionError as e:
                PermissionError_num += 1
                continue


    # Ctrl + C キー で処理を中断した場合はメッセージを変更する
    # わざわざ条件分岐させなくてもいいけど、tqdm のプログレス表示とごちゃついて見づらかったので実装している
    except KeyboardInterrupt as e:
        KeyboardInterrupt_flag = True
    if KeyboardInterrupt_flag:
        print('\nキーボード入力により処理を中断しました。\n中断した時点での結果を出力します。')
    else:
        print('検索を完了しました。')

    print('\n----------------------------------------------------------')
    print(f'{sum(hit_num_list)} 件ヒットしました。')


    '''
    ■ 補足情報の表示とファイル出力処理
    '''

    # 検索したファイルの情報を表示
    print(f'\n検索対象のファイル総件数                               : {search_file_num} 件')
    if UnicodeDecodeError_num:
        print(f'総件数のうち開けなかったファイル（文字コードエラー）   : {UnicodeDecodeError_num} 件')
    if PermissionError_num:
        print(f'総件数のうち開けなかったファイル（アクセス権限エラー） : {PermissionError_num} 件')

    export_flag = input('\n終了する場合は Enter キーを押してください。\n検索結果を CSV ファイルとして出力する場合は c を送信してください\n: ')


    # 検索結果を出力する処理
    if export_flag in ('c', 'C', 'ｃ', 'Ｃ'):
        export_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'export')

        export_input_path  = utils.export_input_data(export_dir, search_dir, keyword, regexp_flag, setup.FILTER_PATH, setup.EXCLUDE_PATH)
        export_result_path = utils.export_result_csv(export_dir, utils.result_multi_list)

        print(f'\n検索結果を以下のファイルへ出力しました。')
        print(f'検索条件: {export_input_path}')
        print(f'検索結果: {export_result_path}')


except KeyboardInterrupt as e:
    print('\nキーボード入力により処理を中断しました。')

print('\nGuruGuruGrep を終了します。\n')
