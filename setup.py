'''
■■■ setup.py ■■■

検索の設定を行うファイルです。
コメントを参考に記述を変更して保存後、ツールを実行してください。

参考:
基本的な正規表現一覧 | murashun.jp
https://murashun.jp/article/programming/regular-expression.html
'''


'''
■ 正規表現の検索設定
検索をする際に正規表現を使用するかどうかを設定します。
USE_REGEXP へ下記のいずれかの値を指定してください。

True : 正規表現を使用して検索する。
False: 正規表現を使用せずに検索する。

デフォルトは False 設定です。
'''
USE_REGEXP = False


'''
■ 検索ファイル設定
検索を限定したいファイルがある場合、 DETECT_PATH へ正規表現で指定してください。
マッチしたファイルのみを検索対象とします。

デフォルトは設定なしです。
除外パス設定に記述されたファイル以外を検索対象とします。

例1) 検索ファイルの設定をしない場合
DETECT_PATH = r''

例2) 拡張子が「.pptx」のファイルのみを検索したい場合
DETECT_PATH = r'\.pptx$'

例3) 拡張子が「.xlsx」もしくは「.txt」のファイルのみを検索したい場合
DETECT_PATH = r'\.xlsx$|\.txt$'
'''
DETECT_PATH = r''


'''
■ 除外ファイル設定
検索から除外したいファイルを EXCLUSION_PATH へ正規表現で指定してください。
マッチしたファイルは検索の対象から除外します。

デフォルトは読み込む必要のなさそうな拡張子を設定しています。

設定例は限定パス設定と同様です。
'''
EXCLUSION_PATH = r'\.doc$|\.xls$|\.ppt$|\.msi$|\.exe$|\.obj$|\.pdb$|\.ilk$|\.res$|\.pch$|\.iobj$|\.ipdb$'
