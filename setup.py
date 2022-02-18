'''
■ 検索ファイル設定
検知を限定したいファイルがある場合、 DETECT_PATH へ正規表現で指定してください。
マッチしたファイルのみを検索対象とします。
設定なしの場合は除外パス設定に記述されたファイル以外を検索対象とします。

例1) 検知ファイルの設定をしない場合
DETECT_PATH = r''

例2) 拡張子が「.pptx」のファイルのみを検索したい場合
DETECT_PATH = r'\.pptx$'

例3) 拡張子が「.xlsx」もしくは「.txt」のファイルのみを検索したい場合
DETECT_PATH = r'\.xlsx$|\.txt$'
'''
# DETECT_PATH = r''
DETECT_PATH = r'\.pptx$'
# DETECT_PATH = r'\.xlsx$'


'''
■ 除外ファイル設定
検知対象から除外したいファイルを EXCLUSION_PATH へ正規表現で指定してください。
マッチしたファイルは検知対象から除外します。

設定例は限定パス設定と同様です。
'''
EXCLUSION_PATH = r'\.doc$|\.xls$|\.ppt$|\.msi$|\.exe$|\.obj$|\.pdb$|\.ilk$|\.res$|\.pch$|\.iobj$|\.ipdb$'
