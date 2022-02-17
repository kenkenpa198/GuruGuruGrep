import os
import shutil

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

# 出力先ディレクトリの作成
os.makedirs(out_dir, exist_ok=True)

# ファイルのコピーと zip 化
shutil.copy(src_file_path, out_file_path)

# zip ファイルの展開
shutil.unpack_archive(out_file_path, out_dir)
