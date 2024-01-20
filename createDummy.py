
from pathlib import Path
import shutil

FOLDER_NAME = ['CAD', 'PDF']

for folder in FOLDER_NAME:
  # フォルダを削除して掃除
  shutil.rmtree(Path(__file__).parent/folder)
  # 新しく再生成
  (Path(__file__).parent/folder).mkdir()
  # ファイルの名前をひとつづつ取得&作成
  for path in (Path(__file__).parent/'secret'/folder).iterdir():
    (Path(__file__).parent/folder/path.name).touch()