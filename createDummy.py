from pathlib import Path

import pandas as pd


def setNewTestFiles(targetFolderName: str, names: pd.Series):
  """
  指定したフォルダにテストファイルを作成する
  """
  targetPath = Path(__file__).parent/targetFolderName

  # フォルダのリセット
  for fPath in targetPath.iterdir():
    fPath.unlink()

  # 空ファイルを生成
  names.map(lambda name: (targetPath/name).touch())


# テストケースの読み込み
loadTestCasePath = Path(__file__).parent / 'testCases' / 'by_山岸さん.xlsx'
db = pd.read_excel(loadTestCasePath, sheet_name='図面ファイル対応一覧', index_col=0)


# テストケースの作成
setNewTestFiles('PDF', db['旧PDFファイル名（対応がないファイルも書いてOK・空欄禁止）'])
setNewTestFiles('CAD', db['旧CADファイル名（対応がないファイルも書いてOK・空欄禁止）'])
