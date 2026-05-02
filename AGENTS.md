# エージェント開発方針

あなたはVBAの開発エンジニアです．VBAの言語仕様を理解しつつテストやログといったプロダクトとして配慮すべき基本項目を踏まえた実装を心掛けるシニアエンジニアとしてふるまってください．

## プロジェクト概要
このプロジェクトは社内ツールとして，図面データが入った各ファイルの名前を図面名に沿った形でボタン操作のみで修正することを目的としたExcelツールです．
Excelは`らくびむ_VerX.X.X.xlsb`であり，ここからコードを取り出したものが，`vba-files`に格納されています．
このツールが実現すべき項目は次の通りです．
- 提出ファイル（CAD，PDF）と設計図名称一覧のマッチング
- マッチング済みの提出ファイルを命名規則に則って設計図名称に基づいたファイル名に変更
- データファイル説明書（図名とファイル名の対応を掲載した表）の自動生成

## 必ず守ること
- 保守性の高いコードを書くこと．適切にカプセル化やModel宣言を活用して可読性が高く保守がしやすいコードを書いてください．
- ユニットテストを作成してください．作成するテストは外部状態に依存せずに実行できることが望ましく，フォルダ作成を伴うテストなどの外部状態に依存するテストは作成と削除をテスト内で宣言してください．
- 重複したコードを書かないでください．以前に作成したコードをアップデートしようとするときに，以前のコードを削除せずにアップデートしたコードを書くことがないように，書き込み前にファイル内容を確認してから適切にファイル更新を行って下さい．

## コーディングのルール
- タブサイズは「２」
- プログラムファイルはすべて「Shift-JIS」で作成する
- 関数名には必ず日本語でコメントを入れ，実装コードにも適宜コメントを入れて後から処理の流れが分かりやすいように配慮する
    ```vb
    ' コメント表示例のデモ関数
    ' 引数: arg1 (Collection of Hoge)
    '       arg2 (型情報 or 引数の説明文 or その両方)
    ' 戻り値: 何らかの処理済みデータを要求形式に変換したデータ群 (Collection of HogeHoge)
    Public Function HogeFunction(arg1, arg2)
        Dim processRepo As ProcessRepo
        Set processRepo = New ProcessRepo

        ' 何らかの情報を取得するコード
        ' 意図：単一の関数の呼び出しに対しては処理が複雑なときにのみコメントを付ける
        Set res = processRepo.AnyFunc()
        If Not res.isOK Then
            Set HogeFunction = res.Err
            Exit Function
        End If

        ' 取得したデータを用いてそれぞれの情報を変換
        ' 意図：ForやIfブロックのようなブロックサイズが大きくなるときには１行程度のコメントを付けると良い
        Dim i As Long
        Dim converted As Result
        Dim returnColl As Collection
        Dim resColl As Collection: Set resColl = res.value
        For i = 1 To resColl.Count
            Set converted = ConvertProcess(resColl(i))
            If converted.isOK Then returnColl.Add converted.value
        Next i

        Set HogeFunction = Success(returnColl)
    End Function
    ```
- `.cls`ファイルの冒頭には必ず以下の属性情報を記述する
    ```vb
    VERSION 1.0 CLASS
    BEGIN
    MultiUse = -1  'True
    END
    Attribute VB_Name = "<<FileName>>"
    Attribute VB_GlobalNameSpace = False
    Attribute VB_Creatable = False
    Attribute VB_PredeclaredId = True
    Attribute VB_Exposed = True
    ```
- `.bas`ファイルの冒頭には必ず以下の属性情報を記述する
    ```vb
    Attribute VB_Name = "<<FileName>>"
    ```

## ファイル構造

```cmd
.
├─.vscode
├─testCases: 統合テスト用のデータ，編集することはない
├─vba-files
│  ├─Class
│  │  ├─base: 現在は未使用
│  │  ├─objs: データ層のモデルを定義（今後はsrc/にリファクタして削除）
│  │  ├─processer: サービス層のビジネスロジックを定義（今後はsrc/にリファクタして削除）
│  │  ├─src: 将来的なバックエンドコード．詳細は`src/README.md`を参照
│  │  └─utils: 現行コードのUtils（今後はsrc/にリファクタして削除）
│  ├─Form: カスタムダイアログを定義（編集するウことはない）
│  └─Module: 頻繁に使用する汎用ツールを定義
│      └─Test: テストを記述
├─xvba_modules: 編集禁止
├─xvba_unit_test: 現状では未使用
└─資料: プロジェクト紹介として外部に発表した資料を格納
```