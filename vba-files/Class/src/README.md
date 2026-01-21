# ファイル構成

呼び出し順序に従って次の通りのフォルダ構成を採用する

> エラー処理は所定のKey（と付属情報）がついたエラーコードを返し、Keyに従ってエラーメッセージを表示する？
> 
> エラーによってユーザー判断の結果に応じて処理を振り分けることを加味して、付属情報には`isOK()`のような関数渡しができるModelの定義が必要 


## 実装上の注意
- `class_initializer()`は使用しない。（使用するとデバッグの際に`New <ClassName>`から先を追えなくなる）


## routes

フロントエンド（SheetX.cls）からのAPIエンドポイントを定義する

関数として定義し、その中身は呼び出すコントローラーの定義のみを記載する


```vb
' Routes.cls
' 目録入力シートに記載された図面群のマッチング結果リストを取得する
Public Function GetMatchingLists(arg1, arg2, ...)
    Dim matchingController As MatchingController
    Set matchingController = New MatchingController
    Set GetMatchingLists = matchingController.Run(arg1, arg2, ...)
End Function
```


## controllers

APIエンドポイントの要求に従って適切なサービスの呼び出し（と受け取ったデータの整形）を担当する。

```vb
' MatchingController.cls
' コントローラーの実行時処理を記載
Public Function Run(arg1, arg2, ...)
    ' 変数定義（Dim x As X, Set x = New X）
    
    ' 図面名称群を読み込み
    Set loadedPapers = paperService.LoadPapers(...)
    ' 処理対象ファイル群を読み込み
    Set loadedFiles = fileService.LoadFiles(...)
    ' マッチング結果を取得
    Set pdfMatchings = matchingService.MatchingPdfs(...)
    Set dwgMatchings = matchingService.MatchingDwgs(...)

    ' 処理結果をモデル化
    matchingResults.InitModel(pdfMatchings, dwgMatchings)

    Run = matchingResults
End Function
```


## services

ビジネスロジックの実装を担当する

様々な情報にアクセスして、システムが必要とする機能のコアを提供する

```vb
' PaperService.cls
' 渡されたファイル情報と図面情報のマッチングを行う
Public Function LoadPapers(arg1, arg2, ...)
    ' 変数定義（Dim x As X, Set x = New X）

    ' 入力された図面名称の生データ（String）を取得
    Set paperSourceNames = sheetRepository.ReadCol(...)
    ' 生データを必要な形式に処理した名称に変更
    Set analyzedNames = ClipAnalyzedNames(...)
    Set SnippedNames = ClipSnippedNames(...)

    ' PaperObjに変換
    Set paperObjs = GetPaperObjs(...)

    Set LoadPapers = paperObjs
End Function
```


## repositories

シートや外部データと直接やり取りを行う。

ここ以外でシートや外部APIの構造を考えることがないようなデータに処理して返す

```vb
' SheetRepository.cls
Public Function ReadCol(arg1, arg2, ...)
    ' 変数定義（Dim x As X, Set x = New X）

    Set ReadCol = Renage2Collection(...)
End Function
```


## models

データ型を定義する。

フロントエンドに返すデータや内部処理で使用する情報はすべてここで定義した型（オブジェクト）に登録したものしか利用しない。

入力データのバリデーションをかける場合はここで実装する
（ただし、Modelのバリデーションはあくまで「データとしての正当性」のみを担保するため、事前に処理できる場合はRepositoriesの中で処理しておくことが望ましい）
（よって、エラーメッセージも淡白なものでOK。リッチなメッセージはRepositoriesなどから発出するべき）

```vb
' PaperModel.cls
Private inner_SourceName As String
Private inner_AnalyzeName As String
Private inner_SnippedName As String

' （オブジェクトから参照できるデータはPropertyで定義する）
' 入力された図面名称そのまま
Public ProPerty Get SourceName() As String
    SourceName = inner_SourceName
End Property

' オブジェクトにデータを登録する
Public Function InitModel(sourceName, analyzeName, snippedName)
    vErr = Validate(...)
    if (vErr) Then
        InitModel = ...
        Exit Function
    End If

    inner_SourceName = sourceName
    inner_AnalyzeName = analyzeName
    inner_SnippedName = snippedname
End Function
```


## configs

静的な設定ファイル群を置いておく

```vb
' ErrMsgs/JaJp.cls
' ErrMsg.cls で表示言語やKeyがなかった場合の処理などを定義
' models/ErrKeys.cls でKeyの情報のみを列挙したモデルを作成？

' Key: Valueの形式になるようにメッセージを定義
```