# バックエンド開発手引き

既存のコードである`vba-files`以下`base`, `objs`, `processer`, `utils`は実装が非常に煩雑になっており，各レイヤでの責任分担が不明瞭になっている

特にエラー処理やログ処理，単体テストを挟もうとしたときの設計上の課題が山積していることから，以下の標準仕様に基づいて既存コードのリファクタリングを行う．

## エラー処理

VBAはエラー表示が極めて不親切であるため，基本的には標準メソッド（`Dir()`や`Workbook.XXX()`等）を呼び出す際には，以下に開設するエラー補足機構でラップしたものを返すことを原則とする．

各APIはラップ済みのデータを返し，エラーがあった場合の対応はフロントエンドで定義するものとする


### エラー補足の概要

エラーを起こす可能性のある関数はすべて`Result`オブジェクトでラップされたデータを返すものとする．
`Result`オブジェクトの利用例は次の通りである．


```vb
' src/ 以下のバックエンドコードでの実装を想定
Function HogeProcess(args)
    Set resultReturn = HogeFunc(...)

    ' 処理が失敗した場合は失敗したことを返す
    ' 失敗した場合のフォールバックがある場合は関数を終了させなくても良い    
    If (Not resultReturn.isOK) Then
        Set HogeProcess = resultReturn
        Exit Function
    End If

    ' 処理が成功した場合は値を取り出して続きの処理を実装する
    successValue = resultReturn.value
End Function
```

フロントエンドにおいて，API呼び出しの結果エラーがあった場合に，ユーザーに通知する際の標準コードは次の通り．

```vb
' SheetN.cls での実装を想定
Sub OnClicked()
    ' こんなイベントの時はこの処理をしてほしい，といったものを関数渡しで定義する
    ' 用途としてプログレス画面の表示と更新を想定
    ' バックエンドの実装側で呼び出しを書いておいて，もし登録されていたら処理が走る格好
    e = New Event()
    e.anyEvent = Process

    ' エラーが起きる可能性のある処理
    Set apiRes = routes.HogeApi(e, ...)

    ' 処理した結果，ユーザーにエラー表示を返す必要がない場合は以下の実装は不要
    If (Not apiRes.isOK) Then
        ' Result.err はResultErrオブジェクト
        ' エラーの情報(表示すべきメッセージや求めるユーザーアクションの処理関数など)を持っているため，この情報に基づいてShowErrMsg関数の中で分岐する
        Call routes.ShowErrMsg(apiRes.err)
    End If
End Sub
```

エラーが起きるかもしれない処理を実行する際には，エラーを`Result`でラップできるようにしなければならない．

```vb
Function GetHoge()
    ' とある辞書型にキーを問い合わせてその値を取得することを考える．
    ' キーが存在しない場合にエラーを吐かれるため，これをResultで受ける．
    ' On Error句はVBAの標準処理の時にのみラップすることとし，Result型しか出てこないController層以上では使用しない
    On Error GoTo ErrHandler
    Set res = sampleDict("sampleKey")

    ' Failure(), Success() は標準モジュールで定義し，どこからでも呼び出せるようにする
    ' どちらを呼び出しても戻り値はResult型で統一されている
    ' 処理が成功した場合はSuccessに取得した値を登録する
    Set GetHoge = Success(res)
    Exit Function
ErrHandler:
        ' `config`に登録しておいたエラーキーの中から今回のエラー理由に適切なものを呼び出す．
        ' `config`に登録する際にメッセージ内で使いたい引数がある場合はそれも登録する
    Set GetHoge = Failure("errKey2", arg1, ...)
End Function
```

### エラーの新規登録

エラーはKey-Valueの静的データとして`config`層に保存しておく．
Keyはプログラム中でエラー時にResult型につけておくものであり，Valueに実際のメッセージが入っている構成とする．

新規登録は`config/errMsg/ja.cls`に保存されているKey-Valueの組を追加することで行う





## イベント処理

「このようなイベントが起きたときにはこういう処理を差し込みたい」といったことが事前に判明している場合は`Event`オブジェクトにイベント名と処理関数を入れておくことで，バックエンドから必要なタイミングで処理を呼び出すことができる．

例えば，プログレス画面の実装は次の要領で行う．


```vb
' SheetN.cls でプログレス画面が必要な処理を開始するボタン
Sub OnClicked()
    ' こんなイベントの時はこの処理をしてほしい，といったものを関数渡しで定義する
    ' 用途としてプログレス画面の表示と更新を想定
    ' バックエンドの実装側で呼び出しを書いておいて，もし登録されていたら処理が走る格好
    Set e = New EventsRepo()
    Set e.EM.Progress = routes.CallProgressWindow(...)

    ' プログレス画面の表示が必要な処理
    Call routes.HogeApi(e.EM, ...)
End Sub
```

```vb
' routes.HogeApi(args) が呼び出すcontroller
Function HogeController(em, ...)
    ' 内部的にProgressがない場合は，Startのダミーが呼び出される
    Call em.Progress.Start()

    ' 時間のかかる処理
    Call hogeService.LongRunningService(em.Progress.Update)

    ' プログレス表示の終了
    Call em.Progress.Close()
End Function

' routes.HogeApi(args) が呼び出すservice
Function LongRunningService(..., updater = Nothing)
    ' 進捗を表示可能な処理
    For i = 0 To N
        Call HogeFunc()
        If Not (updater Is Nothing) Then Call updater((i + 1) / N)
    Next i
End Function
```



## テスト計画

`repositories`層に対する単体テストは原則必須とし，`service`層に対しても可能な限りテストを付けて挙動の確認ができるようにする．

`service`層のテストの際に，`repositories`をモックに差替えられるような実装にできると良い．

テストは本来`xvba_unit_test`に記述すべきだが，現状では`vba-files/Module/Test/HogeTest.bas`に記述してよい．

命名規則は`TestHogeHoge.bas`とし，各ファイル内で作成したテスト関数をまとめて呼び出す`Public Sub RunAllXXXXTests()`を作成する．この`RunAllXXXXTests()`を`TestRunner.bas`の`RunAllTests()`に追記することでまとめて呼び出せるようにする．



## ログ取得

今回のプロジェクトではログの自動送信機能は現状実装しないが，将来的に実装する可能性がある．

ログで収集したい主な情報は以下の通り
- ユーザーがどのような図面に対してマッチング処理を走らせて，そのマッチング結果はどのように出力したか
- ユーザーがマッチング結果をどのように修正したか
- ユーザーに表示されたエラーの内容
- API層の呼び出し履歴

ログは`%HOMEPATH%\Rakubim`に`YYYY-MM-DD_hh:mm:ss.log`の形式で保存するものとする．
ここでログファイルに記載する日時は作成時点でのものを使用する．
ログファイルは作成日から１か月を超えたものは自動的に削除する仕様とする．

ログレベルは`INFO`, `WARN`, `ERROR`の３つ程度で良い．



## フォルダ構造

呼び出し順序に従って次の通りのフォルダ構成を採用する

### routes

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


### controllers

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

    ' 処理結果を書き込み
    sheetRepo.WriteColToCell(sheetName, startCell, coll)
End Function
```


### services

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


### repositories

シートや外部データと直接やり取りを行う。

ここ以外でシートや外部APIの構造を考えることがないようなデータに処理して返す

```vb
' SheetRepository.cls
Public Function ReadCol(arg1, arg2, ...)
    ' 変数定義（Dim x As X, Set x = New X）

    Set ReadCol = Renage2Collection(...)
End Function
```


### models

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


## 作成ファイルの命名規則

原則として作成するファイルは直前のフォルダ名を先頭につけて作成する．

これによりExcelのVBエディターでファイルを見たときに名前順が階層順になって見やすい

ex)
- `src/models`内のファイル：`ModelPaper.cls`, `ModelResult.cls`
- `src/services`内のファイル：`ServiceFile.cls`, `ServicePaper.cls`
- `Module/Test`内のファイル：`TestPathObj.bas`, `TestMatchingService.bas`