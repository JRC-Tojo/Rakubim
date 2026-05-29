VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RepoErrMsgs_JaJp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

' エラーキーに対応するメッセージ群
Public inner_Msgs As ObjDict

Private Sub class_initialize()
  Set inner_Msgs = New ObjDict
  With inner_Msgs
    .Append "repo.sheet.readcol.error", _
      Array("シート読み取りエラー", "{}シート{}列の読み取りに失敗しました．\n(エラーコード：{}, {})")
    .Append "repo.sheet.readcell.empty", _
      Array("セル読み取りエラー", "{}シート{}が空欄です．\n値を入力してください．")
    .Append "repo.sheet.readcell.error", _
      Array("セル読み取りエラー", "{}シート{}の読み取りに失敗しました．\n(エラーコード：{}, {})")
    .Append "repo.sheet.writecol.error", _
      Array("シート書き込みエラー", "{}シート{}列の書き込みに失敗しました．\n(エラーコード：{}, {})")
    .Append "repo.sheet.getlastrow.error", _
      Array("最終データ読み取りエラー", "{}シート{}列の末端データ位置の取得に失敗しました．\n(エラーコード：{}, {})")
    .Append "repo.sheet.clearcells.error", _
      Array("セル削除エラー", "{}シート{}のデータ削除に失敗しました．\n(エラーコード：{}, {})")
    .Append "repo.sheet.export.error", _
      Array("エクスポートエラー", "以下のファイルへのエクスポートに失敗しました．\n{}\n\n(エラーコード：{}, {})")
    .Append "repo.sheet.import.error", _
      Array("インポートエラー", "以下のファイルからのインポートに失敗しました．\n{}\n\n(エラーコード：{}, {})")
    .Append "repo.sheet.registprintarea.error", _
      Array("印刷範囲設定エラー", "{}シートの印刷範囲の設定に失敗しました．\n(エラーコード：{}, {})")
    
    .Append "service.file.load.folder_not_found", _
      Array("フォルダエラー", "以下のパスに示すフォルダが見つかりません．\n{}")
    .Append "service.file.load.empty_folder", _
      Array("フォルダエラー", "以下のパスに示すフォルダにファイルが入っていません．\n{}")
    .Append "service.file.newname.missmatch", _
      Array("新ファイル名エラー", "図面名とPDF、DWGファイルの対応がすべての図面に対して定義されていません．\n目録入力シートのPDFファイルとDWGファイルの列をすべて埋めてください．")
    .Append "service.file.newname.company", _
      Array("新ファイル名エラー", "業務機関名が入力されていません．\n入力シートの業務機関コードのセルを選択し，契約先の業務機関をプルダウンメニューから選択してください．")
    .Append "service.file.newname.project", _
      Array("新ファイル名エラー", "契約番号が入力されていません．\n入力シートの契約番号セルに契約番号となる5桁の数字を入力してください．")
    .Append "service.file.newname.structure", _
      Array("新ファイル名エラー", "構造物通し番号が入力されていません．\n入力シートの構造物通し番号のセルを選択して番号を入力してください．")
    .Append "service.file.rename.missmatch", _
      Array("リネームエラー", "旧ファイル名に対応する新ファイル名がすべての項目に対して定義されていません．\nファイル名変更シートの色が塗られている箇所に新ファイル名を定義してください．")
    .Append "service.file.rename.failed", _
      Array("リネームエラー", "以下のファイルのリネームに失敗しました．\n成果品フォルダが正しく設定されていることを確認し、問題が解決しない場合はご自身のデスクトップ上に成果品ファイルをダウンロードしてお試しください．\n\n【リネームに失敗したファイル一覧】\n{}")
      
    .Append "service.match.noFiles", _
      Array("マッチングエラー", "マッチング対象のファイルが見つかりません．")
    .Append "service.match.error", _
      Array("マッチングエラー", "未知のエラーでマッチング処理に失敗しました．\n(エラーコード：{}, {})")
      
    .Append "repo.nameanalyzer.snippedname.empty", _
      Array("名前解析エラー", "ファイル名「{}」から図面名称を取得できませんでした．")
    .Append "repo.nameanalyzer.analyzename.error", _
      Array("名前解析エラー", "ファイル名「{}」の解析中に未知のエラーが発生しました．\n(エラーコード：{}, {})")
      
    .Append "service.namecleaner.clean.error", _
      Array("図面名正規化エラー", "未知のエラーが発生しました．\n(エラーコード：{}, {})")
      
    .Append "service.paper.build.emptySheet", _
      Array("図面名エラー", "図面番号・図面名を入力してください．")
    .Append "service.paper.build.mismatch", _
      Array("図面名エラー", "図面番号・図面名を入力数が一致しません．\n図面番号・図面名の列に空欄が含まれていないことを確認してください．")
    .Append "service.paper.build.duplicated", _
      Array("図面名エラー", "目録入力シート内の図面名称に重複が含まれます．\n同一名称の図面は処理できないため，図面名称を変更して再度お試しください．\n\n【重複が確認された図面名称】\n・{}")
    
    .Append "service.sheet.pulldown.files", _
      Array("入力エラー", "フォルダ内に存在しないファイル名は入力できません．")
    .Append "service.sheet.newname.duplicated", _
      Array("入力エラー", "新しいファイル名に重複が含まれます．\n同一名称のファイルは生成できないため，新しいファイル名称に変更して再度お試しください．\n\n【重複が確認された新しいファイル名称】\n・{}")
    .Append "service.sheet.old2new.missmatch", .At("service.file.rename.missmatch")
    .Append "service.sheet.relationaltable.missmatch", .At("service.file.newname.missmatch")
    .Append "service.sheet.relationaltable.lostrename", _
      Array("入力エラー", "以下の元ファイル名に対応するリネームが行われていません．\nファイル名変更シートで新ファイル名の生成，適用を再度お試しください．\n\n{}")
  End With
End Sub

' メッセージウィンドウのタイトル
Public Property Get MsgboxTitle() As String
  MsgboxTitle = "らくびむ"
End Property