VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModelFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

' ファイルのパス
Private inner_SourcePath As ObjPath
' マッチング用の処理済みのファイル名
Private inner_AnalyzedName As String
' 元の名称から図面名のみを切り出した名称
Private inner_SnippedName As String
' 同一ファイルに格納している（その〇）に該当する番号
' 連番モードの場合はPart〇に該当する番号
Private inner_PaperNumbers As ObjSet

' ファイルモデルの初期設定
' 引数: sourcePath (ObjPath) - ファイルのパス情報
'       analyzedName (String) - マッチング用の処理済みのファイル名
'       snippedName (String) - 元の名称から図面名のみを切り出した名称
'       paperNumbers (ObjSet of String) - 同一ファイルに格納している（その〇）に該当する番号
' 戻り値: なし
Public Function InitModel(sourcePath As ObjPath, analyzedName As String, snippedName As String, paperNumbers As ObjSet)
  Set inner_SourcePath = sourcePath
  inner_AnalyzedName = analyzedName
  inner_SnippedName = snippedName
  Set inner_PaperNumbers = paperNumbers
End Function

' 元のファイル名（拡張子抜き）を返す
Public Property Get SourceName() As String
  SourceName = inner_SourcePath.Name
End Property

' ファイルの拡張子を返す
Public Property Get SourceExtension() As String
  SourceExtension = inner_SourcePath.Extension
End Property

' 解析によって取得したマッチング用のファイル名を返す
Public Property Get AnalyzedName() As String
  AnalyzedName = inner_AnalyzedName
End Property
Public Property Let AnalyzedName(value As String)
  inner_AnalyzedName = value
End Property

' 元の名称から図面名のみを切り出した名称を返す
Public Property Get SnippedName() As String
  SnippedName = inner_SnippedName
End Property

' 同一ファイルに格納している（その〇）に該当する番号のコレクションを返す
Public Property Get PaperNumbers() As ObjSet
  Set PaperNumbers = inner_PaperNumbers
End Property

' 連番モードのPart〇に該当する番号を返す
' ファイル名称の中に複数の番号が含まれている場合や番号が存在しなかった場合は空文字列を返す
Public Property Get PaperNumber() As Long
  If inner_PaperNumbers.Count = 1 Then
    PaperNumber = inner_PaperNumbers.ToList().At(0)
  Else
    PaperNumber = -1
  End If
End Property