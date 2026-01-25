VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "YearConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' 和暦検索時に現在から遡る対象の年数
' `TargetYearWindow`で指定した数字よりも短い年号が生じると，バグになる
Private Const TargetYearWindow = 10

' 年号の開始日をDateとして定義
Private Const ReiwaStartDate As Date = #5/1/2019#
Private Const HeiseiStartDate As Date = #1/8/1989#
Private Const ShowaStartDate As Date = #12/25/1926#

' 和暦の数字から西暦を取得する
' 取得に失敗した場合は-1を返す
Public Function GetSeireki(yearNum As Integer) As Integer
  ' 西暦が渡された場合はそのまま返す
  If (IsSeireki(yearNum)) Then
    GetSeireki = yearNum
    Exit Function
  End If

  ' 現在の年号を取得する
  Dim nowWareki As YearObj
  Set nowWareki = GetNowWareki()

  ' 数字に対応する年号の辞書を作成
  ' 現在から過去 `TargetYearWindow` 年分を作成する
  ' ex) 3: 令和３年, 2: 令和２年, 1: 令和元年，31: 平成31年, ...
  Dim targetYearDict As Dictionary
  Set targetYearDict = ConstructTargets(nowWareki)

  ' yearNumに対応する年号を取得し，西暦に変換して返す
  On Error Resume Next
  Dim targetYear As YearObj
  Set targetYear = targetYearDict.Item(yearNum)
  If Err.Number > 0 Then
    GetSeireki = -1
    Exit Function
  End If
  
  ' 取得した和暦を西暦に変換して返す
  GetSeireki = ConvertWareki2Seireki(targetYear)
End Function

' 数字に対応する年号の辞書を作成
Private Function ConstructTargets(nowWareki As YearObj) As Dictionary
  Dim dict As Dictionary
  Set dict = New Dictionary
  
  ' 現在の年号から過去TargetYearWindow年分を追加
  Dim i As Integer
  Dim count As Integer
  count = 0
  For i = nowWareki.year To 1 Step -1
    If count >= TargetYearWindow Then Exit For
    Dim y As YearObj
    Set y = New YearObj
    y.SetYear nowWareki.Wareki, i
    dict.Add i, y
    count = count + 1
  Next i

  ' 必要であれば前の年号（令和->平成, 平成->昭和）の年を追加
  If count < TargetYearWindow Then
    Dim prevWareki As String
    Dim prevMaxYear As Integer
    Select Case nowWareki.Wareki
      Case "令和"
        prevWareki = "平成"
        prevMaxYear = Year(ReiwaStartDate) - Year(HeiseiStartDate) + 1
      Case "平成"
        prevWareki = "昭和"
        prevMaxYear = Year(HeiseiStartDate) - Year(ShowaStartDate) + 1
      Case Else
        prevWareki = ""
        prevMaxYear = 0
    End Select

    If prevWareki <> "" Then
      Dim prevYear As Integer
      prevYear = prevMaxYear
      Do While count < TargetYearWindow And prevYear >= 1
        Set y = New YearObj
        y.SetYear prevWareki, prevYear
        ' key はその年号内の年番号 (例: 平成31 -> key=31)
        dict.Add prevYear, y
        prevYear = prevYear - 1
        count = count + 1
      Loop
    End If
  End If

  Set ConstructTargets = dict
End Function

' 現在の年号を取得する関数
Private Function GetNowWareki() As YearObj
  Dim nowObj As YearObj
  Set nowObj = New YearObj

  Dim warekiName As String
  Dim warekiYear As Integer

  ' 各年号の開始日を定数で参照して判定
  ' 各年号の開始日をDate定数で判定
  If Date >= ReiwaStartDate Then
    warekiName = "令和"
    warekiYear = Year(Date) - Year(ReiwaStartDate) + 1
  ElseIf Date >= HeiseiStartDate Then
    warekiName = "平成"
    warekiYear = Year(Date) - Year(HeiseiStartDate) + 1
  ElseIf Date >= ShowaStartDate Then
    warekiName = "昭和"
    warekiYear = Year(Date) - Year(ShowaStartDate) + 1
  Else
    warekiName = "不明"
    warekiYear = 0
  End If

  nowObj.SetYear warekiName, warekiYear
  Set GetNowWareki = nowObj
End Function

' 与えられた数字が西暦か否かを返す
Private Function IsSeireki(num As Integer) As Boolean
  ' 数字が1000以上であれば西暦と判断
  IsSeireki = num > 1000
End Function

' 年号を西暦に変換する
' 年号データセットをファイル冒頭で定数宣言し，その値を用いて変換処理を実行する
Private  Function ConvertWareki2Seireki(wareki As YearObj) As Integer
  Dim startYear As Integer
  Select Case wareki.Wareki
    Case "令和"
      startYear = Year(ReiwaStartDate)
    Case "平成"
      startYear = Year(HeiseiStartDate)
    Case "昭和"
      startYear = Year(ShowaStartDate)
    Case Else
      ConvertWareki2Seireki = -1
      Exit Function
  End Select
  ConvertWareki2Seireki = startYear + wareki.year - 1
End Function

' 西暦を和暦に変換する
Private  Function ConvertSeireki2Wareki(seireki As Integer) As YearObj
  Dim result As YearObj
  Set result = New YearObj

  If seireki >= Year(ReiwaStartDate) Then
    result.SetYear "令和", seireki - Year(ReiwaStartDate) + 1
    Set ConvertSeireki2Wareki = result
    Exit Function
  ElseIf seireki >= Year(HeiseiStartDate) Then
    result.SetYear "平成", seireki - Year(HeiseiStartDate) + 1
    Set ConvertSeireki2Wareki = result
    Exit Function
  ElseIf seireki >= Year(ShowaStartDate) Then
    result.SetYear "昭和", seireki - Year(ShowaStartDate) + 1
    Set ConvertSeireki2Wareki = result
    Exit Function
  Else
    ' 対応しない年は Nothing を返す
    Set ConvertSeireki2Wareki = Nothing
  End If
End Function