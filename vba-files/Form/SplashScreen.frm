VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SplashScreen 
   Caption         =   ""
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7755
   OleObjectBlob   =   "SplashScreen.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "SplashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Activate()
  Dim startTime As Double
  Dim minDuration As Double

  ' --- 設定：最低表示時間（秒） ---
  minDuration = 1.5

  ' 開始時間を記録
  startTime = Timer

  ' フォームの描画をシステムに強制的に反映させる
  Me.Repaint
  DoEvents

  ' 1. 裏で初期化処理を実行
  Dim r As Routes
  Set r = New Routes
  ' 目録入力シートにプルダウンを設置
  Call r.SetPulldownForFiles()
  ' コードシートのプルダウンを設置
  Call r.CodeSheetInitializer()

  ' 2. 最低表示時間が経過するまでループ待機
  ' （DoEventsを入れることで、待機中も画面がフリーズしません）
  Do While Timer - startTime < minDuration
    DoEvents
    
    ' 深夜0時をまたいだ場合の簡易リセット対策
    If Timer < startTime Then startTime = Timer
  Loop

  ' 3. 条件を満たしたらフォームを閉じる
  Unload Me
End Sub

' フォームがユーザーによって×ボタン等で閉じられるのを防ぐ（オプション）
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = vbFormControlMenu Then
    Cancel = True
  End If
End Sub