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

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As LongPtr) As Long

' スタイルの詳細は`https://learn.microsoft.com/ja-jp/windows/win32/winmsg/window-styles`を参照
Private Const GWL_STYLE As Long = -16
Private Const WS_CAPTION As Long = &HC00000     ' タイトルバー
Private Const WS_THICKFRAME As Long = &H40000   ' サイズ変更枠（太い枠）
Private Const WS_DLGFRAME As Long = &H400000    ' ダイアログ枠（二重枠）
Private Const WS_BORDER As Long = &H800000      ' 単線枠

Private Sub UserForm_Initialize()
  Dim hWnd As LongPtr
  Dim currentStyle As LongPtr
  
  ' 処理中メッセージの表示
  MessageArea.Caption = "初期化処理を実行中..."
  
  ' 1. フォームのウィンドウハンドルを取得
  hWnd = FindWindow("ThunderDFrame", Me.Caption)
  
  If hWnd <> 0 Then
    ' 2. 現在のウィンドウスタイルを取得
    currentStyle = GetWindowLongPtr(hWnd, GWL_STYLE)
    
    ' 3. タイトルバーと、あらゆる枠線スタイル（太枠・ダイアログ枠・単線枠）をすべて除去
    currentStyle = currentStyle And Not (WS_CAPTION Or WS_THICKFRAME Or WS_DLGFRAME Or WS_BORDER)
    
    ' 4. 新しいスタイルをウィンドウに適用
    Call SetWindowLongPtr(hWnd, GWL_STYLE, currentStyle)
    
    ' 5. ウィンドウの枠組み（メニューバーなど）を再描画して変更を反映
    Call DrawMenuBar(hWnd)
  End If
End Sub

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