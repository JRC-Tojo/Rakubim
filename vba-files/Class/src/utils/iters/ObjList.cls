VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ObjList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Private inner_Collection As Collection

Private Sub class_initialize()
  Set inner_Collection = New Collection
End Sub

' アイテムの初期宣言
Public Function InitItems(ParamArray items() As Variant)
  Dim i As Long
  For i = LBound(items) To UBound(items)
    Me.Append items(i)
  Next i
  Set InitItems = Me
End Function

' アイテムの初期宣言（From 配列）
Public Function InitItemsFromArray(ByVal items As Variant)
  Dim i As Long
  For i = LBound(items) To UBound(items)
    Me.Append items(i)
  Next i
  Set InitItemsFromArray = Me
End Function

' アイテムの初期宣言（From Collection）
Public Function InitItemsFromCollection(items As Collection)
  Dim i As Long
  For i = 1 To items.Count
    Me.Append items(i)
  Next i
  Set InitItemsFromCollection = Me
End Function

' For Eachで呼び出されるメソッド
Public Function NewEnum() As IUnknown
  Attribute NewEnum.VB_UserMemId = -4
  Set NewEnum = inner_Collection.[_NewEnum]
End Function

' アイテム数の取得
Public Property Get Count() As Long
  Count = inner_Collection.Count
End Property

' アイテムの追加（追加する値はVariant）
Public Function Append(item As Variant)
  inner_Collection.Add item
End Function

' 外部から指定されたインデックスを、内部のCollectionに対応するインデックスに変換する
' 外部からは負の数の時に逆順で取得可能だが，内部用に正のIndexに修正する
Private Function convertIndex(index As Long) As Long
  If (index >= 0) Then
    convertIndex = index + 1
  Else
    convertIndex = Me.Count + index + 1
  End If
End Function

' アイテムの取得
' 負の数の時には逆順で取得可能
Public Function At(index As Long, Optional defValue As Variant = Nothing) As Variant
  Dim targetIndex As Long
  targetIndex = convertIndex(index)
  Dim isValidIdx As Boolean
  isValidIdx = targetIndex <= Me.Count

  If (isValidIdx) Then
    ' valueがObjectかわからないときには変数に代入できないため，このような実装としている
    If IsObject(inner_Collection(targetIndex)) Then
      Set At = inner_Collection(targetIndex)
    Else
      At = inner_Collection(targetIndex)
    End If
  Else
    If IsObject(defValue) Then
      Set At = defValue
    Else
      At = defValue
    End If
  End If
End Function

' アイテムの削除
Public Function Pop(index As Long)
  inner_Collection.Remove convertIndex(index)
End Function

' アイテムの全削除
Public Function Clear()
  Set inner_Collection = New Collection
End Function

' 値の更新
Public Function Update(index As Long, item As Variant)
  Dim targetIndex As Long
  targetIndex = convertIndex(index)
  inner_Collection.Remove targetIndex
  If (targetIndex = 1) Then
    inner_Collection.Add item, before:=1
  Else
    inner_Collection.Add item, after:=targetIndex - 1
  End If
End Function

' 値が含まれているか否かを返す
Public Function Contains(item As Variant) As Boolean
  Dim i As Long
  Contains = False
  For i = 1 To inner_Collection.Count
    If inner_Collection(i) = item Then
      Contains = True
      Exit Function
    End If
  Next i
End Function

' 別のObjListのアイテムをすべて追加する
Public Function Combine(otherList As ObjList) As ObjList
  Dim i As Long
  For i = 0 To otherList.Count - 1
    Call Me.Append(otherList.At(i))
  Next i
  Set Combine = Me
End Function

' 逆順のObjListを返す
Public Function Reverse(Optional updateMe As Boolean = False) As ObjList
  Dim reversedList As ObjList: Set reversedList = New ObjList
  Dim i As Long
  For i = Me.Count - 1 To 0 Step -1
    Call reversedList.Append(Me.At(i))
  Next i

  If (updateMe) Then
    Set inner_Collection = reversedList.ToCollection
  End If

  Set Reverse = reversedList
End Function

' Collection型に変換
Public Function ToCollection() As Collection
  Set ToCollection = inner_Collection
End Function

' 配列に変換
Public Function ToArray() As Variant
  Dim arr() As Variant
  ReDim arr(0 To inner_Collection.Count - 1) As Variant
  
  Dim i As Long
  For i = 1 To inner_Collection.Count
    arr(i - 1) = inner_Collection(i)
  Next i
  
  ToArray = arr
End Function

' 配列を指定した`chrs`で結合した文字列を返す
Public Function Join(chrs As String) As String
  Join = ""
  If (Me.Count = 0) Then Exit Function
  
  Dim joined As String
  Dim i As Long
  For i = 0 To Me.Count - 1
    joined = joined & CStr(Me.At(i)) & chrs
  Next i

  Join = Left(joined, Len(joined) - Len(chrs))
End Function

' 表示用の文字列を返す
Public Function ToString() As String
  ToString = "[" & Me.Join(", ") & "]"
End Function