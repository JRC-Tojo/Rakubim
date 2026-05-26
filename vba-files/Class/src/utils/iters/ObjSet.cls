VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ObjSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Private inner_Dictionary As Dictionary
Private inner_NewEnumCollection As Collection

Private Sub class_initialize()
  Set inner_Dictionary = New Dictionary
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
Public Function InitItemsFromArray(ByVal var As Variant)
  Dim i As Long
  For i = LBound(var) To UBound(var)
    Me.Append var(i)
  Next i
  Set InitItemsFromArray = Me
End Function

' アイテムの初期宣言（From Collection）
Public Function InitItemsFromCollection(coll As Collection)
  Dim i As Long
  For i = 1 To coll.Count
    Me.Append coll(i)
  Next i
  Set InitItemsFromCollection = Me
End Function

' For Eachで呼び出されるメソッド
Public Function NewEnum() As IUnknown
  Attribute NewEnum.VB_UserMemId = -4
  Set inner_NewEnumCollection = Me.ToCollection
  Set NewEnum = inner_NewEnumCollection.[_NewEnum]
End Function

' アイテム数の取得
Public Property Get Count() As Long
  Count = inner_Dictionary.Count
End Property

' アイテムの追加（追加する値はVariant）
Public Function Append(newItem As Variant) As Boolean
  Append = False
  If inner_Dictionary.Exists(newItem) Then Exit Function
  
  inner_Dictionary.Add newItem, newItem
  Append = True
End Function

' アイテムの削除
Public Function Pop(value As Variant) As Boolean
  Pop = False
  If Not inner_Dictionary.Exists(value) Then Exit Function
  
  inner_Dictionary.Remove value
  Pop = True
End Function

' アイテムの全削除
Public Function Clear()
  Set inner_Dictionary = New Dictionary
End Function

' 値が含まれているか否かを返す
Public Function Contains(value As Variant) As Boolean
  Contains = inner_Dictionary.Exists(value)
End Function

' 別の集合との和集合を返す
Public Function Union(otherSet As ObjSet) As ObjSet
  Dim unionSet As ObjSet
  Set unionSet = New ObjSet
  
  Dim myItem As Variant
  For Each myItem In Me.ToCollection()
    Call unionSet.Append(myItem)
  Next myItem

  Dim otherItem As Variant
  For Each otherItem In otherSet.ToCollection()
    Call unionSet.Append(otherItem)
  Next otherItem
  
  Set Union = unionSet
End Function

' 積集合を返す
' 引数: otherSet (ObjSet) ? 比較対象の集合
' 戻り値: 2つの集合に共通して存在する要素からなる新しい ObjSet
Public Function Intersection(otherSet As ObjSet) As ObjSet
  Dim interSet As ObjSet
  Set interSet = New ObjSet

  Dim item As Variant
  For Each item In Me.ToCollection()
    If otherSet.Contains(item) Then Call interSet.Append(item)
  Next item

  Set Intersection = interSet
End Function


' 差集合を返す（この集合 - otherSet）
' 引数: otherSet (ObjSet) ? 引かれる側の集合
' 戻り値: this に存在し otherSet に存在しない要素からなる新しい ObjSet
Public Function Difference(otherSet As ObjSet) As ObjSet
  Dim diffSet As ObjSet
  Set diffSet = New ObjSet

  Dim item As Variant
  For Each item In Me.ToCollection()
    If Not otherSet.Contains(item) Then Call diffSet.Append(item)
  Next item

  Set Difference = diffSet
End Function


' 部分集合かどうかを判定する
' 引数: otherSet (ObjSet) ? 上位集合と比較する集合
' 戻り値: Boolean ? Me が otherSet の部分集合なら True
Public Function IsSubsetOf(otherSet As ObjSet) As Boolean
  Dim item As Variant
  For Each item In Me.ToCollection()
    If Not otherSet.Contains(item) Then
      IsSubsetOf = False
      Exit Function
    End If
  Next item
  IsSubsetOf = True
End Function


' 上位集合かどうかを判定する
' 引数: otherSet (ObjSet) ? 部分集合と比較する集合
' 戻り値: Boolean ? Me が otherSet の上位集合（super set）なら True
Public Function IsSupersetOf(otherSet As ObjSet) As Boolean
  Dim item As Variant
  For Each item In otherSet.ToCollection()
    If Not Me.Contains(item) Then
      IsSupersetOf = False
      Exit Function
    End If
  Next item
  IsSupersetOf = True
End Function


' 完全に一致する集合かどうかを判定する
' 引数: otherSet (ObjSet) ? 比較する集合
' 戻り値: Boolean ? Me と otherSet が完全に同じ集合なら True
Public Function IsEqualTo(otherSet As ObjSet) As Boolean
  IsEqualTo = False
  If Me.Count <> otherSet.Count Then Exit Function
    
  ' １ループで判定するためにForで検証する
  Dim item As Variant
  For Each item In Me.ToCollection()
    If Not otherSet.Contains(item) Then Exit Function
  Next item

  IsEqualTo = True
End Function


' 互いに素（共通要素がない）かどうかを判定する
' 引数: otherSet (ObjSet)
' 戻り値: Boolean ? 共通要素が無ければ True
Public Function IsDisjointWith(otherSet As ObjSet) As Boolean
  Dim item As Variant
  For Each item In Me.ToCollection()
    If otherSet.Contains(item) Then
      IsDisjointWith = False
      Exit Function
    End If
  Next item
  IsDisjointWith = True
End Function

' Collection型に変換
Public Function ToCollection() As Collection
  Dim coll As New Collection
  Dim i As Long
  For i = 0 To inner_Dictionary.Count - 1
    Call coll.Add(inner_Dictionary.Items(i))
  Next i
  Set ToCollection = coll
End Function

' 配列に変換
Public Function ToArray() As Variant
  ToArray = inner_Dictionary.Items
End Function

' ObjList型に変換
Public Function ToList() As ObjList
  Dim list As New ObjList
  Set ToList = list.InitItemsFromArray(Me.ToArray())
End Function

' 配列を指定した`chrs`で結合した文字列を返す
Public Function Join(chrs As String) As String
  Join = ""
  If (Me.Count = 0) Then Exit Function
  
  Dim joined As String
  Dim value As Variant
  For Each value In Me.ToCollection()
    joined = joined & CStr(value) & chrs
  Next value

  Join = Left(joined, Len(joined) - Len(chrs))
End Function

' 表示用の文字列を返す
Public Function ToString() As String
  ToString = "(" & Me.Join(", ") & ")"
End Function