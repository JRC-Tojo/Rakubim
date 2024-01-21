Attribute VB_Name = "CollectionUtils"
' 連番の配列を生成する
Function RangeCollection(startNum As Integer, endNum As Integer)
  Dim returns As Collection
  Set returns = New Collection
  
  For i = 0 To endNum - startNum
    returns.Add (startNum + i)
  Next i
  
  Set RangeCollection = returns
End Function

' 配列が値を有しているかチェック
Function HasItem(objCol As Collection, checkItem As Variant)
  On Error Resume Next
     
  'Itemメソッドを実行
  Call objCol.item(checkItem)
         
  'エラー値がない場合：キー検索はヒット（戻り値：True）
  If Err.Number = 0 Then
    HasItem = True
  Else
    HasItem = False
  End If
End Function

' 最大値のインデックスを取得
Function ArgMax(col As Collection)
  Dim maxVal As Variant
  Dim maxValIdx As Integer
  maxVal = -9999
  maxValIdx = 0
  
  Dim i As Integer
  For i = 1 To col.Count
    If col(i) > maxVal Then
      maxVal = col(i)
      maxValIdx = i
    End If
  Next i
  
  ArgMax = Array(maxVal, maxValIdx)
End Function

' 配列のすべての値を表示する
Function printCollectionItems(col As Collection)
  Dim colItem As Variant
  Dim s As String
  s = ""
  
  For Each colItem In col
    s = s & colItem & ", "
  Next colItem
  
  Debug.Print s
End Function

' dictからキー一覧を取得する
Function dictKeys(dict As Scripting.Dictionary) As Collection
  Dim returnCol As Collection
  Set returnCol = New Collection
  
  For Each dkey In dict
    returnCol.Add dkey
  Next
  
  Set dictKeys = returnCol
End Function

' dictから値一覧を取得する
Function dictValues(dict As Scripting.Dictionary) As Collection
  Dim returnCol As Collection
  Set returnCol = New Collection
  
  For Each dkey In dict
    returnCol.Add dict(dkey)
  Next
  
  Set dictValues = returnCol
End Function

' ２つの配列を結合する
Function CombineCols(ByVal col1 As Collection, ByVal col2 As Collection) As Collection
  For Each item As Object In col2
    Call col1.Add(item)
  Next
  Set CombineCols = col1
End Function