VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ObjDict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Private inner_Dictionary As Dictionary
Private inner_KVcoll As Collection

Private Sub class_initialize()
  Set inner_Dictionary = New Dictionary
End Sub

' For Eachで呼び出されるメソッド
Public Function NewEnum() As IUnknown
  Attribute NewEnum.VB_UserMemId = -4
  ' 直接Me.KVcollを返すと，For Eachの呼び出し時にはKVcollで生成したCollectionが破棄されて参照できなくなる
  ' -> 「オブジェクトが必要です」というエラーになる
  ' -> クラス変数に退避させることでオブジェクトが破棄されないようにする
  Call buildKVcoll()
  Set NewEnum = inner_KVcoll.[_NewEnum]
End Function

' KeyValueオブジェクトのCollectionを返す
Private Function buildKVcoll() As Collection
  Dim coll As Collection
  Set coll = New Collection
  
  Dim key As Variant
  Dim kv As ModelKeyValue
  For Each key In inner_Dictionary.Keys
    Set kv = New ModelKeyValue
    coll.Add kv.InitModel(key, inner_Dictionary(key))
  Next key
  
  Set inner_KVcoll = coll
End Function

' Keyの一覧を取得
Public Property Get Keys() As ObjList
  Dim l As ObjList
  Set l = New ObjList
  Call l.InitItemsFromArray(inner_Dictionary.Keys)
  Set Keys = l
End Property

' Valueの一覧を取得
Public Property Get Values() As ObjList
  Dim l As ObjList
  Set l = New ObjList
  Call l.InitItemsFromArray(inner_Dictionary.Items)
  Set Values = l
End Property

' アイテム数の取得
Public Property Get Count() As Long
  Count = inner_Dictionary.Count
End Property

' アイテムの追加（KeyとValueはVariant）
Public Function Append(newKey As Variant, newValue As Variant) As Boolean
  Append = False
  If inner_Dictionary.Exists(newKey) Then Exit Function
  
  inner_Dictionary.Add newKey, newValue
  Append = True
End Function

' Keyに対応するValueの取得
Public Function At(key As Variant, Optional defValue As Variant = Nothing) As Variant
  If inner_Dictionary.Exists(key) Then
    If IsObject(inner_Dictionary(key)) Then
      Set At = inner_Dictionary(key)
    Else
      At = inner_Dictionary(key)
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
Public Function Remove(key As Variant) As Boolean
  Remove = False
  If Not inner_Dictionary.Exists(key) Then Exit Function
  
  inner_Dictionary.Remove key
  Remove = True
End Function

' 値の更新
Public Function Update(key As Variant, newValue As Variant) As Boolean
  Update = False
  If Not inner_Dictionary.Exists(key) Then Exit Function

  Call Me.Remove(key)
  Call Me.Append(key, newValue)
  Update = True
End Function

' Keyの存在確認
Public Function Exists(key As Variant) As Boolean
  Exists = inner_Dictionary.Exists(key)
End Function

