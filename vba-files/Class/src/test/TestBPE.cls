VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestBPE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Public Function Run()
  Dim Ts As TestSuite
  Set Ts = New TestSuite
  Ts.Description = "TestBPE"

  Dim bpe As BytePairEncoding
  Set bpe = New BytePairEncoding

  Dim lines As Variant
  Dim vocab As Dictionary
  Dim tokenToFreqs As Dictionary
  
  ' 簡易な共通部分の抽出
  With Ts.Test("SimpleExtract")
    lines = Array("P3-P4ラーメン構造一般図", "P3-P4ラーメン計画基準図", "P3-P4ラーメン 設計要旨・注意事項")
    Set vocab = bpe.TrainVocab(lines)
    Set tokenToFreqs = bpe.ExtractSortedTokens(vocab)
    ' Collectionのインデックスは「1」はじまり
    .IsEqual tokenToFreqs.Keys(0), "P3-P4ラーメン"
    .IsEqual tokenToFreqs.Items(0), 3
  End With

  Set Run = Ts
End Function