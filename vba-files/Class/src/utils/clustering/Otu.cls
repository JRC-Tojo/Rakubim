VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Otu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

' --- メイン関数: 分割境界となるインデックスを返す ---
' 引数: sortedVals (Variant: 数値の1次元配列)
' 戻り値: 小さいグループの末尾インデックス（配列インデックスで返す）
Public Function ArgSplit(sortedVals As Variant) As Long
  Dim lb As Long, ub As Long, n As Long
  lb = LBound(sortedVals)
  ub = UBound(sortedVals)
  n = ub - lb + 1
  If n <= 2 Then ArgSplit = lb: Exit Function

  Dim maxVariance As Double: maxVariance = -1#
  Dim bestIdx As Long: bestIdx = lb

  Dim i As Long
  Dim sumTotal As Double: sumTotal = 0#
  For i = lb To ub
    sumTotal = sumTotal + CDbl(sortedVals(i))
  Next i

  ' 各隙間を閾値候補としてループ
  Dim sum1 As Double: sum1 = 0#
  Dim sum2 As Double: sum2 = 0#
  Dim n1 As Long: n1 = 0
  Dim n2 As Long: n2 = 0
  Dim w1 As Double: Dim w2 As Double
  Dim mu1 As Double: Dim mu2 As Double
  Dim variance As Double
  For i = lb To ub - 1
    n1 = (i - lb) + 1
    n2 = n - n1
    w1 = n1 / n
    w2 = n2 / n

    sum1 = sum1 + CDbl(sortedVals(i))
    sum2 = sumTotal - sum1
    mu1 = sum1 / n1
    mu2 = sum2 / n2

    variance = w1 * w2 * ((mu1 - mu2) ^ 2)

    If variance > maxVariance Then
      maxVariance = variance
      bestIdx = i
    End If
  Next i

  ArgSplit = bestIdx
End Function

' --- 閾値（値そのもの）を返すラッパー関数 ---
Public Function SplitThreshold(sortedVals As Variant) As Double
  Dim idx As Long
  idx = ArgSplit(sortedVals)

  ' 境界とその次の値の中間を閾値とする
  SplitThreshold = (CDbl(sortedVals(idx)) + CDbl(sortedVals(idx + 1))) / 2#
End Function