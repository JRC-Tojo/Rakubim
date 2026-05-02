' --- 補助関数: クイックソート ---
Public Function QuickSort(v() As Double, left As Long, right As Long)
  Dim i As Long, j As Long, pivot As Double, tmp As Double
  i = left: j = right: pivot = v((left + right) \ 2)
  Do
    Do While v(i) < pivot: i = i + 1: Loop
    Do While pivot < v(j): j = j - 1: Loop
    If i <= j Then
      tmp = v(i): v(i) = v(j): v(j) = tmp
      i = i + 1: j = j - 1
    End If
  Loop While i <= j
  If left < j Then QuickSort v, left, j
  If i < right Then QuickSort v, i, right
End Function