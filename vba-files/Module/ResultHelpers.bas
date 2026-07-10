Attribute VB_Name = "ResultHelpers"
Option Explicit

' Helper functions to create ObjResult/ObjResultErr objects
Public Function Success(Optional v As Variant = Nothing) As ObjResult
  Dim r As ObjResult
  Set r = New ObjResult
  Set Success = r.InitSuccess(v)
End Function

Public Function Failure(errKey As String, ParamArray args() As Variant) As ObjResult
  Dim re As ObjResultErr
  Set re = New ObjResultErr
  Set re = re.Init(errKey, args)

  ' FailureÇÃèÍçáÇ…ÇÕä»à’LogÇèoóÕÇ∑ÇÈ
  Dim argsList As ObjList
  Set argsList = New ObjList
  Call argsList.InitItemsFromArray(args)
  Debug.Print "[ERROR] (" & errKey & "): " & argsList.Join(", ")

  Dim r As ObjResult
  Set r = New ObjResult
  Set Failure = r.InitFailure(re)
End Function

Public Function IsResultAllOK(ParamArray results() As Variant) As Boolean
  Dim resultList As ObjList
  Set resultList = New ObjList
  Call resultList.InitItemsFromArray(results)

  Dim result As ObjResult
  For Each result In resultList
    If (Not result.isOK) Then
      IsResultAllOK = False
      Exit Function
    End If
  Next

  IsResultAllOK = True
End Function

Public Function PickResultErr(successVal As Variant, ParamArray results() As Variant) As ObjResult
  Dim resultList As ObjList
  Set resultList = New ObjList
  Call resultList.InitItemsFromArray(results)

  Dim result As ObjResult
  For Each result In resultList
    If (Not result.isOK) Then
      Set PickResultErr = result
      Exit Function
    End If
  Next

  Set PickResultErr = Success(successVal)
End Function