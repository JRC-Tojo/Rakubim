VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ObjPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

' ObjPath: ファイルパスを扱うユーティリティオブジェクト
' プロパティ: FullPath, Name, Extension, Stem
Private p_FullPath As String

Public Function Init(path As String) As ObjPath
  p_FullPath = path
  Set Init = Me
End Function

Public Property Get FullPath() As String
  FullPath = p_FullPath
End Property

Public Property Get IsFile() As Boolean
  IsFile = (Len(Me.Extension) > 0)
End Property

Public Property Get Name() As String
  Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
  Name = fso.GetFileName(p_FullPath)
End Property

Public Property Get Extension() As String
  Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
  Extension = LCase$(fso.GetExtensionName(p_FullPath))
End Property

Public Property Get Stem() As String
  Dim nm As String: nm = Me.Name
  Dim pos As Long: pos = InStrRev(nm, ".")
  If pos > 0 Then
      Stem = Left$(nm, pos - 1)
  Else
      Stem = nm
  End If
End Property

Public Function Parent(Optional level As Long = 0) As ObjPath
  Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
  Dim parentPath As String
  parentPath = p_FullPath

  Dim i As Long
  For i = 0 To level
    parentPath = fso.GetParentFolderName(parentPath)
  Next i

  ' ルートまで到達した場合のガード
  If parentPath = "" Then
    parentPath = Left(p_FullPath, 3)  ' "C:\" 等を返す
  End If

  Dim parentObj As ObjPath
  Set parentObj = New ObjPath
  Call parentObj.Init(parentPath)
  Set Parent = parentObj
End Function

Public Function Child(ParamArray names() As Variant) As ObjPath
  Dim combined As String
  combined = p_FullPath
  Dim i As Long
  For i = LBound(names) To UBound(names)
    If Right$(combined, 1) <> "\" Then
      combined = combined & "\"
    End If
    combined = combined & CStr(names(i))
  Next i

  Dim childPath As ObjPath
  Set childPath = New ObjPath
  Call childPath.Init(combined)
  Set Child = childPath
End Function

Public Function ToString() As String
  ToString = p_FullPath
End Function

Public Function Exists() As Boolean
  Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
  If (Me.IsFile) Then
    Exists = fso.FileExists(p_FullPath)
  Else
    Exists = fso.FolderExists(p_FullPath)
  End If
End Function

Public Function Colldir() As ObjList
  Dim pathObj As ObjPath
  Dim childPaths As ObjList
  Set childPaths = New ObjList

  ' 自身が存在しない場合，またはフォルダでない場合は空配列を返す
  If (Not Me.Exists() Or Me.IsFile) Then
    Set Colldir = childPaths
    Exit Function
  End If

  Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
  Dim folder As Object
  Set folder = fso.GetFolder(p_FullPath)

  Dim file As Object
  For Each file In folder.Files
    Set pathObj = New ObjPath
    Call pathObj.Init(file.Path)
    Call childPaths.Append(pathObj)
  Next file

  Dim subfolder As Object
  For Each subfolder In folder.SubFolders
    Set pathObj = New ObjPath
    Call pathObj.Init(subfolder.Path)
    Call childPaths.Append(pathObj)
  Next subfolder

  Set Colldir = childPaths
End Function