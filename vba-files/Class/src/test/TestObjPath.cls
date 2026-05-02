VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestObjPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Public Function Run()
  Dim Ts As TestSuite: Set Ts = New TestSuite
  Ts.Description = "TestObjPath"

  With Ts.Test("Methods")
    Dim p As ObjPath: Set p = New ObjPath
    Call p.Init("C:\TestFolder\SubFolder\TestFile.txt")
    .IsEqual p.FullPath, "C:\TestFolder\SubFolder\TestFile.txt"
    .IsEqual p.Extension, "txt"
    .IsEqual p.Name, "TestFile.txt"
    .IsEqual p.Stem, "TestFile"
    .IsEqual p.ToString(), "C:\TestFolder\SubFolder\TestFile.txt"
    .IsOk p.IsFile
    Dim parentPath As ObjPath: Set parentPath = p.Parent()
    .IsEqual parentPath.FullPath, "C:\TestFolder\SubFolder"
    Dim parent2Path As ObjPath: Set parent2Path = p.Parent(1)
    .IsEqual parent2Path.FullPath, "C:\TestFolder"
    Dim parent3Path As ObjPath: Set parent3Path = p.Parent(2)
    .IsEqual parent3Path.FullPath, "C:\"
    Dim parent4Path As ObjPath: Set parent4Path = p.Parent(3)
    .IsEqual parent4Path.FullPath, "C:\"
    Dim childPath As ObjPath: Set childPath = parent2Path.Child("NewFolder")
    .IsEqual childPath.FullPath, "C:\TestFolder\NewFolder"
    Dim childFilePath As ObjPath: Set childFilePath = parent2Path.Child("ChildFolder", "NewFile.docx")
    .IsEqual childFilePath.FullPath, "C:\TestFolder\ChildFolder\NewFile.docx"
  End With

  Call WithTempFolderHelper(Ts.test("WithTempFolder"))

  Set Run = Ts
End Function

Private Sub WithTempFolderHelper(t As TestCase)
  Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
  Dim tempFolderPath As String: tempFolderPath = ActiveWorkbook.Path & "\TestTempFolder"
  Dim subFolderPath As String: subFolderPath = tempFolderPath & "\SubFolder"
  
  On Error GoTo CleanupAndFail
  If Not fso.FolderExists(tempFolderPath) Then
    fso.CreateFolder tempFolderPath
    fso.CreateFolder subFolderPath
    fso.CreateTextFile tempFolderPath & "\file1.pdf", True
    fso.CreateTextFile tempFolderPath & "\file2.dwg", True
    fso.CreateTextFile subFolderPath & "\file3.pdf", True
  End If

  Dim p As ObjPath: Set p = New ObjPath
  Call p.Init(tempFolderPath)
  t.IsOk p.Exists()
  t.IsEqual p.Colldir().Count, 3

Cleanup:
  On Error Resume Next
  fso.DeleteFolder tempFolderPath, True
  Exit Sub

CleanupAndFail:
  Debug.Print "FAIL: TestObjPath.WithTempFolderHelper - " & Err.Description
  Resume Cleanup
End Sub
