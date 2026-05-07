Attribute VB_Name = "TestAllRunner"
Option Explicit

Public Sub RunAllTests()
  ' Report results to the Immediate Window
  ' (ctrl + g or View > Immediate Window)
  Dim Reporter As New ImmediateReporter
  Dim tests As Collection

  Set tests = New Collection
  tests.Add New TestObjDict
  tests.Add New TestObjList
  tests.Add New TestObjPath
  tests.Add New TestObjRegex
  tests.Add New TestObjSet
  tests.Add New TestRepoSheet
  tests.Add New TestServiceFile
  tests.Add New TestServiceMatching
  tests.Add New TestServiceNameAnalyzer
  tests.Add New TestServiceNameCleaner
  tests.Add New TestBPE
  tests.Add New TestCluster
  tests.Add New TestRealCases

  Dim t As Object
  For Each t In tests
    Reporter.ListenTo t.Run()
  Next t
End Sub


' ================= ŒÂ•ÊƒeƒXƒg =================

Public Sub GenerateRealFiles()
  Dim testRealCases As TestRealCases
  Set testRealCases = New TestRealCases
  Call testRealCases.GenerateRealFiles()
End Sub