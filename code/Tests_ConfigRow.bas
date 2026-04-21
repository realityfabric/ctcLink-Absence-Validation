Attribute VB_Name = "Tests_ConfigRow"
'@TestModule
'@Folder("Tests")

Option Explicit
Option Private Module

Private Assert As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
End Sub

'@TestMethod("Uncategorized")
Private Sub Get_IsInitialized_Uninitialized()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestRow As ConfigRow
    Set TestRow = New ConfigRow
    
    'Assert:
    Assert.IsFalse TestRow.IsInitialized

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Get_IsInitialized_Initialized()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestRow As ConfigRow
    Set TestRow = New ConfigRow
    
    'Act:
    TestRow.Initialize 1, 1, True
    
    'Assert:
    Assert.IsTrue TestRow.IsInitialized

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
