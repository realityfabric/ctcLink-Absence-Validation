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

'@TestMethod("Properties")
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

'@TestMethod("Properties")
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

'@TestMethod("Init")
Private Sub Initialize_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestRowContinuous As ConfigRow
    Dim TestRowNonContinuous As ConfigRow
    Set TestRowContinuous = New ConfigRow
    Set TestRowNonContinuous = New ConfigRow
    
    'Act:
    TestRowContinuous.Initialize 1, 1, True
    TestRowContinuous.Initialize 1, 1, False
    
    'Assert:
    Assert.Succeed

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Properties")
Private Sub Get_YearsWorked()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestRow As ConfigRow
    Set TestRow = New ConfigRow
    TestRow.Initialize 1, 2, True
    
    'Assert:
    Assert.IsTrue 1 = TestRow.YearsWorked

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Properties")
Private Sub Get_AnnualAccrual()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestRow As ConfigRow
    Set TestRow = New ConfigRow
    TestRow.Initialize 1, 2, True
    
    'Assert:
    Assert.IsTrue 2 = TestRow.AnnualAccrual

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Properties")
Private Sub Get_RequiresContinuousEmployment_True()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestRow As ConfigRow
    Set TestRow = New ConfigRow
    TestRow.Initialize 1, 2, True
    
    'Assert:
    Assert.IsTrue TestRow.RequiresContinuousEmployment

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Properties")
Private Sub Get_RequiresContinuousEmployment_False()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestRow As ConfigRow
    Set TestRow = New ConfigRow
    TestRow.Initialize 1, 2, False
    
    'Assert:
    Assert.IsFalse TestRow.RequiresContinuousEmployment

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
