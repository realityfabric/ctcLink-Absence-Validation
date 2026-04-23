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
    Dim Row_A As ConfigRow
    Dim Row_B As ConfigRow
    Set Row_A = New ConfigRow
    Set Row_B = New ConfigRow

    'Act:
    Row_A.Initialize 1, 1, True
    Row_B.Initialize 1, 1, False

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

'@TestMethod("Init")
Private Sub Initialize_NegativeYearsWorked()
    Const ExpectedError As Long = ErrorCode.INVALID_PROPERTY_VALUE
    On Error GoTo TestFail

    'Arrange:
    Dim TestRow As ConfigRow
    Set TestRow = New ConfigRow

    'Act:
    TestRow.Initialize -1, 2, True

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Init")
Private Sub Initialize_NegativeAnnualAccrual()
    Const ExpectedError As Long = ErrorCode.INVALID_PROPERTY_VALUE
    On Error GoTo TestFail

    'Arrange:
    Dim TestRow As ConfigRow
    Set TestRow = New ConfigRow

    'Act:
    TestRow.Initialize 1, -2, True

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Init")
Private Sub Initialize_Reinit_SameValues()
    Const ExpectedError As Long = ErrorCode.LET_READ_ONLY_PROPERTY
    On Error GoTo TestFail

    'Arrange:
    Dim TestRow As ConfigRow
    Set TestRow = New ConfigRow
    TestRow.Initialize 1, 2, True

    'Act:
    TestRow.Initialize 1, 2, True

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@Description("Tests MonthlyAccrual in the case that the result is a natural number.")
'@TestMethod("Function")
Private Sub MonthlyAccrual_N()
Attribute MonthlyAccrual_N.VB_Description = "Tests MonthlyAccrual in the case that the result is a natural number."
    On Error GoTo TestFail

    'Arrange:
    Dim ConfRow As ConfigRow
    Set ConfRow = New ConfigRow
    ConfRow.Initialize 1, 120, True

    'Act:
    'Assert:
    Assert.IsTrue ConfRow.MonthlyAccrual = 10

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@Description("Tests MonthlyAccrual in the case that the result is an irrational number.")
'@TestMethod("Function")
Private Sub MonthlyAccrual_I()
Attribute MonthlyAccrual_I.VB_Description = "Tests MonthlyAccrual in the case that the result is an irrational number."
    On Error GoTo TestFail

    'Arrange:
    Dim ConfRow As ConfigRow
    Set ConfRow = New ConfigRow
    ConfRow.Initialize 1, 136, True

    'Act:
    'Assert:
    Debug.Print ConfRow.MonthlyAccrual
    Debug.Print 136 / 12
    Debug.Print Round(136 / 12, 3)
    Debug.Print Round(ConfRow.AnnualAccrual / 12, 3)
    Assert.IsTrue ConfRow.MonthlyAccrual = 11.333

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
