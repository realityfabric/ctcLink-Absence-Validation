Attribute VB_Name = "Tests_Config"
'@TestModule
'@Folder("Tests")

Option Explicit
Option Private Module

Private Assert As Object
Private Conf As Config

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Set Conf = New Config
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set Conf = Nothing
End Sub

'@Description("Succeeds if the configs load without an error.")
'@TestMethod("Uncategorized")
Private Sub LoadConfig_Defaults_Success()
Attribute LoadConfig_Defaults_Success.VB_Description = "Succeeds if the configs load without an error."
    On Error GoTo TestFail
    
    Conf.LoadConfigs
    
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
Private Sub Let_ConfigDir_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Conf As Config
    Set Conf = New Config
    
    'Act:
    Conf.ConfigDir = "C:\Test_Path"
    
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
Private Sub Get_ConfigDir()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Conf As Config
    Set Conf = New Config
    
    'Act:
    Conf.ConfigDir = "C:\Test_Path"
    
    'Assert:
    Assert.IsTrue "C:\Test_Path" = Conf.ConfigDir

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

