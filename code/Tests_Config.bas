Attribute VB_Name = "Tests_Config"
'@TestModule
'@Folder("Tests")

Option Explicit
Option Private Module

Private Assert As Object
Private Conf As Config

Private Const CONFIG_TEST_DIR As String = "test_data"
Private Const CLA_NONREP_CONFIG_FILENAME As String = "leave-accrual_vac_classified-nonrepresented.csv"
Private Const CLA_REP_CONFIG_FILENAME As String = "leave-accrual_vac_classified-represented.csv"

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
'@TestMethod("IO")
Private Sub LoadConfig_Defaults_NoFail()
Attribute LoadConfig_Defaults_NoFail.VB_Description = "Succeeds if the configs load without an error."
    On Error GoTo TestFail

    Conf.LoadConfigs ' Use defaults

    Assert.Succeed

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@Description("Succeeds if the configs load without an error.")
'@TestMethod("IO")
Private Sub LoadConfig_TestConfig_NoFail()
Attribute LoadConfig_TestConfig_NoFail.VB_Description = "Succeeds if the configs load without an error."
    On Error GoTo TestFail

    Conf.LoadConfigs _
        ConfigDirectory:=CONFIG_TEST_DIR, _
        CLANonRepVacFileName:=CLA_NONREP_CONFIG_FILENAME, _
        CLARepVacFileName:=CLA_REP_CONFIG_FILENAME

    Assert.Succeed

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@Description("Test passes if all years are correct.")
'@TestMethod("IO")
Private Sub LoadConfig_TestConfig_YearsWorkedCorrect_NonRep()
Attribute LoadConfig_TestConfig_YearsWorkedCorrect_NonRep.VB_Description = "Test passes if all years are correct."
    On Error GoTo TestFail

    ' Arrange
    Conf.LoadConfigs _
        ConfigDirectory:=CONFIG_TEST_DIR, _
        CLANonRepVacFileName:=CLA_NONREP_CONFIG_FILENAME, _
        CLARepVacFileName:=CLA_REP_CONFIG_FILENAME

    ' Assert
    '@Ignore UseMeaningfulName
    Dim i As Long
    For i = 1 To 16
        Assert.IsTrue Conf.CLA_NonRep_Item(i).YearsWorked = i
    Next i

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@Description("Test passes if all years are correct.")
'@TestMethod("IO")
Private Sub LoadConfig_TestConfig_YearsWorkedCorrect_Rep()
Attribute LoadConfig_TestConfig_YearsWorkedCorrect_Rep.VB_Description = "Test passes if all years are correct."
    On Error GoTo TestFail

    ' Arrange
    Conf.LoadConfigs _
        ConfigDirectory:=CONFIG_TEST_DIR, _
        CLANonRepVacFileName:=CLA_NONREP_CONFIG_FILENAME, _
        CLARepVacFileName:=CLA_REP_CONFIG_FILENAME

    ' Assert
    '@Ignore UseMeaningfulName
    Dim i As Long
    For i = 1 To 16
        Assert.IsTrue Conf.CLA_Rep_Item(i).YearsWorked = i
    Next i

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
