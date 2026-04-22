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

'@Description("Test passes if all Accurals are correct.")
'@TestMethod("IO")
Private Sub LoadConfig_TestConfig_AnnualAccrualCorrect_Rep()
Attribute LoadConfig_TestConfig_AnnualAccrualCorrect_Rep.VB_Description = "Test passes if all Accurals are correct."
    On Error GoTo TestFail

    ' Arrange
    Conf.LoadConfigs _
        ConfigDirectory:=CONFIG_TEST_DIR, _
        CLANonRepVacFileName:=CLA_NONREP_CONFIG_FILENAME, _
        CLARepVacFileName:=CLA_REP_CONFIG_FILENAME

    ' Assert
    With Conf
        Assert.IsTrue .CLA_Rep_Item(1).AnnualAccrual = 112
        Assert.IsTrue .CLA_Rep_Item(2).AnnualAccrual = 112

        Assert.IsTrue .CLA_Rep_Item(3).AnnualAccrual = 120

        Assert.IsTrue .CLA_Rep_Item(4).AnnualAccrual = 128

        Assert.IsTrue .CLA_Rep_Item(5).AnnualAccrual = 136
        Assert.IsTrue .CLA_Rep_Item(6).AnnualAccrual = 136

        Assert.IsTrue .CLA_Rep_Item(7).AnnualAccrual = 144
        Assert.IsTrue .CLA_Rep_Item(8).AnnualAccrual = 144
        Assert.IsTrue .CLA_Rep_Item(9).AnnualAccrual = 144

        Assert.IsTrue .CLA_Rep_Item(10).AnnualAccrual = 160
        Assert.IsTrue .CLA_Rep_Item(11).AnnualAccrual = 160
        Assert.IsTrue .CLA_Rep_Item(12).AnnualAccrual = 160
        Assert.IsTrue .CLA_Rep_Item(13).AnnualAccrual = 160
        Assert.IsTrue .CLA_Rep_Item(14).AnnualAccrual = 160

        Assert.IsTrue .CLA_Rep_Item(15).AnnualAccrual = 176
        Assert.IsTrue .CLA_Rep_Item(16).AnnualAccrual = 176
        Assert.IsTrue .CLA_Rep_Item(17).AnnualAccrual = 176
        Assert.IsTrue .CLA_Rep_Item(18).AnnualAccrual = 176
        Assert.IsTrue .CLA_Rep_Item(19).AnnualAccrual = 176

        Assert.IsTrue .CLA_Rep_Item(20).AnnualAccrual = 192
        Assert.IsTrue .CLA_Rep_Item(21).AnnualAccrual = 192
        Assert.IsTrue .CLA_Rep_Item(22).AnnualAccrual = 192
        Assert.IsTrue .CLA_Rep_Item(23).AnnualAccrual = 192
        Assert.IsTrue .CLA_Rep_Item(24).AnnualAccrual = 192

        Assert.IsTrue .CLA_Rep_Item(25).AnnualAccrual = 200
    End With
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IO")
Private Sub LoadConfig_AnnualAccrualCorrect_GTMax_Rep()
    On Error GoTo TestFail

    ' Arrange
    Conf.LoadConfigs _
        ConfigDirectory:=CONFIG_TEST_DIR, _
        CLANonRepVacFileName:=CLA_NONREP_CONFIG_FILENAME, _
        CLARepVacFileName:=CLA_REP_CONFIG_FILENAME

    ' Assert
    With Conf
        Dim i As Long
        For i = 26 To 100
            Assert.IsTrue .CLA_Rep_Item(i).AnnualAccrual = 200
        Next i
    End With
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
