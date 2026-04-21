VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConfigRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Config")
Option Explicit

Private Type TConfigRow
    IsInitialized As Boolean
    YearsWorked As Long
    AnnualAccrual As Long
    RequiresContinuousEmployment As Boolean
End Type
Private this As TConfigRow

'@Description("Initialize Class")
Private Sub Class_Initialize()
Attribute Class_Initialize.VB_Description = "Initialize Class"
    this.IsInitialized = False
End Sub

'@Description("Initialize read-only values for the instance.")
Public Sub Initialize( _
        ByVal Years As Long, _
        ByVal Hours As Long, _
        ByVal RequiresContinuous As Boolean)
Attribute Initialize.VB_Description = "Initialize read-only values for the instance."
    ' Values can only be initialized once.
    ' Attempting to initialize already set values should throw an error.
    If this.IsInitialized Then
        ' TODO: Throw error
    Else
        YearsWorked = Years
        AnnualAccrual = Hours
        RequiresContinuousEmployment = RequiresContinuous
        this.IsInitialized = True
    End If
End Sub

'@Description("Update the Years Worked to a number of years equal to or greater than 0.")
Private Property Let YearsWorked(ByVal Years As Long)
Attribute YearsWorked.VB_Description = "Update the Years Worked to a number of years equal to or greater than 0."
    If Years >= 0 Then
        this.YearsWorked = Years
    Else
        ' TODO throw error
    End If
End Property

'@Description("Get the number of years worked.")
Public Property Get YearsWorked() As Long
Attribute YearsWorked.VB_Description = "Get the number of years worked."
    YearsWorked = this.YearsWorked
End Property

'@Description("Update the annual accrual to a number of hours equal to or greater than 0.")
Private Property Let AnnualAccrual(ByVal Hours As Long)
Attribute AnnualAccrual.VB_Description = "Update the annual accrual to a number of hours equal to or greater than 0."
    If Hours >= 0 Then
        this.AnnualAccrual = Hours
    Else
        ' TODO throw error
    End If
End Property

'@Description("Get the number of hours accrued for annually.")
Public Property Get AnnualAccrual() As Long
Attribute AnnualAccrual.VB_Description = "Get the number of hours accrued for annually."
    AnnualAccrual = this.AnnualAccrual
End Property

'@Description("Update whether or not this instance requires continuous employment.")
Private Property Let RequiresContinuousEmployment(ByVal Requirement As Boolean)
Attribute RequiresContinuousEmployment.VB_Description = "Update whether or not this instance requires continuous employment."
    this.RequiresContinuousEmployment = Requirement
End Property

'@Description("Get whether or not this instance requires continuous employment.")
Public Property Get RequiresContinuousEmployment() As Boolean
Attribute RequiresContinuousEmployment.VB_Description = "Get whether or not this instance requires continuous employment."
    RequiresContinuousEmployment = this.RequiresContinuousEmployment
End Property
