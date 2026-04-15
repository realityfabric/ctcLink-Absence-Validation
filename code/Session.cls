VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Session"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("HCM_AB_VALIDATION")
Option Explicit

Private Type TSession
    FormClosedWithoutRunning As Boolean
    fpathABValidation As String
    fpathJobData As String
End Type
Private this As TSession

Private Sub Class_Initialize()
    this.FormClosedWithoutRunning = True
End Sub

Public Property Let FormClosedWithoutRunning(ByVal Value As Boolean)
    this.FormClosedWithoutRunning = Value
End Property

Public Property Get FormClosedWithoutRunning() As Boolean
    FormClosedWithoutRunning = this.FormClosedWithoutRunning
End Property

Public Property Get fpathABValidation() As String
    fpathABValidation = this.fpathABValidation
End Property

Public Property Let fpathABValidation(ByVal Value As String)
    this.fpathABValidation = Value
End Property

Public Property Get fpathJobData() As String
    fpathJobData = this.fpathJobData
End Property

Public Property Let fpathJobData(ByVal Value As String)
    this.fpathJobData = Value
End Property
