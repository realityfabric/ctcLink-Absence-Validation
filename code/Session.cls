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
    fpathABValidation As String
    fpathJobData As String
End Type
Private this As TSession

Public Property Get fpathABValidation() As String
    fpathABValidation = this.fpathABValidation
End Property

Public Property Let fpathABValidation(ByVal value As String)
    this.fpathABValidation = value
End Property

Public Property Get fpathJobData() As String
    fpathJobData = this.fpathJobData
End Property

Public Property Let fpathJobData(ByVal value As String)
    this.fpathJobData = value
End Property
