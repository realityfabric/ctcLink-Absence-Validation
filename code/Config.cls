VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Config")
Option Explicit

Private Const ConfigDir As String = "config"
Private Const fClassifiedNonRepVac As String = "leave-accrual_vac_classified-nonrepresented.csv"
Private Const fClassifiedRepVac As String = "leave-accrual_vac_classified-represented.csv"

Private Type TConfig
    CLANonRepConfig As Collection ' Collection of ConfigRow objects
    CLARepConfig As Collection ' Collection of ConfigRow objects
End Type
Private this As TConfig

Private Sub Class_Initialize()
    ' TODO write private methods to load the config files into this
    ' TODO write a private method to sort the config
    ' TODO throw an error if the years are not consecutive starting at either 0 or 1
End Sub

' TODO write getters, include logic for years > max index
'   Use the last member of the collections to determine the
'   max years, after which that ConfigRow will be re-used.
' TODO decide if there should be private letters
