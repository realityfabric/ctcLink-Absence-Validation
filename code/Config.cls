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

Private Type TConfig
    ConfigDir As String
    fClassifiedNonRepVac As String
    fClassifiedRepVac As String
    CLANonRepConfig As Collection ' Collection of ConfigRow objects
    CLARepConfig As Collection ' Collection of ConfigRow objects
End Type
Private this As TConfig

Private Sub Class_Initialize()
    Set this.CLANonRepConfig = New Collection
    Set this.CLARepConfig = New Collection
End Sub

'@Description("Load configurations, optionally change config filenames and directory.")
'@Ignore MultilineParameter
Public Sub LoadConfigs( _
        Optional ByVal ConfigDirectory As String _
            = "config", _
        Optional ByVal CLANonRepVacFileName As String _
            = "leave-accrual_vac_classified-nonrepresented.csv", _
        Optional ByVal CLARepVacFileName As String _
            = "leave-accrual_vac_classified-represented.csv")
Attribute LoadConfigs.VB_Description = "Load configurations, optionally change config filenames and directory."
    ' TODO write private methods to load the config files into this
    ' TODO write a private method to sort the config
    ' TODO throw an error if the years are not consecutive starting at either 0 or 1
End Sub

'@Description("Update the name of the directory where configuration files are kept.")
Public Property Let ConfigDir(ByVal DirString As String)
Attribute ConfigDir.VB_Description = "Update the name of the directory where configuration files are kept."
    this.ConfigDir = DirString
End Property

'@Description("Returns the name of the directory where configuration files are kept.")
Public Property Get ConfigDir() As String
Attribute ConfigDir.VB_Description = "Returns the name of the directory where configuration files are kept."
    ConfigDir = this.ConfigDir
End Property

'@Description("Update the file name of the configuration for Classified Non-Represented vacation accrual.")
Public Property Let fClassifiedNonRepVac(ByVal FileName As String)
Attribute fClassifiedNonRepVac.VB_Description = "Update the file name of the configuration for Classified Non-Represented vacation accrual."
    this.fClassifiedNonRepVac = FileName
End Property

'@Description("Returns the file name of the configuration for Classified Non-Represented vacation accrual.")
Public Property Get fClassifiedNonRepVac() As String
Attribute fClassifiedNonRepVac.VB_Description = "Returns the file name of the configuration for Classified Non-Represented vacation accrual."
    fClassifiedNonRepVac = this.fClassifiedNonRepVac
End Property

'@Description("Update the file name of the configuration for Classified Represented vacation accrual.")
Public Property Let fClassifiedRepVac(ByVal FileName As String)
Attribute fClassifiedRepVac.VB_Description = "Update the file name of the configuration for Classified Represented vacation accrual."
    this.fClassifiedRepVac = FileName
End Property

'@Description("Returns the file name of the configuration for Classified Represented vacation accrual.")
Public Property Get fClassifiedRepVac() As String
Attribute fClassifiedRepVac.VB_Description = "Returns the file name of the configuration for Classified Represented vacation accrual."
    fClassifiedRepVac = this.fClassifiedRepVac
End Property

' TODO write getters, include logic for years > max index
'   Use the last member of the collections to determine the
'   max years, after which that ConfigRow will be re-used.
' TODO decide if there should be private letters
