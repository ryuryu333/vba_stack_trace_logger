VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_IConfigReceiver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for receiving configuration"
'@Folder("VbaStackTraceLogger.Core.Engine")
'@Interface
'@ModuleDescription("Interface for receiving configuration")
Option Explicit

'@Description("Initialize with configuration data")
Public Sub Initialize(ByVal config As Logger_ConfigurationDTO)
End Sub
