VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_IConfigBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for Configuration Builder"
'@Folder("VbaStackTraceLogger.Core.Config")
'@Interface
'@ModuleDescription("Interface for Configuration Builder")
Option Explicit

'@Description("Initialize builder with configuration receiver")
Public Sub Initialize(ByVal configReceiver As Logger_IConfigReceiver)
End Sub
