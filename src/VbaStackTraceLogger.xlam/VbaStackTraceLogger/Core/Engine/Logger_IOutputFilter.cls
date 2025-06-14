VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_IOutputFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for output filtering"
'@Folder("VbaStackTraceLogger.Core.Engine")
'@Interface
'@ModuleDescription("Interface for output filtering")
Option Explicit

'@Description("Determine if log should be written based on tag and configuration")
Public Function ShouldWriteLog(ByVal selectedTag As LoggerLogTag, ByVal config As Logger_ConfigurationDTO) As Boolean
End Function
