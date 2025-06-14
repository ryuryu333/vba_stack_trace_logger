VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_ITraceLogOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for trace log output"
'@Folder("VbaStackTraceLogger.Core.Engine")
'@Interface
'@ModuleDescription("Interface for trace log output")
Option Explicit

'@Description("Write trace log message with specified tag")
Public Sub WriteTraceLog(ByVal writeMessage As String, ByVal logTag As LoggerLogTag)
End Sub
