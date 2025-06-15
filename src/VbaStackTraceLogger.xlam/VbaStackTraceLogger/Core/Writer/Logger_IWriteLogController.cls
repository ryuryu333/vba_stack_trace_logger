VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_IWriteLogController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for Write Log Controller"
'@Folder("VbaStackTraceLogger.Core.Writer")
'@Interface
'@ModuleDescription("Interface for Write Log Controller")
Option Explicit

'@Description("Initialize with collection of writers")
Public Sub Initialize(ByVal writers As Collection)
End Sub

'@Description("Write log message to all writers")
Public Sub WriteLogToAll(ByVal writeMessage As String, ByVal logTag As LoggerLogTag, ByVal callPath As String)
End Sub

'@Description("Terminate all writers")
Public Sub Terminate()
End Sub
