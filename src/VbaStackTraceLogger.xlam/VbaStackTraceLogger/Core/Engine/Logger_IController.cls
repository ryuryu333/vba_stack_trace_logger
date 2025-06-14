VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_IController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for Logger Controller"
'@Folder("VbaStackTraceLogger.Core.Engine")
'@Interface
'@ModuleDescription("Interface for Logger Controller")
Option Explicit

'@Description("Write log message with specified tag")
Public Sub WriteLog(ByVal writeMessage As String, ByVal logTag As LoggerLogTag)
End Sub

'@Description("Create procedure tracer for stack trace logging")
Public Function UsingTracer(ByVal ModuleName As String, ByVal ProcedureName As String) As Logger_ProcedureTracer
End Function

'@Description("Terminate logger resources")
Public Sub Terminate()
End Sub
