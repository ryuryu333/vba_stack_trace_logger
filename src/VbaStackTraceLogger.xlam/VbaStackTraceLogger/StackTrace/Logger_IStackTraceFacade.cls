VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_IStackTraceFacade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for Stack Trace Facade"
'@Folder("VbaStackTraceLogger.StackTrace")
'@Interface
'@ModuleDescription("Interface for Stack Trace Facade")
Option Explicit

'@Description("Initialize stack trace controller with settings")
Public Sub Initialize(ByVal IsStackTraceEnabled As Boolean, ByVal traceLogOutput As Logger_ITraceLogOutput)
End Sub

'@Description("Create procedure tracer for specified module and procedure")
Public Function UsingTracer(ByVal ModuleName As String, ByVal ProcedureName As String) As Logger_ProcedureTracer
End Function

'@Description("Get current call path as string")
Public Function GetCurrentCallPath() As String
End Function

'@Description("Terminate stack trace controller")
Public Sub Terminate()
End Sub
