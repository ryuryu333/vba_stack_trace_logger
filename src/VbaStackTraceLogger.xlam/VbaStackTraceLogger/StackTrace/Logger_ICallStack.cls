VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_ICallStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for Call Stack Management"
'@Folder("VbaStackTraceLogger.StackTrace")
'@Interface
'@ModuleDescription("Interface for Call Stack Management")
Option Explicit

'@Description("Enter a procedure and update call stack")
Public Sub EnterProcedure(ByVal callSite As String)
End Sub

'@Description("Exit current procedure and update call stack")
Public Sub ExitProcedure()
End Sub

'@Description("Get current call Site as string")
Public Function GetCurrentCallSite() As String
End Function

'@Description("Get current call path as string")
Public Function GetCurrentCallPath() As String
End Function

