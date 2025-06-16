VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_IStackTraceEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for Stack Trace Event Handler"
'@Folder("VbaStackTraceLogger.StackTrace")
'@Interface
'@ModuleDescription("Interface for Stack Trace Event Handler")
Option Explicit

'@Description("Handle procedure enter event")
Public Sub OnProcedureEnter(ByVal callSite As String)
End Sub

'@Description("Handle procedure exit event")
Public Sub OnProcedureExit()
End Sub
