VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_ILogWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Logger_ILogWriter.cls
'@Folder("VbaStackTraceLogger.Core.Writer")
'@Interface
'@ModuleDescription("Interface for Log Writer")

Option Explicit

'@Description("Initialize the log writer")
Public Sub Initialize()
End Sub

'@Description("Write log using DTO")
Public Sub WriteLog(ByVal logRequest As Logger_LogWriteRequestDTO)
End Sub

'@Description("Terminate the log writer")
Public Sub Terminate()
End Sub
