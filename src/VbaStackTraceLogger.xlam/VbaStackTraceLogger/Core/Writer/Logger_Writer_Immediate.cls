VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_Writer_Immediate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Logger_ImmediateWriter.cls
'@Folder("VbaStackTraceLogger.Core.Writer")

Option Explicit
Implements Logger_ILogWriter

Private Sub Logger_ILogWriter_Initialize()
    Debug.Print "Finished initializing Immediate Window writer."
End Sub

Private Sub Logger_ILogWriter_WriteLog(ByVal logRequest As Logger_LogWriteRequestDTO)
    Debug.Print "[" & logRequest.Timestamp & "]" & _
                "[" & logRequest.TagName & "]" & _
                IIf(Len(logRequest.CallPath) > 0, "[" & logRequest.CallPath & "]", "") & _
                IIf(Len(logRequest.Message) > 0, " " & logRequest.Message, "")
End Sub

Private Sub Logger_ILogWriter_Terminate()
    ' Nnothing to do for Immediate Window writer
End Sub
