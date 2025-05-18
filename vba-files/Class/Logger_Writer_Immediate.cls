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
Option Explicit
Implements Logger_ILogWriter

Private Sub Logger_ILogWriter_Initialize()
    Debug.Print "イミディエイトウィンドウ用のロガーを初期化しました"
End Sub

Private Sub Logger_ILogWriter_WriteLog(logInfo As LoggerLogInfoStruct)
    Dim outputContext As String
    
    If logInfo.IsIncludeCallerInfo Then
        If Len(logInfo.CallPath) > 0 Then
            outputContext = "[" & logInfo.CallPath & "]"
        ElseIf Len(logInfo.ModuleName) > 0 And Len(logInfo.ProcedureName) > 0 Then
            ' Fallback for direct Log calls if CallPath is empty but caller info is requested
            outputContext = "[" & logInfo.ModuleName & "." & logInfo.ProcedureName & "]"
        Else
            outputContext = ""
        End If
    Else
        outputContext = ""
    End If

    Debug.Print "[" & logInfo.Timestamp & "]" & _
                "[" & logInfo.TagName & "]" & _
                IIf(Len(outputContext) > 0, " " & outputContext, "") & _
                IIf(Len(logInfo.Message) > 0, " " & logInfo.Message, "")
End Sub

Private Sub Logger_ILogWriter_Terminate()
    ' 特になし
End Sub
