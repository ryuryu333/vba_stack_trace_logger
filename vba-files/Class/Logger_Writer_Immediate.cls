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
    If logInfo.IsIncludeCallerInfo Then
        Debug.Print "[" & logInfo.Timestamp & "]" & _
                    "[" & logInfo.TagName & "]" & _
                    "[" & logInfo.ModuleName & "." & logInfo.ProcedureName & "] " & _
                    logInfo.Message
    Else
        Debug.Print "[" & logInfo.Timestamp & "]" & _
                    "[" & logInfo.TagName & "]" & _
                    logInfo.Message
    End If
End Sub

Private Sub Logger_ILogWriter_Terminate()
    ' 特になし
End Sub
