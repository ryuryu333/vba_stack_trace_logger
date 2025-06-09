VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_ExcelOutputBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Logger_ExcelOutputBuilder.cls
'@Folder("VbaStackTraceLogger.Config")

Option Explicit
Private Const CLASS_NAME As String = "ExcelOutputBuilder"

Private parentBuilder As Logger_ConfigBuilder
Private mIsApplied As Boolean

Friend Sub SetParentBuilder(ByVal parent As Logger_ConfigBuilder)
    Set parentBuilder = parent
End Sub

Public Function SetOutputExcelSheet(sheet As Worksheet) As Logger_ConfigBuilder
    parentBuilder.SetOutputExcelSheet sheet
    mIsApplied = True
    Set SetOutputExcelSheet = parentBuilder
End Function

Private Sub Class_Terminate()
    If Not mIsApplied Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
                  LOGGER_NAMESPACE & "." & CLASS_NAME & ".Class_Terminate", _
                  "SetOutputExcelSheet() が呼ばれていません、Excel出力の設定が適用されていないため正常に動作しません"
    End If
End Sub
