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
'@Folder("VbaStackTraceLogger.Core.Config")

Option Explicit
Private Const CLASS_NAME As String = "Logger_ExcelOutputBuilder"

Private mParentBuilder As Logger_ConfigBuilder
Private mIsApplied As Boolean

Friend Sub SetParentBuilder(ByVal parent As Logger_ConfigBuilder)
    Set mParentBuilder = parent
End Sub

Public Function SetOutputExcelSheet(ByVal sheet As Worksheet) As Logger_ConfigBuilder
    If mParentBuilder Is Nothing Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".SetOutputExcelSheet", _
            "Unexpected error. mParentBuilder is Nothing."
    End If

    mParentBuilder.SetOutputExcelSheet sheet
    mIsApplied = True
    Set SetOutputExcelSheet = mParentBuilder
End Function

Private Sub EnsureApply()
    If Not mIsApplied Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".EnsureApply", _
             "Logger initialization is not complete. Call SetOutputExcelSheet()."
    End If
End Sub

Private Sub Class_Terminate()
    EnsureApply
End Sub
