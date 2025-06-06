VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_TagFilterBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Logger_TagFilterBuilder.cls
Option Explicit
Private Const CLASS_NAME As String = "TagFilterBuilder"

Private parentBuilder As Logger_ConfigBuilder
Private tags() As LoggerLogTag
Private count As Long
Private mIsApplied As Boolean

Private Sub Class_Initialize()
    ReDim tags(0)
    count = 0
End Sub

Friend Sub SetParentBuilder(ByVal parent As Logger_ConfigBuilder)
    Set parentBuilder = parent
End Sub

Public Function Add(tag As LoggerLogTag) As Logger_TagFilterBuilder
    ReDim Preserve tags(0 To count)
    tags(count) = tag
    count = count + 1
    Set Add = Me
End Function

Public Function Apply() As Logger_ConfigBuilder
    parentBuilder.SetExcludedTags = tags
    mIsApplied = True
    Set Apply = parentBuilder
End Function

Private Sub Class_Terminate()
    If Not mIsApplied Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
                  LOGGER_NAMESPACE & "." & CLASS_NAME & ".Class_Terminate", _
                  "Apply() が呼ばれていません、タグフィルタの設定が適用されていないため正常に動作しません"
    End If
End Sub
