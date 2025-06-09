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
'@Folder("VbaStackTraceLogger.Config")

Option Explicit
Private Const CLASS_NAME As String = "TagFilterBuilder"

Private parentBuilder As Logger_ConfigBuilder
Private mTags() As LoggerLogTag
Private mCount As Long
Private mIsAdded As Boolean
Private mIsApplied As Boolean

Private Sub Class_Initialize()
    ReDim mTags(1 To 1)
    mCount = 0
End Sub

Friend Sub SetParentBuilder(ByVal parent As Logger_ConfigBuilder)
    Set parentBuilder = parent
End Sub

Public Function Add(tag As LoggerLogTag) As Logger_TagFilterBuilder
    mCount = mCount + 1
    ReDim Preserve mTags(1 To mCount)
    mTags(mCount) = tag
    
    mIsAdded = True
    Set Add = Me
End Function

Public Function Apply() As Logger_ConfigBuilder
    If mIsAdded = False Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".Apply", _
            "除外対象のタグが指定されていません、Add で指定した後に Apply してください"
    End If
    
    parentBuilder.SetExcludedTags mTags
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
