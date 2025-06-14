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
'@Folder("VbaStackTraceLogger.Core.Config")

Option Explicit
Private Const CLASS_NAME As String = "Logger_TagFilterBuilder"

Private mParentBuilder As Logger_ConfigBuilder
Private mTags() As LoggerLogTag
Private mCount As Long
Private mIsAdded As Boolean
Private mIsApplied As Boolean

Private Sub Class_Initialize()
    ReDim mTags(1 To 1)
    mCount = 0
End Sub

Friend Sub SetParentBuilder(ByVal parent As Logger_ConfigBuilder)
    Set mParentBuilder = parent
End Sub

Public Function Add(ByVal tag As LoggerLogTag) As Logger_TagFilterBuilder
    EnsureAdd
    mCount = mCount + 1
    ReDim Preserve mTags(1 To mCount)
    mTags(mCount) = tag
    
    mIsAdded = True
    Set Add = Me
End Function

Public Function Apply() As Logger_ConfigBuilder
    EnsureAdd
    
    mParentBuilder.SetExcludedTags mTags
    mIsApplied = True
    Set Apply = mParentBuilder
End Function

Private Sub EnsureAdd()
    If Not mIsAdded Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".EnsureAdd", _
            "Logger initialization is not complete. Call Add()"
    End If
End Sub

Private Sub EnsureApply()
    If Not mIsApplied Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".EnsureApply", _
            "Logger initialization is not complete. Call Apply()"
    End If
End Sub

Private Sub Class_Terminate()
    EnsureApply
End Sub
