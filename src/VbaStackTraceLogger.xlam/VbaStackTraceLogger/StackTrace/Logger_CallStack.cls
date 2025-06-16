VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_CallStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Logger_CallStack.cls
'@Folder("VbaStackTraceLogger.StackTrace")

Option Explicit

Implements Logger_ICallStack
Private Const CLASS_NAME As String = "Logger_CallStack"

Private mCallNestLevel As Long
Private mCallSite As Collection
Private mCallPath As Collection
' e.g. When call FugaCls.FugaProc() in HogeCls.HogeProc()
' mCallNestLevel = 2
' | mCallSite(1) = HogeCls.HogeProc | mCallPath(1) = HogeCls.HogeProc |
' | mCallSite(2) = FugaCls.FugaProc | mCallPath(2) = FugaCls.FugaProc < HogeCls.HogeProc |

Private Sub Class_Initialize()
    Set mCallSite = New Collection
    Set mCallPath = New Collection
    mCallNestLevel = 0
End Sub

Private Sub EnterProcedure(ByVal callSite As String)
    Dim newCallSite As String
    newCallSite = callSite
    Dim newCallPath As String
    If mCallNestLevel > 0 Then
        newCallPath = newCallSite & " < " & mCallPath(mCallNestLevel)
    Else
        newCallPath = newCallSite
    End If
    ' Update call stack information
    mCallSite.Add newCallSite
    mCallPath.Add newCallPath
    mCallNestLevel = mCallNestLevel + 1
    ValidateCallStack
End Sub

Private Sub ExitProcedure()
    If mCallNestLevel = 0 Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".ExitProcedure", _
            "Unexpected error. ExitProcedure() was called before EnterProcedure()."
    End If
    ' Update call stack information
    mCallSite.Remove mCallNestLevel
    mCallPath.Remove mCallNestLevel
    mCallNestLevel = mCallNestLevel - 1
    ValidateCallStack
End Sub

Private Function GetCurrentCallSite() As String
    If mCallNestLevel = 0 Then
        GetCurrentCallSite = ""
    Else
        GetCurrentCallSite = mCallSite(mCallNestLevel)
    End If
End Function

Private Function GetCurrentCallPath() As String
    If mCallNestLevel = 0 Then
        GetCurrentCallPath = ""
    Else
        GetCurrentCallPath = mCallPath(mCallNestLevel)
    End If
End Function

Private Sub ValidateCallStack()
    If mCallNestLevel < 0 Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".ValidateCallStack", _
            "Unexpected error. CallNestLevel is under 0."
    End If
    
    If mCallSite.Count <> mCallNestLevel Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".ValidateCallStack", _
            "Unexpected error. CallSite.Count is not equal CallNestLevel."
    End If
    
    If mCallPath.Count <> mCallNestLevel Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".ValidateCallStack", _
            "Unexpected error. mCallPath.Count is not equal CallNestLevel."
    End If
End Sub

Private Sub Class_Terminate()
    Set mCallSite = Nothing
    Set mCallPath = Nothing
End Sub

' === Interface Implementation ===

' Logger_ICallStack implementation
Private Sub Logger_ICallStack_EnterProcedure(ByVal callSite As String)
    EnterProcedure callSite
End Sub

Private Sub Logger_ICallStack_ExitProcedure()
    ExitProcedure
End Sub

Private Function Logger_ICallStack_GetCurrentCallSite() As String
    Logger_ICallStack_GetCurrentCallSite = GetCurrentCallSite
End Function

Private Function Logger_ICallStack_GetCurrentCallPath() As String
    Logger_ICallStack_GetCurrentCallPath = GetCurrentCallPath
End Function

