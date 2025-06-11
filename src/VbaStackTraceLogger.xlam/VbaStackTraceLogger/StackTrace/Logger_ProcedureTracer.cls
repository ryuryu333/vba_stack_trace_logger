VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_ProcedureTracer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Logger_ProcedureTracer.cls
'@Folder("VbaStackTraceLogger.StackTrace")

' Mimics RAII pattern, updates stack trace information tied to lifecycle
' Class used as variable
Option Explicit
Private Const CLASS_NAME As String = "Logger_ProcedureTracer"

Private mCallSite As String

Private mStackTraceController As Logger_StackTraceController
Private mIsInitialized As Boolean

' Initialization is instructed after instantiation
' Substitute for parameterized constructor
Public Sub Initialize(ByVal currentModuleName As String, ByVal currentProcedureName As String, _
                      ByRef stackTraceController As Logger_StackTraceController)
    If stackTraceController Is Nothing Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".Initialize", _
            "StackTraceController instance not found"
    End If
    
    Set mStackTraceController = stackTraceController
    mCallSite = currentModuleName & "." & currentProcedureName
    
    ' Notify procedure entry to StackTraceController
    mStackTraceController.NotifyProcedureEnter mCallSite
    
    mIsInitialized = True
End Sub

' Destructor
Private Sub Class_Terminate()
    If mIsInitialized = False Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".Class_Terminate", _
            "Cleanup called without initialization"
    End If
    
    ' Notify procedure exit to StackTraceController
    mStackTraceController.NotifyProcedureExit mCallSite
    
    ' Release resources
    Set mStackTraceController = Nothing
End Sub


