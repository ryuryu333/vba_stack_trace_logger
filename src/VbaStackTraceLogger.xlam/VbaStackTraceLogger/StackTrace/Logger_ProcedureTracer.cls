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

' Automatically tracks procedure entry and exit
' NOTE:
'  This is based on the RAII (Resource Acquisition Is Initialization) concept.
'  User stores an instance of this class as a variable at the beginning of the procedure.
'    e.g. Dim scopeGuard As Variant: Set scopeGuard = MyLogger.UsingTracer("ModuleName", "ProcedureName")
'  By calling UsingTracer(), the instance is created and initialized internally.
'    -> At that time (= constructor), this class notifies the start-of-procedure event.
'  When the user's procedure ends, the reference to the variable drops to zero,
'  and the variable is automatically destroyed by the system.
'    -> At that time (= destructor), this class notifies the end-of-procedure event.

Option Explicit
Private Const CLASS_NAME As String = "Logger_ProcedureTracer"

Private mStackTraceEventHandler As Logger_IStackTraceEventHandler
Private mIsInitialized As Boolean

' Initialization is instructed after instantiation
Public Sub Initialize(ByVal currentModuleName As String, ByVal currentProcedureName As String, _
                      ByVal stackTraceEventHandler As Logger_IStackTraceEventHandler)
    If stackTraceEventHandler Is Nothing Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".Initialize", _
            "Unexpected error. StackTraceEventHandler instance not found"
    End If
    Set mStackTraceEventHandler = stackTraceEventHandler
    
    ' Notify procedure entry to StackTraceEventHandler
    mStackTraceEventHandler.OnProcedureEnter currentModuleName & "." & currentProcedureName
    
    mIsInitialized = True
End Sub

' Destructor
Private Sub Class_Terminate()
    If mIsInitialized = False Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".Class_Terminate", _
            "Unexpected error. Destructor called without initialization"
    End If
    
    ' Notify procedure exit to StackTraceEventHandler
    mStackTraceEventHandler.OnProcedureExit
    
    ' Release resources
    Set mStackTraceEventHandler = Nothing
End Sub


