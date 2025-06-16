VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_StackTraceController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Logger_StackTraceController.cls
'@Folder("VbaStackTraceLogger.StackTrace")

Option Explicit

Implements Logger_IStackTraceFacade
Implements Logger_IStackTraceEventHandler
Private Const CLASS_NAME As String = "Logger_StackTraceController"

' Updated to match class diagram design
Private mCallStack As Logger_ICallStack
Private mTraceLogOutput As Logger_ITraceLogOutput
Private mIsStackTraceEnabled As Boolean
Private mIsInitialized As Boolean

Private Sub Class_Initialize()
    Set mCallStack = New Logger_CallStack
End Sub

' === Initialization ===
Private Sub Initialize(ByVal IsStackTraceEnabled As Boolean, ByVal traceLogOutput As Logger_ITraceLogOutput)
    If traceLogOutput Is Nothing Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".Initialize", _
            "Unexpected error. TraceLogOutput instance not found"
    End If
    
    Set mTraceLogOutput = traceLogOutput
    mIsStackTraceEnabled = IsStackTraceEnabled
    
    mIsInitialized = True
End Sub

' === Stack Trace ===
Private Function UsingTracer(ByVal currentModuleName As String, _
                            ByVal currentProcName As String) As Logger_ProcedureTracer
    EnsureInitialized
    
    ' If stack trace is disabled, return Nothing
    If Not mIsStackTraceEnabled Then
        Set UsingTracer = Nothing
        Exit Function
    End If
    
    ' Create tracer
    Set UsingTracer = CreateTracer(currentModuleName, currentProcName)
End Function

' Create procedure tracer
Private Function CreateTracer(ByVal ModuleName As String, ByVal ProcedureName As String) As Logger_ProcedureTracer
    Dim tracer As Logger_ProcedureTracer
    Set tracer = New Logger_ProcedureTracer
    tracer.Initialize ModuleName, ProcedureName, Me
    Set CreateTracer = tracer
End Function
End Function

' === Get current call path ===
Private Function GetCurrentCallPath() As String
    EnsureInitialized
    GetCurrentCallPath = mCallStackManager.GetCurrentCallPath()
End Function

' === Callback from ProcedureTracer ===
' Event handlers called by ProcedureTracer
Private Sub NotifyProcedureEnter(ByVal callSite As String)
    OnProcedureEnter callSite
End Sub

Private Sub NotifyProcedureExit(ByVal callSite As String)
    OnProcedureExit
End Sub

' Internal event handling
Private Sub OnProcedureEnter(ByVal callSite As String)
    EnsureInitialized
    
    ' 1. Update call stack
    mCallStack.EnterProcedure callSite
    ' 2. Request log output
    mTraceLogOutput.WriteTraceLog ">> Enter " & callSite, LogTag_Trace
End Sub

Private Sub OnProcedureExit()
    EnsureInitialized
    
    ' 1. Get current call site before exit
    Dim currentCallPath As String
    currentCallPath = mCallStack.GetCurrentCallPath
    
    ' 2. Request log output
    If Len(currentCallPath) > 0 Then
        mTraceLogOutput.WriteTraceLog "<< Exit " & currentCallPath, LogTag_Trace
    End If
    
    ' 3. Update call stack
    mCallStack.ExitProcedure
End Sub

' === Cleanup ===
Private Sub Terminate()
    EnsureInitialized
    
    Set mCallStack = Nothing
    Set mTraceLogOutput = Nothing
    mIsInitialized = False
End Sub

' Check if initialized
Private Sub EnsureInitialized()
    If Not mIsInitialized Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".EnsureInitialized", _
            "StackTraceController is not initialized"
    End If
End Sub

Private Sub Class_Terminate()
    ' Fail-safe
    Set mCallStackManager = Nothing
    Set mProcedureTracerProvider = Nothing
    Set mMainController = Nothing
End Sub

' === Interface Implementations ===

' Logger_IStackTraceFacade implementation
Private Sub Logger_IStackTraceFacade_Initialize(ByVal IsStackTraceEnabled As Boolean, ByVal traceLogOutput As Logger_ITraceLogOutput)
    Initialize IsStackTraceEnabled, traceLogOutput
End Sub

Private Function Logger_IStackTraceFacade_UsingTracer(ByVal ModuleName As String, ByVal ProcedureName As String) As Logger_ProcedureTracer
    Set Logger_IStackTraceFacade_UsingTracer = UsingTracer(ModuleName, ProcedureName)
End Function

Private Function Logger_IStackTraceFacade_GetCurrentCallPath() As String
    Logger_IStackTraceFacade_GetCurrentCallPath = GetCurrentCallPath
End Function

Private Sub Logger_IStackTraceFacade_Terminate()
    Terminate
End Sub

' Logger_IStackTraceEventHandler implementation
Private Sub Logger_IStackTraceEventHandler_OnProcedureEnter(ByVal callSite As String)
    OnProcedureEnter callSite
End Sub

Private Sub Logger_IStackTraceEventHandler_OnProcedureExit()
    OnProcedureExit
End Sub

