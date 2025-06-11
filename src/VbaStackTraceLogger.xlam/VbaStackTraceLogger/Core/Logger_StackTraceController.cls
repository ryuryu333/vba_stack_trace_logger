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
'@Folder("VbaStackTraceLogger.Core")

Option Explicit
Private Const CLASS_NAME As String = "Logger_StackTraceController"

Private mCallStackManager As Logger_CallStackManager
Private mProcedureTracerProvider As Logger_ProcedureTracerProvider
Private mMainController As Logger_Controller
Private mIsInitialized As Boolean

Private Sub Class_Initialize()
    Set mCallStackManager = New Logger_CallStackManager
    Set mProcedureTracerProvider = New Logger_ProcedureTracerProvider
End Sub

' === Initialization ===
Public Sub Initialize(ByVal IsStackTraceEnabled As Boolean, ByRef mainController As Logger_Controller)
    If mainController Is Nothing Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".Initialize", _
            "Main controller instance not found"
    End If
    
    Set mMainController = mainController
    
    ' === Initialize ProcedureTracerProvider ===
    mProcedureTracerProvider.Initialize IsStackTraceEnabled
    
    mIsInitialized = True
End Sub

' === Stack Trace ===
Public Function UsingTracer(ByVal currentModuleName As String, _
                            ByVal currentProcName As String) As Logger_ProcedureTracer
    EnsureInitialized
    ' Request Tracer creation from Provider
    Set UsingTracer = mProcedureTracerProvider.CreateTracer(currentModuleName, currentProcName, Me)
End Function

' === Get CallStackManager reference ===
Public Function GetCallStackManager() As Logger_CallStackManager
    EnsureInitialized
    Set GetCallStackManager = mCallStackManager
End Function

' === Callback from ProcedureTracer ===
Public Sub NotifyProcedureEnter(ByVal callSite As String)
    EnsureInitialized
    
    ' 1. Update call stack
    mCallStackManager.EnterProcedure callSite
    ' 2. Request log output
    mMainController.WriteLog ">> Enter " & callSite, LogTag_Trace
End Sub

Public Sub NotifyProcedureExit(ByVal callSite As String)
    EnsureInitialized
    
    ' 1. Request log output
    mMainController.WriteLog "<< Exit " & callSite, LogTag_Trace
    ' 2. Update call stack
    mCallStackManager.ExitProcedure
End Sub

' === Cleanup ===
Public Sub Terminate()
    EnsureInitialized
    
    Set mProcedureTracerProvider = Nothing
    Set mCallStackManager = Nothing
    Set mMainController = Nothing
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
