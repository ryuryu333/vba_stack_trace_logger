VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_Facade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Logger_Facade.cls
'@Folder("VbaStackTraceLogger.Core.API")

' User interface aggregation class
' Users access Logger_Facade via Logger_Provider
' (Actual flow: Logger_Provider -> Logger_SingletonManager -> Logger_Facade)
' Usage: myLogger.Log

Option Explicit
Private Const CLASS_NAME As String = "Logger_Facade"

Private mController As Logger_Controller
Private mIsInitialized As Boolean

Private Sub Class_Initialize()
    Set mController = New Logger_Controller
End Sub

' === Logger initialization ===
' Processing order: User > Facade.Initialize > ConfigBuilder.Build > Facade.CompleteInitialization
Public Function Initialize() As Logger_ConfigBuilder
    If mIsInitialized Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".Initialize", _
            "Logger is already initialized, second initialization cannot be executed"
        Exit Function
    End If

    ' Create configuration values using builder
    ' Then returns to CompleteInitialization
    Dim configBuilder As New Logger_ConfigBuilder
    configBuilder.Initialize Me
    Set Initialize = configBuilder
End Function

' Processing order: User > Facade.Initialize > ConfigBuilder.Build > Facade.CompleteInitialization
Public Sub CompleteInitialization(ByRef userConfig As Logger_ConfigStruct)
    If mIsInitialized Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".Initialize", _
            "Logger is already initialized, second initialization cannot be executed"
    End If
    
    ' Initialize Controller using user-created configuration values
    mController.Initialize userConfig
    
    mIsInitialized = True
End Sub

' === Log output instruction ===
Public Sub Log(ByVal writeMessage As String, Optional ByVal userSelectedTag As LoggerLogTag = LoggerLogTag.LogTag_Info)
    EnsureInitialized

    mController.WriteLog writeMessage, userSelectedTag
End Sub

' === Logger termination instruction ===
Public Sub Terminate()
    EnsureInitialized
    
    mController.Terminate
    Set mController = Nothing
    
    mIsInitialized = False
End Sub

' === Stack trace related ===
' ProcedureTracerProvider > Create ProcedureTracer > User get ProcedureTracer instance
' User holds ProcedureTracer as variable at procedure start, variable is destroyed when procedure ends
' Utilizes instance lifecycle to detect procedure start/end and execute processing
Public Function UsingTracer(ByVal currentModuleName As String, _
                            ByVal currentProcName As String) As Logger_ProcedureTracer
    EnsureInitialized
    
    Set UsingTracer = mController.UsingTracer(currentModuleName, currentProcName)
End Function

' === Utility functions ===
' Destroy Logger_Facade instance in Logger_SingletonManager
Public Sub Reset()
    ReleaseMyLogger
End Sub

' === Private functions used throughout the module ===
' Check if initialized
Private Sub EnsureInitialized()
    If Not mIsInitialized Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".EnsureInitialized", _
            "Logger is not initialized"
    End If
End Sub

Private Sub Class_Terminate()
    ' Fail-safe
    Set mController = Nothing
End Sub


