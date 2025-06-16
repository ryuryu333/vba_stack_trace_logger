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

Option Explicit

Private Const CLASS_NAME As String = "Logger_Facade"

Private mIsInitialized As Boolean
Private mController As Logger_IController

' === Logger initialization ===
' Start configuration, user set options via methode cahin
' Finally call back Controller's initialize method
Public Function StartConfiguration() As Logger_ConfigBuilder
    ' If already initialized, automatically reset and reinitialize
    If mIsInitialized Then
        Terminate
        mIsInitialized = False
    End If
    
    ' Create concrete class instance
    Dim concreteController As New Logger_Controller ' Used by Facade & Builder
    Dim concreteConfigBuilder As New Logger_ConfigBuilder
    
    ' Set member variable
    Set mController = concreteController

    ' Preparation builder
    Dim configBuilder As Logger_IConfigBuilder
    Set configBuilder = concreteConfigBuilder

    ' Set call back receiver
    Dim receiver As Logger_IConfigReceiver
    Set receiver = concreteController
    configBuilder.Initialize receiver

    ' Finish
    mIsInitialized = True
    Set StartConfiguration = configBuilder
End Function

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
Public Function UsingTracer(ByVal currentModuleName As String, _
                            ByVal currentProcName As String) As Logger_ProcedureTracer
    EnsureInitialized
    
    Set UsingTracer = mController.UsingTracer(currentModuleName, currentProcName)
End Function

Private Sub Class_Terminate()
    ' Fail-safe
    Set mController = Nothing
End Sub

' === Private Method ===
' Ensure initialized before operations
Private Sub EnsureInitialized()
    If Not mIsInitialized Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".EnsureInitialized", _
            "Logger is not initialized. Call StartConfiguration().Build() first."
    End If
End Sub

