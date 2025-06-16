VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_Controller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Logger_Controller.cls
'@Folder("VbaStackTraceLogger.Core.Engine")

Option Explicit

Implements Logger_IController
Implements Logger_IConfigReceiver
Implements Logger_ITraceLogOutput
Private Const CLASS_NAME As String = "Logger_Controller"

Private mStackTraceController As Logger_IStackTraceFacade
Private mWriteLogController As Logger_IWriteLogController
Private mOutputFilter As Logger_IOutputFilter
Private mLoggerConfig As Logger_ConfigurationDTO
Private mIsInitialized As Boolean

' === Initialization processing ===
Private Sub Initialize()
    ' Use configuration received from Logger_ConfigBuilder
    ' === Create LoggerConfig ===
    Set mOutputFilter = New Logger_OutputFilter

    ' === Create and Initialize StackTraceController ===
    Set mStackTraceController = New Logger_StackTraceController
    mStackTraceController.Initialize mLoggerConfig.IsStackTraceEnabled, Me
      ' === Create and initialize Writers ===
    Dim writerFactory As Logger_WriterFactory
    Set writerFactory = New Logger_WriterFactory
    
    Dim writers As Collection
    Set writers = writerFactory.CreateWriters(mLoggerConfig.IsWriteToImmediate, _
                                              mLoggerConfig.IsWriteToExcelSheet, _
                                              mLoggerConfig.OutputExcelSheet)
    
    Set mWriteLogController = New Logger_WriteLogController
    mWriteLogController.Initialize writers
    
    ' Finish
    mIsInitialized = True
End Sub

' === Log output ===
' Write information to logInfo, then pass to Writer for log output
Private Sub WriteLog(ByVal writeMessage As String, ByVal userSelectedTag As LoggerLogTag)
    EnsureInitialized
    ' === Based on logger settings, determine if log should be output ===
    If mOutputFilter.ShouldWriteLog(userSelectedTag, _
                                    mLoggerConfig.IsLoggingEnabled, _
                                    mLoggerConfig.IsTagFilteringEnabled, _
                                    mLoggerConfig.ExcludedTags) _
                                        = False Then Exit Sub
    
    ' === Get call path information ===
    Dim CallPath As String
    If mLoggerConfig.IsStackTraceEnabled Then
        CallPath = mStackTraceController.GetCurrentCallPath()
    Else
        CallPath = ""
    End If
    
    ' === Use WriteLogController to output log ===
    mWriteLogController.WriteLogToAll writeMessage, userSelectedTag, CallPath
End Sub

' === Stack trace ===
Private Function UsingTracer(ByVal currentModuleName As String, _
                            ByVal currentProcName As String) As Logger_ProcedureTracer
    EnsureInitialized
    
    ' Request Tracer creation from StackTraceController
    Set UsingTracer = mStackTraceController.UsingTracer(currentModuleName, currentProcName)
End Function

' === Termination processing ===
Private Sub Terminate()
    EnsureInitialized
    mWriteLogController.Terminate
    mStackTraceController.Terminate
    Set mStackTraceController = Nothing
    Set mWriteLogController = Nothing
    Set mOutputFilter = Nothing
    Set mLoggerConfig = Nothing
End Sub

' Check if initialized
Private Sub EnsureInitialized()
    If Not mIsInitialized Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".EnsureInitialized", _
            "Controller is not initialized"
    End If
End Sub

Private Sub Class_Terminate()
    ' Fail-safe
    If Not mWriteLogController Is Nothing Then mWriteLogController.Terminate
    If Not mStackTraceController Is Nothing Then mStackTraceController.Terminate
    Set mStackTraceController = Nothing
    Set mWriteLogController = Nothing
    Set mOutputFilter = Nothing
    Set mLoggerConfig = Nothing
End Sub

' === Interface Implementations ===

' Logger_IController implementation
Private Sub Logger_IController_WriteLog(ByVal writeMessage As String, ByVal logTag As LoggerLogTag)
    WriteLog writeMessage, logTag
End Sub

Private Function Logger_IController_UsingTracer(ByVal ModuleName As String, ByVal ProcedureName As String) As Logger_ProcedureTracer
    Set Logger_IController_UsingTracer = UsingTracer(ModuleName, ProcedureName)
End Function

Private Sub Logger_IController_Terminate()
    Terminate
End Sub

' Logger_IConfigReceiver implementation
Private Sub Logger_IConfigReceiver_Initialize(ByVal config As Logger_ConfigurationDTO)
    ' Recive configuration from Logger_ConfigBuilder
    Set mLoggerConfig = config
    ' start initialization
    Initialize
End Sub

' Logger_ITraceLogOutput implementation
Private Sub Logger_ITraceLogOutput_WriteTraceLog(ByVal writeMessage As String, ByVal logTag As LoggerLogTag)
    WriteLog writeMessage, logTag
End Sub

