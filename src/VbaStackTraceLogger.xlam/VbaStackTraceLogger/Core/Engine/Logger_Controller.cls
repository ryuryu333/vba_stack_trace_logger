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
'@Folder("VbaStackTraceLogger.Core")

Option Explicit
Private Const CLASS_NAME As String = "Logger_Controller"

Private mStackTraceController As Logger_StackTraceController
Private mLogInfoMetadataHelper As Logger_LogInfoMetadataHelper
Private mWriterManager As Logger_WriterManager
Private mOutputFilter As Logger_OutputFilter
Private mLoggerConfig As Logger_ConfigStruct
Private mIsInitialized As Boolean

Private Sub Class_Initialize()
    Set mStackTraceController = New Logger_StackTraceController
    Set mLogInfoMetadataHelper = New Logger_LogInfoMetadataHelper
    Set mWriterManager = New Logger_WriterManager
    Set mOutputFilter = New Logger_OutputFilter
End Sub

' === Initialization processing ===
Public Sub Initialize(ByRef userConfig As Logger_ConfigStruct)
    ' === Validate each argument, then assign to member variables ===
    ValidateConfig userConfig
    Set mLoggerConfig = userConfig
    ' === Initialize StackTraceController ===
    mStackTraceController.Initialize userConfig.IsStackTraceEnabled, Me
    
    ' === Create and initialize Writers ===
    mWriterManager.CreateWriters userConfig
    mWriterManager.InitializeWriters
    
    mIsInitialized = True
End Sub

' Execute dynamic validation, static validation is performed in Logger_ConfigBuilder
Private Sub ValidateConfig(ByRef userConfig As Logger_ConfigStruct)
    ' Validate Excel output settings
    If userConfig.IsWriteToExcelSheet Then
        ' Access Excel sheet and verify it's accessible
        On Error Resume Next
        Dim testValue As String
        testValue = userConfig.OutputExcelSheet.Name
        If Err.Number <> 0 Then
            On Error GoTo 0
            Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
                LOGGER_NAMESPACE & "." & CLASS_NAME & ".ValidateConfig", _
                "Cannot access the specified Excel sheet"
        End If
        On Error GoTo 0
    End If
End Sub

' === Log output ===
' Write information to logInfo, then pass to Writer for log output
Public Sub WriteLog(ByVal writeMessage As String, ByVal userSelectedTag As LoggerLogTag)
    EnsureInitialized
    ' === Based on logger settings, determine if log should be output ===
    If mOutputFilter.ShouldWriteLog(userSelectedTag, mLoggerConfig) = False Then Exit Sub
    ' === Write information to logInfo ===
    Dim CallPath As String
    If mLoggerConfig.IsStackTraceEnabled Then
        CallPath = mStackTraceController.GetCurrentCallPath()
    Else
        CallPath = ""
    End If
    
    Dim logInfo As Logger_LogInfoStruct
    Set logInfo = mLogInfoMetadataHelper.BuildLogInfo(writeMessage, userSelectedTag, mLoggerConfig, CallPath)
    
    ' === Command each Writer to output log ===
    mWriterManager.WriteLogToAll logInfo
End Sub

' === Stack trace ===
Public Function UsingTracer(ByVal currentModuleName As String, _
                            ByVal currentProcName As String) As Logger_ProcedureTracer
    EnsureInitialized
    
    ' Request Tracer creation from StackTraceController
    Set UsingTracer = mStackTraceController.UsingTracer(currentModuleName, currentProcName)
End Function


' === Termination processing ===
Public Sub Terminate()
    EnsureInitialized
      mWriterManager.TerminateWriters
    mStackTraceController.Terminate
    Set mLogInfoMetadataHelper = Nothing
    Set mStackTraceController = Nothing
    Set mWriterManager = Nothing
    Set mOutputFilter = Nothing
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
    If Not mWriterManager Is Nothing Then
        mWriterManager.TerminateWriters
    End If
    Set mStackTraceController = Nothing
    Set mLogInfoMetadataHelper = Nothing
    Set mWriterManager = Nothing
    Set mOutputFilter = Nothing
End Sub


