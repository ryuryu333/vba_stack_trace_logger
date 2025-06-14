VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_ConfigBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Logger_ConfigBuilder.cls
'@Folder("VbaStackTraceLogger.Core.Config")

Option Explicit

Implements Logger_IConfigBuilder

Private Const CLASS_NAME As String = "Logger_ConfigBuilder"

Private mConfig As Logger_ConfigurationDTO
Private mReceiver As Logger_IConfigReceiver
Private mIsBuilt As Boolean

' Specify initial values for logger settings
' Settings are configured to enable only essential minimum functionality
Private Sub Class_Initialize()
    Set mConfig = New Logger_ConfigurationDTO
    With mConfig
        .IsLoggingEnabled = True
        .IsTagFilteringEnabled = False
        .IsStackTraceEnabled = False
        .IsWriteToImmediate = True
        .IsWriteToExcelSheet = False
        Set .OutputExcelSheet = Nothing
    End With
End Sub

' Configuration methods with method chaining
Public Function DisableLogging() As Logger_ConfigBuilder
    mConfig.IsLoggingEnabled = False
    Set DisableLogging = Me
End Function

Public Function EnableTagFiltering() As Logger_TagFilterBuilder
    mConfig.IsTagFilteringEnabled = True
    ' Chain to sub-builder and force ExcludedTags settings
    Dim nextStep As New Logger_TagFilterBuilder
    nextStep.SetParentBuilder Me
    Set EnableTagFiltering = nextStep
End Function

' Called by sub-builder Logger_TagArrayBuilder
Friend Sub SetExcludedTags(ByRef tags() As LoggerLogTag)
    Dim createdArray() As LoggerLogTag
    createdArray = tags
    mConfig.ExcludedTags = createdArray
End Sub

Public Function EnableStackTrace() As Logger_ConfigBuilder
    mConfig.IsStackTraceEnabled = True
    Set EnableStackTrace = Me
End Function

Public Function DisableWriteToImmediate() As Logger_ConfigBuilder
    mConfig.IsWriteToImmediate = False
    Set DisableWriteToImmediate = Me
End Function

Public Function EnableWriteToExcelSheet() As Logger_ExcelOutputBuilder
    mConfig.IsWriteToExcelSheet = True
    ' Chain to sub-builder
    Dim nextStep As New Logger_ExcelOutputBuilder
    nextStep.SetParentBuilder Me
    Set EnableWriteToExcelSheet = nextStep
End Function

' Called from sub-builder Logger_ExcelOutputBuilder
Friend Sub SetOutputExcelSheet(ByVal sheet As Worksheet)
    Set mConfig.OutputExcelSheet = sheet
End Sub

' Finalize configuration and call initialization
Public Sub Build()
    EnsureInitialized
    mIsBuilt = True
    mReceiver.Initialize mConfig
End Sub

Private Sub EnsureBuild()
    If Not mIsBuilt Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".EnsureBuild", _
            "Logger initialization is not complete. Call Build() to finalize configuration."
    End If
End Sub

Private Sub Class_Terminate()
    EnsureBuild
End Sub

' Ensure initialized before operations
Private Sub EnsureInitialized()
    If Not mIsInitialized Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".EnsureInitialized", _
            "Unexpected error. Not initialized."
    End If
End Sub

' === Interface Implementation ===

' Logger_IConfigBuilder implementation
Private Sub Logger_IConfigBuilder_Initialize(ByVal configReceiver As Logger_IConfigReceiver)
    If configReceiver is Nothing Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".Logger_IConfigBuilder_Initialize", _
            "Unexpected error. configReceiver is Nothing."
    End If
    Set mReceiver = configReceiver
    mIsInitialized = True
End Sub
