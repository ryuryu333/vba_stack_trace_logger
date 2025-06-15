VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_WriteLogController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Controller for writing logs to all writers"
'@Folder("VbaStackTraceLogger.Core.Writer")
'@ModuleDescription("Controller for writing logs to all writers")
Option Explicit

Implements Logger_IWriteLogController

Private Const CLASS_NAME As String = "Logger_WriteLogController"

Private mWriters As Collection
Private mLogMetadataProvider As Logger_ILogMetadataProvider

Private Sub Class_Initialize()
    Set mWriters = New Collection
    Set mLogMetadataProvider = New Logger_LogMetadataProvider
End Sub

'@Description("Initialize with collection of writers")
Public Sub Initialize(ByVal writers As Collection)
    Set mWriters = writers
    
    ' Initialize all writers
    If mWriters.Count > 0 Then
        Dim writer As Logger_ILogWriter
        For Each writer In mWriters
            writer.Initialize
        Next
    End If
End Sub

'@Description("Write log message to all writers")
Public Sub WriteLogToAll(ByVal writeMessage As String, ByVal logTag As LoggerLogTag, ByVal callPath As String)
    If mWriters Is Nothing Then Exit Sub
    If mWriters.Count = 0 Then Exit Sub
    
    ' Create DTO with metadata
    Dim logRequest As Logger_LogWriteRequestDTO
    Set logRequest = New Logger_LogWriteRequestDTO
    
    With logRequest
        .Message = writeMessage
        .TagType = logTag
        .TagName = mLogMetadataProvider.GetTagName(logTag)
        .Timestamp = mLogMetadataProvider.GetTimeStamp
        .IsStackTraceEnabled = (Len(callPath) > 0)
        .CallPath = callPath
    End With
    
    ' Send to all writers
    Dim writer As Logger_ILogWriter
    For Each writer In mWriters
        writer.WriteLog logRequest
    Next
End Sub

'@Description("Terminate all writers")
Public Sub Terminate()
    If mWriters Is Nothing Then Exit Sub
    If mWriters.Count = 0 Then Exit Sub
    
    Dim writer As Logger_ILogWriter
    For Each writer In mWriters
        writer.Terminate
    Next
End Sub

' === Interface Implementation ===

' Logger_IWriteLogController implementation
Private Sub Logger_IWriteLogController_Initialize(ByVal writers As Collection)
    Initialize writers
End Sub

Private Sub Logger_IWriteLogController_WriteLogToAll(ByVal writeMessage As String, ByVal logTag As LoggerLogTag, ByVal callPath As String)
    WriteLogToAll writeMessage, logTag, callPath
End Sub

Private Sub Logger_IWriteLogController_Terminate()
    Terminate
End Sub

Private Sub Class_Terminate()
    ' Fail-safe
    Terminate
    Set mWriters = Nothing
    Set mLogMetadataProvider = Nothing
End Sub
