VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_LogInfoMetadataHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Logger_LogInfoMetadataHelper.cls
'@Folder("VbaStackTraceLogger.Core")

Option Explicit
Private Const CLASS_NAME As String = "Logger_LogInfoMetadataHelper"

Private mMetadataProvider As Logger_LogMetadataProvider

Private Sub Class_Initialize()
    Set mMetadataProvider = New Logger_LogMetadataProvider
End Sub

' Build complete log information structure with metadata
Public Function BuildLogInfo(ByVal userMessage As String, ByVal userSelectedTag As LoggerLogTag, ByRef userConfig As Logger_ConfigStruct, ByVal currentCallPath As String) As Logger_LogInfoStruct
    
    Dim logInfo As Logger_LogInfoStruct
    Set logInfo = New Logger_LogInfoStruct
      With logInfo
        .Message = userMessage
        .TagType = userSelectedTag
        .TagName = mMetadataProvider.GetTagName(.TagType)
        .Timestamp = mMetadataProvider.GetTimeStamp
        .IsStackTraceEnabled = userConfig.IsStackTraceEnabled
        If .IsStackTraceEnabled Then .CallPath = currentCallPath
    End With
    
    Set BuildLogInfo = logInfo
End Function

Private Sub Class_Terminate()
    ' Fail-safe
    Set mMetadataProvider = Nothing
End Sub
