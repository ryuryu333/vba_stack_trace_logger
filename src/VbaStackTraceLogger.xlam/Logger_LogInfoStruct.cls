VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_LogInfoStruct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Logger_LogInfoStruct.cls
'@Folder("VbaStackTraceLogger.Utilities")

Option Explicit

Private pMessage As String
Private pTagType As LoggerLogTag
Private pTagName As String
Private pTimestamp As String
Private pIsStackTraceEnabled As Boolean
Private pCallPath As String

Friend Property Get Message() As String
    Message = pMessage
End Property
Friend Property Let Message(value As String)
    pMessage = value
End Property

Friend Property Get TagType() As LoggerLogTag
    TagType = pTagType
End Property
Friend Property Let TagType(value As LoggerLogTag)
    pTagType = value
End Property

Friend Property Get TagName() As String
    TagName = pTagName
End Property
Friend Property Let TagName(value As String)
    pTagName = value
End Property

Friend Property Get Timestamp() As String
    Timestamp = pTimestamp
End Property
Friend Property Let Timestamp(value As String)
    pTimestamp = value
End Property

Friend Property Get IsStackTraceEnabled() As Boolean
    IsStackTraceEnabled = pIsStackTraceEnabled
End Property
Friend Property Let IsStackTraceEnabled(value As Boolean)
    pIsStackTraceEnabled = value
End Property

Friend Property Get CallPath() As String
    CallPath = pCallPath
End Property
Friend Property Let CallPath(value As String)
    pCallPath = value
End Property

