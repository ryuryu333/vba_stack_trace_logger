VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_LogWriteRequestDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Data Transfer Object for Log Write Request"
'@Folder("VbaStackTraceLogger.Utilities")
'@ModuleDescription("Data Transfer Object for Log Write Request")
Option Explicit

' Private fields
Private mMessage As String
Private mTagType As LoggerLogTag
Private mTagName As String
Private mTimestamp As String
Private mCallPath As String

' As an alternative to a parameterized constructor
Friend Sub Initialize(ByVal Message As String, _
                    ByVal TagType As LoggerLogTag, _
                    ByVal TagName As String, _
                    ByVal Timestamp As String, _
                    ByVal CallPath As String)
    mMessage = Message
    mTagType = TagType
    mTagName = TagName
    mTimestamp = Timestamp
    mCallPath = CallPath
End Sub

' Property: Message
Public Property Get Message() As String
    Message = mMessage
End Property

Public Property Let Message(ByVal value As String)
    mMessage = value
End Property

' Property: TagType
Public Property Get TagType() As LoggerLogTag
    TagType = mTagType
End Property

Public Property Let TagType(ByVal value As LoggerLogTag)
    mTagType = value
End Property

' Property: TagName
Public Property Get TagName() As String
    TagName = mTagName
End Property

Public Property Let TagName(ByVal value As String)
    mTagName = value
End Property

' Property: Timestamp
Public Property Get Timestamp() As String
    Timestamp = mTimestamp
End Property

Public Property Let Timestamp(ByVal value As String)
    mTimestamp = value
End Property

' Property: CallPath
Public Property Get CallPath() As String
    CallPath = mCallPath
End Property

Public Property Let CallPath(ByVal value As String)
    mCallPath = value
End Property
