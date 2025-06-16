VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_ILogMetadataProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for Log Metadata Provider"
'@Folder("VbaStackTraceLogger.Core.Writer")
'@Interface
'@ModuleDescription("Interface for Log Metadata Provider")
Option Explicit

'@Description("Get current timestamp string")
Public Function GetTimeStamp() As String
End Function

'@Description("Get tag name string for specified log tag")
Public Function GetTagName(ByVal logTag As LoggerLogTag) As String
End Function
