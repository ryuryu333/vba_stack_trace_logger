VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_OutputFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Logger_OutputFilter.cls
'@Folder("VbaStackTraceLogger.Core.Engine")

Option Explicit

Implements Logger_IOutputFilter
Private Const CLASS_NAME As String = "Logger_OutputFilter"

' === Log Output Filtering ===
' Determine whether log should be written based on configuration
Public Function ShouldWriteLog(ByVal selectedTag As LoggerLogTag, _
                              ByVal configIsLoggingEnabled As Boolean, _
                              ByVal configIsTagFilteringEnabled As Boolean, _
                              ByRef configExcludedTags() As LoggerLogTag) As Boolean
    ShouldWriteLog = True
    
    ' If logging is disabled
    If configIsLoggingEnabled = False Then
        ShouldWriteLog = False
        Exit Function
    End If
    
    ' If specific tag logging is disabled
    If configIsTagFilteringEnabled Then
        ' Check if it matches user-configured excluded logs
        Dim i As Integer
        For i = LBound(configExcludedTags) To UBound(configExcludedTags)
            If selectedTag = configExcludedTags(i) Then
                ShouldWriteLog = False
                Exit Function
            End If
        Next i
    End If
End Function

' === Interface Implementation ===

' Logger_IOutputFilter implementation
Private Function Logger_IOutputFilter_ShouldWriteLog(ByVal selectedTag As LoggerLogTag, _
                                                     ByVal configIsLoggingEnabled As Boolean, _
                                                     ByVal configIsTagFilteringEnabled As Boolean, _
                                                     ByRef configExcludedTags() As LoggerLogTag) As Boolean
    ' Delegate to main method
    Logger_IOutputFilter_ShouldWriteLog = ShouldWriteLog(selectedTag, configIsLoggingEnabled, configIsTagFilteringEnabled, configExcludedTags)
End Function

