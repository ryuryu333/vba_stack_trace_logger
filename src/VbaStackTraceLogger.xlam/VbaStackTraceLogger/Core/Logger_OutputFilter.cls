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
'@Folder("VbaStackTraceLogger.Core")

Option Explicit
Private Const CLASS_NAME As String = "Logger_OutputFilter"

' === Log Output Filtering ===
' Determine whether log should be written based on configuration
Public Function ShouldWriteLog(ByVal selectedTag As LoggerLogTag, ByRef config As Logger_ConfigStruct) As Boolean
    ShouldWriteLog = True
    
    ' If logging is disabled
    If config.IsLoggingEnabled = False Then
        ShouldWriteLog = False
        Exit Function
    End If
    
    ' If specific tag logging is disabled
    If config.IsTagFilteringEnabled Then
        Dim userDefinedExcludedTags() As LoggerLogTag
        userDefinedExcludedTags = config.ExcludedTags
        ' Check if it matches user-configured excluded logs
        Dim i As Integer
        For i = LBound(userDefinedExcludedTags) To UBound(userDefinedExcludedTags)
            If selectedTag = userDefinedExcludedTags(i) Then
                ShouldWriteLog = False
                Exit Function
            End If
        Next i
    End If
End Function
