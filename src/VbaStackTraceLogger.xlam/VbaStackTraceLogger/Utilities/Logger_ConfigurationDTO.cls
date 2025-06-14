VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_ConfigurationDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Data Transfer Object for Logger Configuration"
'@Folder("VbaStackTraceLogger.Utilities")
'@ModuleDescription("Data Transfer Object for Logger Configuration")
Option Explicit

' Private fields
Private pIsLoggingEnabled As Boolean
Private pIsTagFilteringEnabled As Boolean
Private pExcludedTags() As LoggerLogTag
Private pIsStackTraceEnabled As Boolean
Private pIsWriteToImmediate As Boolean
Private pIsWriteToExcelSheet As Boolean
Private pOutputExcelSheet As Worksheet

Public Property Get IsLoggingEnabled() As Boolean
    IsLoggingEnabled = pIsLoggingEnabled
End Property
Public Property Let IsLoggingEnabled(ByVal value As Boolean)
    pIsLoggingEnabled = value
End Property

Public Property Get IsTagFilteringEnabled() As Boolean
    IsTagFilteringEnabled = pIsTagFilteringEnabled
End Property
Public Property Let IsTagFilteringEnabled(ByVal value As Boolean)
    pIsTagFilteringEnabled = value
End Property

Public Property Get ExcludedTags() As LoggerLogTag()
    ExcludedTags = pExcludedTags
End Property
Public Property Let ExcludedTags(ByRef value() As LoggerLogTag)
    pExcludedTags = value
End Property

Public Property Get IsStackTraceEnabled() As Boolean
    IsStackTraceEnabled = pIsStackTraceEnabled
End Property
Public Property Let IsStackTraceEnabled(ByVal value As Boolean)
    pIsStackTraceEnabled = value
End Property

Public Property Get IsWriteToImmediate() As Boolean
    IsWriteToImmediate = pIsWriteToImmediate
End Property
Public Property Let IsWriteToImmediate(ByVal value As Boolean)
    pIsWriteToImmediate = value
End Property

Public Property Get IsWriteToExcelSheet() As Boolean
    IsWriteToExcelSheet = pIsWriteToExcelSheet
End Property
Public Property Let IsWriteToExcelSheet(ByVal value As Boolean)
    pIsWriteToExcelSheet = value
End Property

Public Property Get OutputExcelSheet() As Worksheet
    Set OutputExcelSheet = pOutputExcelSheet
End Property
Public Property Set OutputExcelSheet(ByVal value As Worksheet)
    Set pOutputExcelSheet = value
End Property

' Validate configuration settings
Public Sub Validate()
    If IsTagFilteringEnabled Then
        ' ExcludedTags must be set by the builder.
        If ExcludedTags Is Nothing Then
            Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
                LOGGER_NAMESPACE & "." & CLASS_NAME & ".Validate", _
                "Unexpected error. ExcludedTags is Nothing."
        End If
        On Error Resume Next
        Dim boundCheck As Long
        boundCheck = LBound(ExcludedTags) ' If not array, this will raise an error
        If Err.Number <> 0 Then
            On Error GoTo 0
            Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
                LOGGER_NAMESPACE & "." & CLASS_NAME & ".Validate", _
                "Unexpected error. ExcludedTags have non-initialized array."
        End If
        On Error GoTo 0
        ' ExcludedTags must be set to an array with more than one tag by the builder.
        Dim tagCount As Long
        tagCount = UBound(ExcludedTags) - LBound(ExcludedTags) + 1
        If tagCount <= 0 Then
            Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
                LOGGER_NAMESPACE & "." & CLASS_NAME & ".Validate", _
                "Unexpected error. ExcludedTags have empty array."
        End If
        ' Check for duplicates in ExcludedTags
        Dim tagDict As Object
        Set tagDict = CreateObject("Scripting.Dictionary")
        Dim tag As LoggerLogTag
        Dim isDuplicate As Boolean
        isDuplicate = False
        For Each tag In ExcludedTags
            If tagDict.Exists(tag) Then
                isDuplicate = True
                Exit For
            End If
            tagDict.Add tag, True
        Next
        If isDuplicate Then
            Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
                LOGGER_NAMESPACE & "." & CLASS_NAME & ".Validate", _
                "ExcludedTags have duplicate tags." _&
                " Please ensure that each tag is unique."
        End If
    End If

    If IsWriteToExcelSheet Then
        ' OutputExcelSheet must be set by the builder.
        If OutputExcelSheet Is Nothing Then
            Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
                LOGGER_NAMESPACE & "." & CLASS_NAME & ".Validate", _
                "Unexpected error. OutputExcelSheet is Nothing."
        End If
        ' Access Excel sheet and verify it's accessible
        On Error Resume Next
        Dim testValue As String
        testValue = OutputExcelSheet.Name
        If Err.Number <> 0 Then
            On Error GoTo 0
            Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
                LOGGER_NAMESPACE & "." & CLASS_NAME & ".Validate", _
                "Cannot access the user-defined Excel sheet." & _
                "Please ensure that the OutputExcelSheet is accessible."
        End If
        On Error GoTo 0
    End If    
End Sub

