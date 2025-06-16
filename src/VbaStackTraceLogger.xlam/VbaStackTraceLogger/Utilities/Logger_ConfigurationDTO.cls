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

Option Explicit

Private Const CLASS_NAME As String = "Logger_ConfigurationDTO"

' This class is used to store configuration settings for the logger.
' Builder class set prperties, call Validate() and MarkAsFinalized().
' After that, Builder class send the instance to Controller class.
' NOTE : VBA does not support immutable classes, so this class is mutable.
Private mIsLoggingEnabled As Boolean
Private mIsTagFilteringEnabled As Boolean
Private mExcludedTags() As LoggerLogTag
Private mIsStackTraceEnabled As Boolean
Private mIsWriteToImmediate As Boolean
Private mIsWriteToExcelSheet As Boolean
Private mOutputExcelSheet As Worksheet

Private mIsFinalized As Boolean

Friend Property Get IsLoggingEnabled() As Boolean
    IsLoggingEnabled = mIsLoggingEnabled
End Property
Friend Property Let IsLoggingEnabled(ByVal value As Boolean)
    VerifySetAllowed
    mIsLoggingEnabled = value
End Property

Friend Property Get IsTagFilteringEnabled() As Boolean
    IsTagFilteringEnabled = mIsTagFilteringEnabled
End Property
Friend Property Let IsTagFilteringEnabled(ByVal value As Boolean)
    VerifySetAllowed
    mIsTagFilteringEnabled = value
End Property

Friend Property Get ExcludedTags() As LoggerLogTag()
    ExcludedTags = mExcludedTags
End Property
Friend Property Let ExcludedTags(ByRef value() As LoggerLogTag)
    VerifySetAllowed
    mExcludedTags = value
End Property

Friend Property Get IsStackTraceEnabled() As Boolean
    IsStackTraceEnabled = mIsStackTraceEnabled
End Property
Friend Property Let IsStackTraceEnabled(ByVal value As Boolean)
    VerifySetAllowed
    mIsStackTraceEnabled = value
End Property

Friend Property Get IsWriteToImmediate() As Boolean
    IsWriteToImmediate = mIsWriteToImmediate
End Property
Friend Property Let IsWriteToImmediate(ByVal value As Boolean)
    VerifySetAllowed
    mIsWriteToImmediate = value
End Property

Friend Property Get IsWriteToExcelSheet() As Boolean
    IsWriteToExcelSheet = mIsWriteToExcelSheet
End Property
Friend Property Let IsWriteToExcelSheet(ByVal value As Boolean)
    VerifySetAllowed
    mIsWriteToExcelSheet = value
End Property

Friend Property Get OutputExcelSheet() As Worksheet
    Set OutputExcelSheet = mOutputExcelSheet
End Property
Friend Property Set OutputExcelSheet(ByVal value As Worksheet)
    VerifySetAllowed
    Set mOutputExcelSheet = value
End Property

Private Sub VerifySetAllowed()
    If mIsFinalized Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".EnsureFinalized", _
            "Unexpected error. After build configuration, cannot set any properties."
    End If
End Sub

Friend Sub MarkAsFinalized()
    mIsFinalized = True
End Sub

Friend Sub Validate()
    If IsTagFilteringEnabled Then
        ' ExcludedTags must be set by the builder.
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
        Dim tag As Variant
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
                "ExcludedTags have duplicate tags." _
                & "Please ensure that each tag is unique."
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

