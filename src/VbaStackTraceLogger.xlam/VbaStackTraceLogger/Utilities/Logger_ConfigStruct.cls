VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_ConfigStruct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Logger_ConfigStruct.cls
'@Folder("VbaStackTraceLogger.Utilities")

Option Explicit

Private pIsLoggingEnabled As Boolean
Private pIsTagFilteringEnabled As Boolean
Private pExcludedTags() As LoggerLogTag
Private pIsStackTraceEnabled As Boolean
Private pIsWriteToImmediate As Boolean
Private pIsWriteToExcelSheet As Boolean
Private pOutputExcelSheet As Worksheet

Friend Property Get IsLoggingEnabled() As Boolean
    IsLoggingEnabled = pIsLoggingEnabled
End Property
Friend Property Let IsLoggingEnabled(value As Boolean)
    pIsLoggingEnabled = value
End Property

Friend Property Get IsTagFilteringEnabled() As Boolean
    IsTagFilteringEnabled = pIsTagFilteringEnabled
End Property
Friend Property Let IsTagFilteringEnabled(value As Boolean)
    pIsTagFilteringEnabled = value
End Property

Friend Property Get ExcludedTags() As LoggerLogTag()
    ExcludedTags = pExcludedTags
End Property
Friend Property Let ExcludedTags(value() As LoggerLogTag)
    pExcludedTags = value
End Property

Friend Property Get IsStackTraceEnabled() As Boolean
    IsStackTraceEnabled = pIsStackTraceEnabled
End Property
Friend Property Let IsStackTraceEnabled(value As Boolean)
    pIsStackTraceEnabled = value
End Property

Friend Property Get IsWriteToImmediate() As Boolean
    IsWriteToImmediate = pIsWriteToImmediate
End Property
Friend Property Let IsWriteToImmediate(value As Boolean)
    pIsWriteToImmediate = value
End Property

Friend Property Get IsWriteToExcelSheet() As Boolean
    IsWriteToExcelSheet = pIsWriteToExcelSheet
End Property
Friend Property Let IsWriteToExcelSheet(value As Boolean)
    pIsWriteToExcelSheet = value
End Property

Friend Property Get OutputExcelSheet() As Worksheet
    Set OutputExcelSheet = pOutputExcelSheet
End Property
Friend Property Set OutputExcelSheet(value As Worksheet)
    Set pOutputExcelSheet = value
End Property
