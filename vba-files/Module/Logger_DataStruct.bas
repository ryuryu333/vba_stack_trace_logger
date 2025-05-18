Attribute VB_Name = "Logger_DataStruct"
' Logger_DataStruct.bas

Public Type LoggerConfigStruct
    IsIncludeCallerInfo As Boolean
    IsWriteToImmediate As Boolean
    IsWriteToExcelSheet As Boolean
    OutputExcelSheet As Worksheet
End Type

Public Type LoggerLogInfoStruct
    Message As String
    TagType As LoggerLogTag
    TagName As String
    Timestamp As String
    IsIncludeCallerInfo As Boolean
    ModuleName As String
    ProcedureName  As String
End Type

Public Enum LoggerLogTag
    LogTag_Debug = 0
    LogTag_Info = 1
    LogTag_Warning = 2
    LogTag_Error = 3
    LogTag_Critical = 4
End Enum
