Attribute VB_Name = "Logger_DataStruct"
' Logger_DataStruct.bas

Public Type LoggerConfigStruct
    IsLoggingEnabled As Boolean
    IsTagFilteringEnabled  As Boolean
    ExcludedTags() As LoggerLogTag
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
    CallPath As String
End Type

' 無効値 を表す要素として LogTag_None = -1 も定義すべきだが
' インテリセンスにユーザーが使用しないものを表示したくなにので未定義
' タグを追加するときは -1 以外を指定する
Public Enum LoggerLogTag
    LogTag_Debug = 0
    LogTag_Info = 1
    LogTag_Warning = 2
    LogTag_Error = 3
    LogTag_Critical = 4
    LogTag_Trace = 5
End Enum

