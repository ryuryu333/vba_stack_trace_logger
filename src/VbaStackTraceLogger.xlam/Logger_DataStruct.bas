Attribute VB_Name = "Logger_DataStruct"
' Logger_DataStruct.bas

Public Type LoggerConfigStruct
    isLoggingEnabled As Boolean
    isTagFilteringEnabled  As Boolean
    excludedTags() As LoggerLogTag
    isIncludeCallerInfo As Boolean
    isWriteToImmediate As Boolean
    isWriteToExcelSheet As Boolean
    outputExcelSheet As Worksheet
End Type

Public Type LoggerLogInfoStruct
    Message As String
    tagType As LoggerLogTag
    TagName As String
    Timestamp As String
    isIncludeCallerInfo As Boolean
    ModuleName As String
    ProcedureName  As String
    CallPath As String
End Type

' �����l ��\���v�f�Ƃ��� LogTag_None = -1 ����`���ׂ�����
' �C���e���Z���X�Ƀ��[�U�[���g�p���Ȃ����̂�\���������Ȃɂ̂Ŗ���`
' �^�O��ǉ�����Ƃ��� -1 �ȊO���w�肷��
Public Enum LoggerLogTag
    LogTag_Debug = 0
    LogTag_Info = 1
    LogTag_Warning = 2
    LogTag_Error = 3
    LogTag_Critical = 4
    LogTag_Trace = 5
End Enum

