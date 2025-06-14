VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_LogMetadataProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Logger_LogMetadataProvider.cls
'@Folder("VbaStackTraceLogger.Core")

Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
#Else
    Private Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
#End If

Private Type SYSTEMTIME
    wYear        As Integer
    wMonth       As Integer
    wDayOfWeek   As Integer
    wDay         As Integer
    wHour        As Integer
    wMinute      As Integer
    wSecond      As Integer
    wMilliseconds As Integer
End Type

' 現在の時刻を導出
Public Function GetTimeStamp() As String
    Dim st As SYSTEMTIME
    GetLocalTime st

    GetTimeStamp = _
        Format$(st.wYear, "0000") & "-" & _
        Format$(st.wMonth, "00") & "-" & _
        Format$(st.wDay, "00") & " " & _
        Format$(st.wHour, "00") & ":" & _
        Format$(st.wMinute, "00") & ":" & _
        Format$(st.wSecond, "00") & "." & _
        Format$(st.wMilliseconds, "000")
End Function

' ログタグの Enum 値から、ログ表示用の文字列を導出
Public Function GetTagName(ByVal tagEnum As LoggerLogTag) As String
    Select Case tagEnum
        Case LogTag_Debug:    GetTagName = "DEBUG"
        Case LogTag_Info:     GetTagName = "INFO"
        Case LogTag_Warning:  GetTagName = "WARNING"
        Case LogTag_Error:    GetTagName = "ERROR"
        Case LogTag_Critical: GetTagName = "CRITICAL"
        Case LogTag_Trace:    GetTagName = "Trace"
        Case Else:            GetTagName = "INFO"
    End Select
End Function
