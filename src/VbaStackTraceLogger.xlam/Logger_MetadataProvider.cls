VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_MetadataProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Logger_MetadataProvider.cls
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
#Else
    Private Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
#End If

Private Type logCallerInfo
    ModuleName As String
    ProcedureName As String
End Type

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

' メタ情報を付与
Public Sub EnrichLogEntry(ByRef logInfo As LoggerLogInfoStruct)
    ' TimeStamp
    logInfo.Timestamp = GetTimeStamp()
    ' TagName
    logInfo.TagName = GetTagName(logInfo.tagType)
    ' ModuleName、ProcedureName
    Dim callerInfo As logCallerInfo
    callerInfo = GetCallerInfo(logInfo.isIncludeCallerInfo, logInfo.CallPath, logInfo.ModuleName, logInfo.ProcedureName)
    logInfo.ModuleName = callerInfo.ModuleName
    logInfo.ProcedureName = callerInfo.ProcedureName
End Sub

' 現在の時刻を導出
Private Function GetTimeStamp() As String
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
Private Function GetTagName(ByVal tagEnum As LoggerLogTag) As String
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

' 呼び出し元情報を導出
Private Function GetCallerInfo(ByVal isIncludeCallerInfoSetting As Boolean, _
                               ByVal currentCallPath As String, _
                               ByVal currentModuleName As String, _
                               ByVal currentProcedureName As String) As logCallerInfo
    ' 結果保存用の変数
    Dim callerInfo As logCallerInfo
    
    ' （early return）呼び出し元情報を含めない、または、CallPath が空の場合
    ' ModuleName と ProcedureName を確実に空にする
    If Not isIncludeCallerInfoSetting Or Len(currentCallPath) = 0 Then
        callerInfo.ModuleName = ""
        callerInfo.ProcedureName = ""
        GetCallerInfo = callerInfo
        Exit Function
    End If
    
    ' （early return）ModuleName と ProcedureName が既に何らかの形で設定されている場合
    ' Facade からログ出力指示を受けた時点では両方とも空である想定なので、これは例外に備えた処理
    ' 値を変更せずに処理を終える
    If Len(currentModuleName) <> 0 Or Len(currentProcedureName) <> 0 Then
        callerInfo.ModuleName = currentModuleName
        callerInfo.ProcedureName = currentProcedureName
        GetCallerInfo = callerInfo
        Exit Function
    End If
    
    ' currentCallPath から呼び出し元情報（ModuleName、ProcedureName）を導出する
    callerInfo = DeriveModuleAndProcedureNamesFromCallPath(currentCallPath, currentModuleName)
    GetCallerInfo = callerInfo
End Function

' currentCallPath から ModuleName と ProcedureName を導出
Private Function DeriveModuleAndProcedureNamesFromCallPath(ByVal currentCallPath As String, _
                                                           ByVal currentModuleName As String) As logCallerInfo
    Dim callerInfo As logCallerInfo     ' 結果保存用の変数
    Dim firstSegment As String          ' currentCallPath の最初のセグメント (例: "Module1.ProcA")
    Dim dotPos As Long                  ' "." の位置
    Dim bracketPos As Long              ' " < " の位置 (コールスタック区切り)

    ' currentCallPath から最初のプロシージャ呼び出し部分 (例: "Module1.ProcA < Module2.ProcB" から "Module1.ProcA") を抽出
    bracketPos = InStr(currentCallPath, " < ")
    If bracketPos > 0 Then
        firstSegment = Trim(Left(currentCallPath, bracketPos - 1))
    Else
        firstSegment = Trim(currentCallPath) ' currentCallPath が単一の場合
    End If

    ' firstSegment が空なら、何も解析できないのでリターン
    If Len(firstSegment) = 0 Then
        callerInfo.ModuleName = ""
        callerInfo.ProcedureName = ""
        DeriveModuleAndProcedureNamesFromCallPath = callerInfo
        Exit Function
    End If

    ' 抽出したセグメントから ModuleName と ProcedureName を分離
    dotPos = InStrRev(firstSegment, ".") ' 末尾から "." を検索 (例: "ModuleName.ProcedureName")
    If dotPos > 0 Then
        ' "." が見つかれば、それより前を ModuleName、後を ProcedureName とする
        callerInfo.ModuleName = Left(firstSegment, dotPos - 1)
        callerInfo.ProcedureName = Mid(firstSegment, dotPos + 1)
    Else
        ' "." が見つからない場合
        ' firstSegment 全体を ProcedureName として扱う
        ' ModuleName はこのセグメントからは導出できないため、元の値を保持する (通常は空のまま)
        callerInfo.ModuleName = currentModuleName
        callerInfo.ProcedureName = firstSegment
    End If
    
    DeriveModuleAndProcedureNamesFromCallPath = callerInfo
End Function
