VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_Controller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Logger_Controller.cls
Option Explicit

Private pMetadataProvider As Logger_MetadataProvider
Private pWriters As Collection
Private logInfo As LoggerLogInfoStruct
Private pIsIncludeCallerInfo As Boolean
Private pModuleName As String
Private pProcedureName As String
Private pCurrentCallPath As String

Public Sub SetModuleName(ModuleName As String)
    pModuleName = ModuleName
End Sub

Public Sub SetProcedureName(ProcedureName As String)
    pProcedureName = ProcedureName
End Sub

Public Sub Initialize(LogConfig As LoggerConfigStruct)
    Set pMetadataProvider = New Logger_MetadataProvider
    Set pWriters = New Collection

    If LogConfig.IsWriteToImmediate Then
        Dim immediateWriter As Logger_Writer_Immediate
        Set immediateWriter = New Logger_Writer_Immediate
        AddWriter immediateWriter
    End If

    If LogConfig.IsWriteToExcelSheet Then
        Dim excelWriter As Logger_Writer_Excel
        Set excelWriter = New Logger_Writer_Excel
        If LogConfig.OutputExcelSheet Is Nothing Then
            MsgBox "Excel出力が有効になっていますが、出力先のシートが設定されていません。" & vbCrLf & _
                   "ログは正しく出力されません。", vbExclamation, "ログ設定の警告"
        Else
            excelWriter.SetTargetSheet LogConfig.OutputExcelSheet
        End If
        AddWriter excelWriter
    End If
    
    pIsIncludeCallerInfo = LogConfig.IsIncludeCallerInfo
    
    InitializeWriters
End Sub

' Writerを追加登録
Public Sub AddWriter(writer As Logger_ILogWriter)
    pWriters.Add writer
End Sub

' 各Writerを初期化
Public Sub InitializeWriters()
    Dim writer As Logger_ILogWriter
    For Each writer In pWriters
        writer.Initialize
    Next
End Sub

' ログを出力
Public Sub WriteLog(logInfo As LoggerLogInfoStruct)
    ' ログ情報の前処理
    logInfo.TagName = GetTagName(logInfo.TagType)
    logInfo.ModuleName = pModuleName
    logInfo.ProcedureName = pProcedureName
    pMetadataProvider.AttachMetadata logInfo
    logInfo.IsIncludeCallerInfo = pIsIncludeCallerInfo
    If pIsIncludeCallerInfo Then
        logInfo.CallPath = pCurrentCallPath
    Else
        logInfo.CallPath = ""
    End If
    
    ' ログ出力命令
    Dim writer As Logger_ILogWriter
    For Each writer In pWriters
        writer.WriteLog logInfo
    Next
End Sub

' 新しい関数が呼び出された際の処理
Public Function EnterProcedureContext(ByVal ModuleName As String, ByVal ProcName As String) As String
    Dim newPathSegment As String
    newPathSegment = ModuleName & "." & ProcName

    ' 新しい関数を呼ばれた時点でのスタックトレース情報を返した後、情報を更新する
    EnterProcedureContext = pCurrentCallPath

    If Len(pCurrentCallPath) > 0 Then
        pCurrentCallPath = newPathSegment & " < " & pCurrentCallPath
    Else
        pCurrentCallPath = newPathSegment
    End If
End Function

Public Sub ExitProcedureContext(ByVal previousPath As String)
    pCurrentCallPath = previousPath
End Sub

' 各Writerの終了処理
Public Sub TerminateWriters()
    Dim writer As Logger_ILogWriter
    For Each writer In pWriters
        writer.Terminate
    Next
End Sub

Private Function GetTagName(tagEnum As LoggerLogTag) As String
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


