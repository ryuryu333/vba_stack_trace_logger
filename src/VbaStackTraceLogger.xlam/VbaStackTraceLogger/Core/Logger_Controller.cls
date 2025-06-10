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
'@Folder("VbaStackTraceLogger.Core")

Option Explicit
Private Const CLASS_NAME As String = "Logger_Controller"

Private mCallStackManager As Logger_CallStackManager
Private mProcedureTracerProvider As Logger_ProcedureTracerProvider
Private mMetadataHelper As Logger_MetadataHelper
Private mWriterManager As Logger_WriterManager
Private mLoggerConfig As Logger_ConfigStruct
Private mIsInitialized As Boolean

Private Sub Class_Initialize()
    Set mCallStackManager = New Logger_CallStackManager
    Set mProcedureTracerProvider = New Logger_ProcedureTracerProvider
    Set mMetadataHelper = New Logger_MetadataHelper
    Set mWriterManager = New Logger_WriterManager
End Sub

' === 初期化処理 ===
Public Sub Initialize(ByRef userConfig As Logger_ConfigStruct)
    ' === 各引数を検証後、メンバ変数へ代入 ===
    ValidateConfig userConfig
    Set mLoggerConfig = userConfig
    ' === ProcedureTracerProvider を初期化 ===
    mProcedureTracerProvider.Initialize userConfig.IsStackTraceEnabled
    
    ' === Writers を作成・初期化 ===
    mWriterManager.CreateWriters userConfig
    mWriterManager.InitializeWriters
    
    mIsInitialized = True
End Sub

' 動的検証を実行、静的検証は Logger_ConfigBuilder にて実施
Private Sub ValidateConfig(ByRef userConfig As Logger_ConfigStruct)
    ' Excel出力設定の検証
    If userConfig.IsWriteToExcelSheet Then
        ' エクセルシートにアクセスし、参照可能な状態かを確認
        On Error Resume Next
        Dim testValue As String
        testValue = userConfig.OutputExcelSheet.Name
        If Err.Number <> 0 Then
            On Error GoTo 0
            Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
                LOGGER_NAMESPACE & "." & CLASS_NAME & ".ValidateConfig", _
                "指定されたエクセルシートにアクセスできません"
        End If
        On Error GoTo 0
    End If
End Sub

' === ログ出力 ===
' logInfo に情報を書き込んだ後、Writer に渡してログを出力させる
Public Sub WriteLog(ByVal writeMessage As String, ByVal userSelectedTag As LoggerLogTag)
    EnsureInitialized
    
    ' === ロガーの設定を元に、ログ出力すべきかを判定 ===
    If ShouldWriteLog(userSelectedTag) = False Then Exit Sub
    
    ' === logInfo に情報を書き込む ===
    Dim logInfo As Logger_LogInfoStruct
    Set logInfo = New Logger_LogInfoStruct
    With logInfo
        .Message = writeMessage
        .TagType = userSelectedTag
        .TagName = mMetadataHelper.GetTagName(.TagType)
        .Timestamp = mMetadataHelper.GetTimeStamp
        .IsStackTraceEnabled = mLoggerConfig.IsStackTraceEnabled
        If .IsStackTraceEnabled Then .CallPath = mCallStackManager.GetCurrentCallPath
    End With
    ' === 各 Writer にログ出力を命令 ===
    mWriterManager.WriteLogToAll logInfo
End Sub

' ロガーの設定を元に、ログ出力すべきかを判定
' False となるパターンを順番に評価、どれにも当てはまらなければTrue
Private Function ShouldWriteLog(ByVal selectedTag As LoggerLogTag) As Boolean
    ShouldWriteLog = True
    
    ' ログ出力が無効の場合
    If mLoggerConfig.IsLoggingEnabled = False Then
        ShouldWriteLog = False
        Exit Function
    End If
    
    ' 特定のタグのログ出力が無効の場合
    If mLoggerConfig.IsTagFilteringEnabled Then
        Dim userDefinedExcludedTags() As LoggerLogTag
        userDefinedExcludedTags = mLoggerConfig.ExcludedTags
        ' ユーザー設定された除外ログと一致するか確認
        Dim i As Integer
        For i = LBound(userDefinedExcludedTags) To UBound(userDefinedExcludedTags)
            If selectedTag = userDefinedExcludedTags(i) Then
                ShouldWriteLog = False
                Exit Function
            End If
        Next i
    End If
End Function

' === スタックトレース ===
Public Function UsingTracer(ByVal currentModuleName As String, _
                            ByVal currentProcName As String) As Logger_ProcedureTracer
    EnsureInitialized
    
    ' Provider に Tracer の生成を依頼
    Set UsingTracer = mProcedureTracerProvider.CreateTracer(currentModuleName, currentProcName, mCallStackManager, Me)
End Function


' === 終了処理 ===
Public Sub Terminate()
    EnsureInitialized

    mWriterManager.TerminateWriters
    
    Set mMetadataHelper = Nothing
    Set mProcedureTracerProvider = Nothing
    Set mCallStackManager = Nothing
    Set mWriterManager = Nothing
End Sub
End Sub

' 初期化済みか確認
Private Sub EnsureInitialized()
    If Not mIsInitialized Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".EnsureInitialized", _
            "Controller が初期化されていません"
    End If
End Sub

Private Sub Class_Terminate()
    ' フェイルセーフ
    If Not mWriterManager Is Nothing Then
        mWriterManager.TerminateWriters
    End If
    Set mCallStackManager = Nothing
    Set mProcedureTracerProvider = Nothing
    Set mMetadataHelper = Nothing
    Set mWriterManager = Nothing
End Sub


