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

Private mStackTraceController As Logger_StackTraceController
Private mLogInfoBuilder As Logger_LogInfoBuilder
Private mWriterManager As Logger_WriterManager
Private mOutputFilter As Logger_OutputFilter
Private mLoggerConfig As Logger_ConfigStruct
Private mIsInitialized As Boolean

Private Sub Class_Initialize()
    Set mStackTraceController = New Logger_StackTraceController
    Set mLogInfoBuilder = New Logger_LogInfoBuilder
    Set mWriterManager = New Logger_WriterManager
    Set mOutputFilter = New Logger_OutputFilter
End Sub

' === 初期化処理 ===
Public Sub Initialize(ByRef userConfig As Logger_ConfigStruct)
    ' === 各引数を検証後、メンバ変数へ代入 ===
    ValidateConfig userConfig
    Set mLoggerConfig = userConfig
    ' === StackTraceController を初期化 ===
    mStackTraceController.Initialize userConfig.IsStackTraceEnabled, Me
    
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
    If mOutputFilter.ShouldWriteLog(userSelectedTag, mLoggerConfig) = False Then Exit Sub
    
    ' === logInfo に情報を書き込む ===
    Dim logInfo As Logger_LogInfoStruct
    Set logInfo = mLogInfoBuilder.BuildLogInfo(writeMessage, userSelectedTag, mLoggerConfig, mStackTraceController.GetCallStackManager)
    
    ' === 各 Writer にログ出力を命令 ===
    mWriterManager.WriteLogToAll logInfo
End Sub

' === スタックトレース ===
Public Function UsingTracer(ByVal currentModuleName As String, _
                            ByVal currentProcName As String) As Logger_ProcedureTracer
    EnsureInitialized
    
    ' StackTraceController に Tracer の生成を依頼
    Set UsingTracer = mStackTraceController.UsingTracer(currentModuleName, currentProcName)
End Function


' === 終了処理 ===
Public Sub Terminate()
    EnsureInitialized
    
    mWriterManager.TerminateWriters
    mStackTraceController.Terminate
    Set mLogInfoBuilder = Nothing
    Set mStackTraceController = Nothing
    Set mWriterManager = Nothing
    Set mOutputFilter = Nothing
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
    Set mStackTraceController = Nothing
    Set mLogInfoBuilder = Nothing
    Set mWriterManager = Nothing
    Set mOutputFilter = Nothing
End Sub


