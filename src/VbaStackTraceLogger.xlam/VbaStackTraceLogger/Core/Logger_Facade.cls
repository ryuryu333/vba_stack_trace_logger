VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_Facade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Logger_Facade.cls
'@Folder("VbaStackTraceLogger.Core")

' ユーザーインターフェースをここに集約している
' ユーザーは Logger_Provider 経由で Logger_Facade にアクセスする
' （実際の流れは Logger_Provider -> Logger_SingletonManager -> Logger_Facade）
' myLogger.Log の様に利用する

Option Explicit
Private Const CLASS_NAME As String = "Logger_Facade"

Private mController As Logger_Controller
Private mIsInitialized As Boolean

Private Sub Class_Initialize()
    Set mController = New Logger_Controller
End Sub

' === ロガーの初期化を指示 ===
' User > Facade.Initialize > ConfigBuilder.Build > Facade.CompleteInitialization
' の順番で処理が実施される
Public Function Initialize() As Logger_ConfigBuilder
    If mIsInitialized Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".Initialize", _
            "ロガーはすでに初期化済みです、2回目の初期化は実行できません"
        Exit Function
    End If

    ' ビルダーで初期化に用いる設定値を作成
    ' その後、CompleteInitialization に返ってくる
    Dim configBuilder As New Logger_ConfigBuilder
    configBuilder.Initialize Me
    Set Initialize = configBuilder
End Function

' User > Facade.Initialize > ConfigBuilder.Build > Facade.CompleteInitialization
' の順番で処理が実施される
Public Sub CompleteInitialization(ByRef userConfig As Logger_ConfigStruct)
    If mIsInitialized Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".Initialize", _
            "ロガーはすでに初期化済みです、2回目の初期化は実行できません"
        Exit Sub
    End If
    
    ' ユーザーが作成した設定値を用いて、Controller を初期化
    mController.Initialize userConfig
    
    mIsInitialized = True
End Sub

' === ログ出力を指示 ===
Public Sub Log(ByVal writeMessage As String, Optional ByVal userSelectedTag As LoggerLogTag = LoggerLogTag.LogTag_Info)
    EnsureInitialized

    mController.WriteLog writeMessage, userSelectedTag
End Sub

' === ロガーの終了処理を指示 ===
Public Sub Terminate()
    EnsureInitialized
    
    mController.Terminate
    Set mController = Nothing
    
    mIsInitialized = False
End Sub

' === スタックトレース関連 ===
' ProcedureTracerProvider > Create ProcedureTracer > User get ProcedureTracer instance
' ユーザーが ProcedureTracer をプロシージャの冒頭で変数として保持、プロシージャの終了と同時に変数が破棄される
' インスタンスのライフサイクルを利用して、プロシージャの開始・終了を検知し処理を実施
Public Function UsingTracer(ByVal currentModuleName As String, _
                            ByVal currentProcName As String) As Logger_ProcedureTracer
    EnsureInitialized
    
    Set UsingTracer = mController.UsingTracer(currentModuleName, currentProcName)
End Function

' === utility 関数 ===
' Logger_SingletonManager の Logger_Facade インスタンスを破棄する
Public Sub Reset()
    ReleaseMyLogger
End Sub

' === モジュール全体で利用される Private 関数 ===
' 初期化済みか確認
Private Sub EnsureInitialized()
    If Not mIsInitialized Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".EnsureInitialized", _
            "ロガーが初期化されていません"
    End If
End Sub

Private Sub Class_Terminate()
    ' フェイルセーフ
    Set mController = Nothing
End Sub


