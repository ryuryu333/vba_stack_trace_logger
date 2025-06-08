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
' ユーザーインターフェースをここに集約している
' ユーザーは Logger_Provider 経由で Logger_Facade にアクセスする
' （実際の流れは Logger_Provider -> Logger_SingletonManager -> Logger_Facade）
' myLogger.Log の様に利用する
Option Explicit
Private Const CLASS_NAME As String = "Logger_Facade"

Private mController As Logger_Controller
Private mSharedCallStackManager As Logger_CallStackManager ' Facade クラスにてインスタン化・解放、Controller と共有
Private mProcedureTracerProvider As Logger_ProcedureTracerProvider
Private mIsInitialized As Boolean

Private Sub Class_Initialize()
    Set mController = New Logger_Controller
    Set mSharedCallStackManager = New Logger_CallStackManager
    Set mProcedureTracerProvider = New Logger_ProcedureTracerProvider
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
Public Sub CompleteInitialization(ByRef config As LoggerConfigStruct)
    If mIsInitialized Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".Initialize", _
            "ロガーはすでに初期化済みです、2回目の初期化は実行できません"
        Exit Sub
    End If
    ' ユーザーが作成した設定値を用いて、各クラスを初期化
    mController.Initialize config, mSharedCallStackManager
    mProcedureTracerProvider.Initialize config.IsIncludeCallerInfo
    
    mIsInitialized = True
End Sub

' === ログ出力を指示 ===
Public Sub Log(ByVal writeMessage As String, Optional ByVal userSelectedTag As LoggerLogTag = LogTag_Info)
    EnsureInitialized
    
    mController.WriteLog writeMessage, userSelectedTag
End Sub

' === ロガーの終了処理を指示 ===
Public Sub Terminate()
    EnsureInitialized
    
    ' Controller は SharedStackTraceManager を参照しているため
    ' Controller の終了処理が終わった後に、SharedStackTraceManager を解放する
    mController.Terminate
    Set mProcedureTracerProvider = Nothing
    Set mSharedCallStackManager = Nothing
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
    
    ' Provider に Tracer の生成を依頼
    Set UsingTracer = mProcedureTracerProvider.CreateTracer(currentModuleName, currentProcName, mSharedCallStackManager, Me)
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
    Set mSharedCallStackManager = Nothing
    Set mProcedureTracerProvider = Nothing
End Sub


