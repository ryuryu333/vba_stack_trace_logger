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

Private mController As Logger_Controller
Private mIsInitialized As Boolean ' VBA の仕様で bool 初期値は False になるので、宣言だけして初期化は未記載

' コンストラクタ
Private Sub Class_Initialize()
    Set mController = New Logger_Controller
End Sub

' === ロガーの初期化処理を指示 ===
Public Function Initialize() As Logger_ConfigBuilder
    If mIsInitialized Then
        Debug.Print "ロガーはすでに初期化されています、2回目の初期化は実行できません"
        Exit Function
    End If
    
    ' ビルダーで初期化に用いる設定値を作成
    ' その後、Controller の初期化処理が呼び出される
    Dim configBuilder As New Logger_ConfigBuilder
    configBuilder.SetLoggerController mController
    Set Initialize = configBuilder
    mIsInitialized = True
End Function

' ログ出力を指示
Public Sub Log(Message As String, Optional tag As LoggerLogTag = LogTag_Info)
    EnsureInitialized

    Dim logInfo As LoggerLogInfoStruct
    logInfo.Message = Message
    logInfo.tagType = tag

    mController.WriteLog logInfo
End Sub

' ロガーの終了処理を指示
Public Sub Terminate()
    EnsureInitialized
    
    mController.Terminate
    Set mController = Nothing
    mIsInitialized = False
End Sub

' === スタックトレース関連 ===
' スタックトレース機能を利用する
' Logger_ProcedureTracer のライフサイクルを利用しているので、インスタンスを返す
Public Function UsingTracer(ByVal currentModuleName As String, _
                            ByVal currentProcName As String) As Logger_ProcedureTracer
    EnsureInitialized
    
    Dim tracer As New Logger_ProcedureTracer
    tracer.Init currentModuleName, currentProcName
    Set UsingTracer = tracer
End Function

' ユーザーが直接呼出し不可
' Logger_ProcedureTracer 生成時に呼び出し
Friend Function EnterProcedureContext(ByVal currentModuleName As String, ByVal currentProcName As String) As String
    EnsureInitialized
    
    EnterProcedureContext = mController.EnterProcedureContext(currentModuleName, currentProcName)
End Function

' ユーザーが直接呼出し不可
' Logger_ProcedureTracer 破棄時に呼び出し
Friend Sub ExitProcedureContext(ByVal previousPath As String)
    EnsureInitialized
    
    mController.ExitProcedureContext previousPath
End Sub

' === utility 関数 ===
' Logger_SingletonManager の Logger_Facade インスタンスを破棄する
Public Sub Reset()
    ReleaseMyLogger
End Sub

' === モジュール全体で利用される Private 関数 ===
' 初期化済みか確認
Private Sub EnsureInitialized()
    If Not mIsInitialized Then
        Err.Raise vbObjectError + 513, "Logger_Facade.Log", "ロガーが初期化されていません。先に Initialize を呼び出してください。"
        Exit Sub
    End If
End Sub
