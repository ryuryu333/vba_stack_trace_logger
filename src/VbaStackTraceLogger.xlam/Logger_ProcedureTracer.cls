VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_ProcedureTracer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Logger_ProcedureTracer.cls
'@Folder("VbaStackTraceLogger.StackTrace")

' RAII を模倣し、ライフサイクルと紐づけてスタックトレース情報の更新を実施
' 変数として使用するクラス
Option Explicit
Private Const CLASS_NAME As String = "Logger_ProcedureTracer"

Private Const ENTER_MESSAG As String = ">> Enter "
Private Const EXIT_MESSAG As String = "<< Exit "
Private mCallSite As String

Private mStackTraceManager As Logger_CallStackManager
Private mController As Logger_Controller
Private mIsInitialized As Boolean

' インスタンス化の後に初期化が指示される
' 引数付きコンストラクタの代わり
Public Sub Initialize(ByVal currentModuleName As String, ByVal currentProcedureName As String, _
                      ByRef stackTraceManager As Logger_CallStackManager, ByRef controller As Logger_Controller)
    If stackTraceManager Is Nothing Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".Initialize", _
            "StackTraceManager のインスタンスが見つかりません"
    End If
    If controller Is Nothing Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".Initialize", _
            "controller のインスタンスが見つかりません"
    End If
    
    Set mStackTraceManager = stackTraceManager
    Set mController = controller
    mCallSite = currentModuleName & "." & currentProcedureName
    
    ' コールスタック更新を指示
    mStackTraceManager.EnterProcedure (mCallSite)
    
    ' プロシージャ開始のログ出力を指示
    mController.WriteLog ENTER_MESSAG & mCallSite, LogTag_Trace
    
    mIsInitialized = True
End Sub

' デストラクタ
Private Sub Class_Terminate()
    If mIsInitialized = False Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".Class_Terminate", _
            "初期化が未実施の状態で解放処理が呼ばれました"
    End If
    
    ' プロシージャ終了のログ出力を指示
    mController.WriteLog EXIT_MESSAG & mCallSite, LogTag_Trace
    
    ' コールスタック更新を指示
    mStackTraceManager.ExitProcedure
    
    ' リソース解放
    Set mStackTraceManager = Nothing
    Set mController = Nothing
End Sub


