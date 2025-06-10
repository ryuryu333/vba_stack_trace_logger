VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_StackTraceController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Logger_StackTraceController.cls
'@Folder("VbaStackTraceLogger.Core")

Option Explicit
Private Const CLASS_NAME As String = "Logger_StackTraceController"

Private mCallStackManager As Logger_CallStackManager
Private mProcedureTracerProvider As Logger_ProcedureTracerProvider
Private mMainController As Logger_Controller
Private mIsInitialized As Boolean

Private Sub Class_Initialize()
    Set mCallStackManager = New Logger_CallStackManager
    Set mProcedureTracerProvider = New Logger_ProcedureTracerProvider
End Sub

' === 初期化処理 ===
Public Sub Initialize(ByVal IsStackTraceEnabled As Boolean, ByRef mainController As Logger_Controller)
    If mainController Is Nothing Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".Initialize", _
            "メインコントローラーのインスタンスが見つかりません"
    End If
    
    Set mMainController = mainController
    
    ' === ProcedureTracerProvider を初期化 ===
    mProcedureTracerProvider.Initialize IsStackTraceEnabled
    
    mIsInitialized = True
End Sub

' === スタックトレース ===
Public Function UsingTracer(ByVal currentModuleName As String, _
                            ByVal currentProcName As String) As Logger_ProcedureTracer
    EnsureInitialized
      ' Provider に Tracer の生成を依頼（現在は循環参照があるがそのまま残す）
    Set UsingTracer = mProcedureTracerProvider.CreateTracer(currentModuleName, currentProcName, mCallStackManager, mMainController)
End Function

' === CallStackManager への参照を取得 ===
Public Function GetCallStackManager() As Logger_CallStackManager
    EnsureInitialized
    Set GetCallStackManager = mCallStackManager
End Function

' === 終了処理 ===
Public Sub Terminate()
    EnsureInitialized
    
    Set mProcedureTracerProvider = Nothing
    Set mCallStackManager = Nothing
    Set mMainController = Nothing
    mIsInitialized = False
End Sub

' 初期化済みか確認
Private Sub EnsureInitialized()
    If Not mIsInitialized Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".EnsureInitialized", _
            "StackTraceController が初期化されていません"
    End If
End Sub

Private Sub Class_Terminate()
    ' フェイルセーフ
    Set mCallStackManager = Nothing
    Set mProcedureTracerProvider = Nothing
    Set mMainController = Nothing
End Sub
