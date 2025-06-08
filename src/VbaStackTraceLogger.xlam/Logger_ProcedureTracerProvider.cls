VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_ProcedureTracerProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Logger_ProcedureTracerProvider.cls
Option Explicit
Private Const CLASS_NAME As String = "Logger_ProcedureTracerProvider"

Private mIsStackTraceEnabled As Boolean
Private mIsInitialized As Boolean

Public Sub Initialize(ByVal IsStackTraceEnabled As Boolean)
    mIsStackTraceEnabled = IsStackTraceEnabled
    mIsInitialized = True
End Sub

Public Function CreateTracer(ByVal currentModuleName As String, ByVal currentProcedureName As String, _
                             ByRef stackTraceManager As Logger_CallStackManager, _
                             ByRef loggerFacade As Logger_Facade) As Logger_ProcedureTracer
    EnsureInitialized
    
    ' スタックトレース機能が有効な場合のみ、Logger_ProcedureTracer を返す
    ' ユーザー視点では、関数の冒頭で Logger.UsingTracer 経由で Variant 型の変数を受け取る
    ' Logger_ProcedureTracer を直接操作しない
    ' スタックトレース機能が無効な場合は Nothing を返す（機能が複雑化したら interface を検討）
    Dim tracer As Logger_ProcedureTracer
    If mIsStackTraceEnabled Then
        Set tracer = New Logger_ProcedureTracer
        tracer.Initialize currentModuleName, currentProcedureName, stackTraceManager, loggerFacade
    Else
        Set tracer = Nothing
    End If
    
    Set CreateTracer = tracer
End Function

Private Sub EnsureInitialized()
    If Not mIsInitialized Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".EnsureInitialized", _
            "ProcedureTracerProvider が初期化されていません"
    End If
End Sub


