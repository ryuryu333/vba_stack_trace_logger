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

' === ロガーの基本的な処理 ===
' ロガーの初期化処理を指示
Public Sub Initialize(Optional includeCallerInfo As Boolean = False, _
                               Optional writeToImmediate As Boolean = True, _
                               Optional writeToExcel As Boolean = False, _
                               Optional targetSheet As Worksheet = Nothing)
    Dim config As LoggerConfigStruct
    With config
        .IsIncludeCallerInfo = includeCallerInfo
        .IsWriteToImmediate = writeToImmediate
        .IsWriteToExcelSheet = writeToExcel
        Set .OutputExcelSheet = targetSheet
    End With
    
    If mIsInitialized Then
        Debug.Print "ロガーはすでに初期化されています、2回目の初期化は実行できません"
        Exit Sub
    End If
    
    Debug.Print "ロガーの初期化を開始します"
    Set mController = New Logger_Controller
    mController.Initialize config
    mIsInitialized = True
    Debug.Print "ロガーの初期化が完了しました"
End Sub

' ログ出力を指示
Public Sub Log(message As String, Optional tag As LoggerLogTag = LogTag_Info)
    EnsureInitialized

    Dim logInfo As LoggerLogInfoStruct
    logInfo.message = message
    logInfo.TagType = tag

    mController.WriteLog logInfo
End Sub

' ロガーの終了処理を指示
Public Sub Terminate()
    EnsureInitialized
    
    If Not mController Is Nothing Then
        mController.TerminateWriters
        Set mController = Nothing
    End If
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


' === Private 関数 ===
' 初期化済みか確認
Private Sub EnsureInitialized()
    If Not mIsInitialized Then
        Err.Raise vbObjectError + 513, "Logger_Facade.Log", "ロガーが初期化されていません。先に Initialize を呼び出してください。"
        Exit Sub
    End If
End Sub

' デストラクタ
Private Sub Class_Terminate()
    ' クラスが破棄されるときの後処理
    Debug.Print "Logger_Facade が破棄されました"
End Sub
