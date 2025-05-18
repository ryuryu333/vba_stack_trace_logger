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
Option Explicit

Private pController As Logger_Controller
Private pIsInitialized As Boolean ' VBA の仕様で bool 初期値は Flase になるので、宣言だけして初期化は記載していません

Public Sub Initialize(LogConfig As LoggerConfigStruct)
    If pIsInitialized Then
        Debug.Print "ロガーはすでに初期化されています、2回目の初期化は実行できません"
        Exit Sub
    End If
    Debug.Print "ロガーの初期化を開始します"
    Set pController = New Logger_Controller
    pController.Initialize LogConfig
    pIsInitialized = True
    Debug.Print "ロガーの初期化が完了しました"
End Sub

Public Sub Log(Message As String, Optional Tag As LoggerLogTag = LogTag_Info)
    If Not pIsInitialized Then
        Err.Raise vbObjectError, "Logger_Facade.Initialize が実行されていません、先に初期化してください"
    End If

    Dim logInfo As LoggerLogInfoStruct
    logInfo.Message = Message
    logInfo.TagType = Tag

    pController.WriteLog logInfo
End Sub

Public Function EnterProcedureContext(ByVal ModuleName As String, ByVal ProcName As String) As String
    If Not pIsInitialized Or pController Is Nothing Then
        Err.Raise vbObjectError + 513, "Logger_Facade.EnterProcedureContext", "ロガーが初期化されていません。"
        Exit Function ' Should not happen if initialized
    End If
    EnterProcedureContext = pController.EnterProcedureContext(ModuleName, ProcName)
End Function

Public Sub ExitProcedureContext(ByVal previousPath As String)
    If Not pIsInitialized Or pController Is Nothing Then
        Err.Raise vbObjectError + 513, "Logger_Facade.ExitProcedureContext", "ロガーが初期化されていません。"
        Exit Sub ' Should not happen if initialized
    End If
    pController.ExitProcedureContext previousPath
End Sub

Public Sub SetCallerInfo(ModuleName As String, ProcedureName As String)
    If Not pIsInitialized Then
        Err.Raise vbObjectError, "Logger_Facade.Initialize が実行されていません、先に初期化してください"
    End If

    pController.SetModuleName ModuleName
    pController.SetProcedureName ProcedureName
End Sub

Public Sub Terminate()
    If Not pController Is Nothing Then
        pController.TerminateWriters
    End If
End Sub

