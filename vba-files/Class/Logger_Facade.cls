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

Private mController As Logger_Controller
Private msInitialized As Boolean ' VBA の仕様で bool 初期値は Flase になるので、宣言だけして初期化は記載していません

Public Sub Initialize(userLogSetting As LoggerConfigStruct)
    If msInitialized Then
        Debug.Print "ロガーはすでに初期化されています、2回目の初期化は実行できません"
        Exit Sub
    End If
    
    Debug.Print "ロガーの初期化を開始します"
    Set mController = New Logger_Controller
    mController.Initialize userLogSetting
    msInitialized = True
    Debug.Print "ロガーの初期化が完了しました"
End Sub

Public Sub Log(message As String, Optional tag As LoggerLogTag = LogTag_Info)
    If Not msInitialized Then
        Err.Raise vbObjectError + 513, "Logger_Facade.Log", "ロガーが初期化されていません。先に Initialize を呼び出してください。"
        Exit Sub
    End If

    Dim logInfo As LoggerLogInfoStruct
    logInfo.message = message
    logInfo.TagType = tag

    mController.WriteLog logInfo
End Sub

Friend Function EnterProcedureContext(ByVal currentModuleName As String, ByVal currentProcName As String) As String
    If Not msInitialized Then
        Err.Raise vbObjectError + 513, "Logger_Facade.EnterProcedureContext", "ロガーが初期化されていません。"
        Exit Function
    End If
    
    EnterProcedureContext = mController.EnterProcedureContext(currentModuleName, currentProcName)
End Function

Friend Sub ExitProcedureContext(ByVal previousPath As String)
    If Not msInitialized Then
        Err.Raise vbObjectError + 513, "Logger_Facade.ExitProcedureContext", "ロガーが初期化されていません。"
        Exit Sub
    End If
    mController.ExitProcedureContext previousPath
End Sub

Public Sub Terminate()
    If Not mController Is Nothing Then
        mController.TerminateWriters
        Set mController = Nothing
    End If
    msInitialized = False
End Sub

