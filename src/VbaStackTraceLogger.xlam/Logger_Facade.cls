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

' === ロガーの初期化処理を指示 ===
Public Sub Initialize(Optional isLoggingEnabled As Boolean = True, _
                               Optional isTagFilteringEnabled As Boolean = True, _
                               Optional excludedTags As Logger_TagArrayWrapper = Nothing, _
                               Optional isIncludeCallerInfo As Boolean = True, _
                               Optional isWriteToImmediate As Boolean = True, _
                               Optional isWriteToExcelSheet As Boolean = False, _
                               Optional outputExcelSheet As Worksheet = Nothing)
    
    If mIsInitialized Then
        Debug.Print "ロガーはすでに初期化されています、2回目の初期化は実行できません"
        Exit Sub
    End If
    
    Dim config As LoggerConfigStruct
    With config
        .isLoggingEnabled = isLoggingEnabled
        .isTagFilteringEnabled = isTagFilteringEnabled
        .isIncludeCallerInfo = isIncludeCallerInfo
        .isWriteToImmediate = isWriteToImmediate
        .isWriteToExcelSheet = isWriteToExcelSheet
        If excludedTags Is Nothing Then
            .excludedTags = InitialValueOfExcludedTags
        Else
            .excludedTags = excludedTags.GetArray
        End If
        Set .outputExcelSheet = outputExcelSheet
    End With
    
    Debug.Print "ロガーの初期化を開始します"
    Set mController = New Logger_Controller
    mController.Initialize config
    mIsInitialized = True
    Debug.Print "ロガーの初期化が完了しました"
End Sub

' ExcludedTags 初期値を設定
' ExcludedTags は LoggerLogTag の配列であり、配列を Optional な引数にはできない
' ラッパークラスを引数に指定することで、LoggerLogTag の配列のみを代入するよう強制する仕組みにしている
Private Function InitialValueOfExcludedTags() As LoggerLogTag()
    InitialValueOfExcludedTags = CreateExcludedTagArray.ToArray.GetArray
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

Public Function CreateExcludedTagArray() As Logger_TagArrayBuilder
    Set CreateExcludedTagArray = New Logger_TagArrayBuilder
End Function

' LoggerConfigStruct の DisabledTags() As LoggerLogTag に代入するための配列を作成する
' インテリセンスにより、LoggerLogTag の具体的な名前を知らなくても、引数を記入できる
' ParamArray はインテリセンスの補助が得られないので使用しない
Function CreateDisabledTags( _
    Optional tag1 As LoggerLogTag = -1, _
    Optional tag2 As LoggerLogTag = -1, _
    Optional tag3 As LoggerLogTag = -1, _
    Optional tag4 As LoggerLogTag = -1, _
    Optional tag5 As LoggerLogTag = -1, _
    Optional tag6 As LoggerLogTag = -1 _
) As LoggerLogTag()

    ' 引数の数を拡張する際はここを変更
    Dim inputTagSlotCount As Long: inputTagSlotCount = 6
    Dim inputTagsArray() As LoggerLogTag: inputTagsArray = Array(tag1, tag2, tag3, tag4, tag5, tag6)

    ' 有効なタグの数をカウント
    Dim validTagCount As Integer
    Dim i As Integer
    validTagCount = 0
    For i = 0 To inputTagSlotCount
        If inputTagsArray(i) <> -1 Then ' -1 = Enum で未定義の値 = 無効
            validTagCount = validTagCount + 1
        End If
    Next i
    
    ' 有効なタグがない場合は空配列を返す
    If validTagCount = 0 Then
        Dim emptyArray() As LoggerLogTag
        CreateDisabledTags = emptyArray
        Exit Function
    End If
    
    ' 有効なタグのみを結果配列にコピー
    Dim result() As LoggerLogTag
    ReDim result(validTagCount - 1)
    
    Dim resultIndex As Integer
    resultIndex = 0
    
    For i = 0 To 5
        If inputTagsArray(i) <> -1 Then
            result(resultIndex) = inputTagsArray(i)
            resultIndex = resultIndex + 1
        End If
    Next i
    
    CreateDisabledTags = result
End Function

' === モジュール全体で利用される Private 関数 ===
' 初期化済みか確認
Private Sub EnsureInitialized()
    If Not mIsInitialized Then
        Err.Raise vbObjectError + 513, "Logger_Facade.Log", "ロガーが初期化されていません。先に Initialize を呼び出してください。"
        Exit Sub
    End If
End Sub
