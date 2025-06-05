VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_Controller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Logger_Controller.cls
Option Explicit

Private mMetadataProvider As Logger_MetadataProvider
Private mWriters As Collection
Private mLoggerConfig As LoggerConfigStruct
Private mCurrentCallPath As String

' === 初期化処理 ===
Public Sub Initialize(ByRef userConfig As LoggerConfigStruct)
    ' === メンバ変数を初期化 ===
    Set mMetadataProvider = New Logger_MetadataProvider
    Set mWriters = New Collection

    ' === ユーザーが入力したロガーの設定値が正常か判定、メンバ変数に代入 ===
    ValidateConfig userConfig
    mLoggerConfig = userConfig
    
    '  === Writers の作成 ===
    ' どこにログを出力するかを確認し、Logger_ILogWriter の具象クラスのインスタンスを作成、mWriters に追加
    ' イミディエイトウインドウ（VBE）
    If mLoggerConfig.isWriteToImmediate Then
        Dim immediateWriter As Logger_Writer_Immediate: Set immediateWriter = New Logger_Writer_Immediate
        AddWriter immediateWriter
    End If
    ' エクセルシート（どのエクセルシートに書き込むかも設定）
    If mLoggerConfig.isWriteToExcelSheet Then
        Dim excelWriter As Logger_Writer_Excel: Set excelWriter = New Logger_Writer_Excel
        excelWriter.SetTargetSheet mLoggerConfig.outputExcelSheet
        AddWriter excelWriter
    End If
    
    ' === Logger_ILogWriter の具象クラスの初期化を実施 ===
    InitializeWriters
End Sub

' ユーザーが入力したロガーの設定値が正常か判定
Private Function ValidateConfig(ByRef userConfig As LoggerConfigStruct)
    ' 検証結果を収納する
    Dim isValid As Boolean: isValid = True
    Dim errorMessage As String: errorMessage = ""
    
    ' タグフィルタリング設定の検証
    If userConfig.isTagFilteringEnabled Then
        ' excludedTags の検証
        ' Function CreateDisabledTags を利用して excludedTags 値が代入されている前提
        ' -> 重複のない、0 個以上の要素を持つ、LoggerLogTag 配列、となっている
        
        ' 配列数を確認
        Dim tagCount As Long
        tagCount = UBound(userConfig.excludedTags) - LBound(userConfig.excludedTags) + 1
        If tagCount <= 0 Then
            isValid = False
            errorMessage = errorMessage & "・excludedTags にタグが指定されていません"
        End If
    End If
    
    ' Excel出力設定の検証
    If userConfig.isWriteToExcelSheet Then
        ' エクセルシートが指定されているか確認
        If userConfig.outputExcelSheet Is Nothing Then
            isValid = False
            errorMessage = errorMessage & "・Excel出力が有効ですが、出力先シート(outputExcelSheet)が設定されていません"
        End If
        ' エクセルシートにアクセスし、参照可能な状態かを確認
        On Error Resume Next
        Dim testValue As Variant
        testValue = userConfig.outputExcelSheet.Name
        If Not Err.Number = 0 Then
            isValid = False
            errorMessage = errorMessage & "・指定されたエクセルシートにアクセスできません"
        End If
        On Error GoTo 0
    End If
    
    ' 検証結果
    If Not isValid Then Err.Raise vbObjectError + 513, "Logger_Controller.ValidateConfig", errorMessage
End Function

Sub a()
    
End Sub

' === ログ出力 ===
' logInfo に情報を書き込んだ後、Writer に渡してログを出力させる
Public Sub WriteLog(logInfo As LoggerLogInfoStruct)
    ' === ロガーの設定を元に、ログ出力すべきかを判定 ===
    If ShouldWriteLog(logInfo.tagType) = False Then Exit Sub
    
    ' === logInfo に情報を書き込む ===
    With logInfo
        .isIncludeCallerInfo = mLoggerConfig.isIncludeCallerInfo
        .CallPath = IIf(.isIncludeCallerInfo, mCurrentCallPath, "")
    End With
    
    ' === logInfo のメタ情報を付与 ===
    ' （TimeStamp、TagName、CallPath、ModuleName、ProcedureName）
    mMetadataProvider.EnrichLogEntry logInfo
    
    ' === 各 Writer にログ出力を命令 ===
    Dim writer As Logger_ILogWriter
    For Each writer In mWriters
        writer.WriteLog logInfo
    Next
End Sub

' ロガーの設定を元に、ログ出力すべきかを判定
' Falseとなるパターンを順番に評価、どれにも当てはまらなければTrue
Private Function ShouldWriteLog(tagType As LoggerLogTag) As Boolean
    ' ログ出力が有効かを確認
    If mLoggerConfig.isLoggingEnabled Then
        ShouldWriteLog = True
        Exit Function
    End If
    
    ' スタックトレース機能が無効の場合
    If mLoggerConfig.isIncludeCallerInfo = False And tagType = LogTag_Trace Then
        ShouldWriteLog = False
        Exit Function
    End If
    
    ' 特定のタグのログ出力が無効の場合
    If mLoggerConfig.isTagFilteringEnabled Then
        Dim i As Integer
        For i = LBound(mLoggerConfig.excludedTags) To UBound(mLoggerConfig.excludedTags)
            If mLoggerConfig.excludedTags(i) = tagType Then
                ShouldWriteLog = False
                Exit Function
            End If
        Next i
    End If
End Function

' === スタックトレースの呼び出し階層情報を更新 ===
' 新しい関数が呼び出された際の処理
Public Function EnterProcedureContext(ByVal currentModuleName As String, ByVal currentProcName As String) As String
    Dim newPathSegment As String
    newPathSegment = currentModuleName & "." & currentProcName

    ' 新しい関数を呼ばれた時点でのスタックトレース情報を返した後、情報を更新する
    EnterProcedureContext = mCurrentCallPath
    If Len(mCurrentCallPath) > 0 Then
        mCurrentCallPath = newPathSegment & " < " & mCurrentCallPath
    Else
        mCurrentCallPath = newPathSegment
    End If
End Function

' 関数が終了した際の処理
Public Sub ExitProcedureContext(ByVal previousPath As String)
    mCurrentCallPath = previousPath
End Sub

' === 終了処理 ===
Public Sub Terminate()
    TerminateWriters
End Sub

' === Writer 関連の処理 ===
' Writer を追加登録
Private Sub AddWriter(writer As Logger_ILogWriter)
    mWriters.Add writer
End Sub

' 各 Writer を初期化
Private Sub InitializeWriters()
    Dim writer As Logger_ILogWriter
    For Each writer In mWriters
        writer.Initialize
    Next
End Sub

' 各Writerの終了処理
Private Sub TerminateWriters()
    Dim writer As Logger_ILogWriter
    For Each writer In mWriters
        writer.Terminate
    Next
End Sub



