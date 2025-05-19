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
Private mIsIncludeCallerInfo As Boolean
Private mCurrentCallPath As String

' 初期化処理
Public Sub Initialize(userLogSetting As LoggerConfigStruct)
    ' メンバ変数を初期化
    Set mMetadataProvider = New Logger_MetadataProvider
    Set mWriters = New Collection

    ' ユーザー指定の設定を反映
    
    ' どこにログを出力するかを確認し、Logger_ILogWriter の具象クラスのインスタンスを作成、mWriters に追加
    ' イミディエイトウインドウ（VBE）
    If userLogSetting.IsWriteToImmediate Then
        Dim immediateWriter As Logger_Writer_Immediate
        Set immediateWriter = New Logger_Writer_Immediate
        AddWriter immediateWriter
    End If
    ' エクセルシート（どのエクセルシートに書き込むかも確認）
    If userLogSetting.IsWriteToExcelSheet Then
        Dim excelWriter As Logger_Writer_Excel
        Set excelWriter = New Logger_Writer_Excel
        If userLogSetting.OutputExcelSheet Is Nothing Then
            MsgBox "Excel出力が有効になっていますが、出力先のシートが設定されていません。" & vbCrLf & _
                   "ログは正しく出力されません。", vbExclamation, "ログ設定の警告"
        Else
            excelWriter.SetTargetSheet userLogSetting.OutputExcelSheet
        End If
        AddWriter excelWriter
    End If
    
    ' コールスタック情報を含めるかを設定
    mIsIncludeCallerInfo = userLogSetting.IsIncludeCallerInfo
    
    ' Logger_ILogWriter の具象クラスの初期化を実施
    InitializeWriters
End Sub

' Writer を追加登録
Public Sub AddWriter(writer As Logger_ILogWriter)
    mWriters.Add writer
End Sub

' 各 Writer を初期化
Public Sub InitializeWriters()
    Dim writer As Logger_ILogWriter
    For Each writer In mWriters
        writer.Initialize
    Next
End Sub

' ログ出力
' logInfo に情報を書き込んだ後、writer に渡してログを出力させる
Public Sub WriteLog(logInfo As LoggerLogInfoStruct)
    ' 呼び出し元情報を含めるかどうかのフラグを logInfo に設定
    ' このフラグは LoggerConfigStruct で初期化時に設定されたもの
    logInfo.IsIncludeCallerInfo = mIsIncludeCallerInfo
    
    ' IsIncludeCallerInfo フラグに基づいて CallPath を設定
    ' mCurrentCallPath（コールスタック情報）は別の関数で処理している
    logInfo.CallPath = IIf(mIsIncludeCallerInfo, mCurrentCallPath, "")
    
    ' logInfo のメタ情報を付与
    ' （TimeStamp、TagName、CallPath、ModuleName、ProcedureName）
    mMetadataProvider.EnrichLogEntry logInfo
    
    ' 各 writer にログ出力を命令
    Dim writer As Logger_ILogWriter
    For Each writer In mWriters
        writer.WriteLog logInfo
    Next
End Sub

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

Public Sub ExitProcedureContext(ByVal previousPath As String)
    mCurrentCallPath = previousPath
End Sub

' 各Writerの終了処理
Public Sub TerminateWriters()
    Dim writer As Logger_ILogWriter
    For Each writer In mWriters
        writer.Terminate
    Next
End Sub


