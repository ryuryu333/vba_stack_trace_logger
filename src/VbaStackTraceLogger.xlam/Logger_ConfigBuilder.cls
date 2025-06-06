VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_ConfigBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Logger_ConfigBuilder.cls
Option Explicit
Private Const CLASS_NAME As String = "ConfigBuilder"

Private mConfig As LoggerConfigStruct
Private mLoggerController As Logger_Controller
Private mIsBuilt As Boolean

' ロガーの設定の初期値を指定
' 必要最低限な機能のみを有効化する設定にしている
Private Sub Class_Initialize()
    Dim defaultConfig As LoggerConfigStruct
    With defaultConfig
        .IsLoggingEnabled = True
        .IsTagFilteringEnabled = False
        ' Enum の配列である ExcludedTags は defaultConfig 宣言時点で空配列が代入されている
        .IsIncludeCallerInfo = False
        .IsWriteToImmediate = True
        .IsWriteToExcelSheet = False
        Set .OutputExcelSheet = Nothing
    End With
    mConfig = defaultConfig
End Sub

Friend Sub SetLoggerController(ByRef logger As Logger_Controller)
    Set mLoggerController = logger
End Sub

Public Function DisableLogging() As Logger_ConfigBuilder
    ' IsLoggingEnabled の初期値は True
    mConfig.IsLoggingEnabled = False
    Set DisableLogging = Me
End Function

Public Function EnableTagFiltering() As Logger_TagFilterBuilder
    ' IsLoggingEnabled の初期値は False
    mConfig.IsTagFilteringEnabled = True
    ' サブビルダーへ連鎖させる
    Dim nextStep As New Logger_TagFilterBuilder
    nextStep.SetParentBuilder Me
    Set EnableTagFiltering = nextStep
End Function

' サブビルダー Logger_TagArrayBuilder からここへ連鎖する
Friend Sub SetExcludedTags(ByRef tags() As LoggerLogTag)
    Set mConfig.ExcludedTags = tags
End Sub

Public Function EnableStackTrace() As Logger_ConfigBuilder
    ' IsIncludeCallerInfo の初期値は False
    mConfig.IsIncludeCallerInfo = True
    Set EnableStackTrace = Me
End Function

Public Function DisableWriteToImmediate() As Logger_ConfigBuilder
    ' IsWriteToImmediate の初期値は True
    mConfig.IsWriteToImmediate = False
    Set DisableWriteToImmediate = Me
End Function

Public Function EnableWriteToExcelSheet() As Logger_ExcelOutputBuilder
    ' IsWriteToExcelSheet の初期値は False
    mConfig.IsWriteToExcelSheet = True
    ' サブビルダーへ連鎖させる
    Dim nextStep As New Logger_ExcelOutputBuilder
    nextStep.SetParentBuilder Me
    Set EnableWriteToExcelSheet = nextStep
End Function

' サブビルダー Logger_ExcelOutputBuilder からここへ連鎖する
Friend Sub SetOutputExcelSheet(ByVal sheet As Worksheet)
    Set mConfig.OutputExcelSheet = sheet
End Sub

' ロガーの設定用の値を確定、初期化処理を呼び出す
Public Sub Build()
    Validate
    If Not mLoggerController Is Nothing Then
        mLoggerController.Initialize mConfig
    End If
    mIsBuilt = True
End Sub

' 静的検証を実行、実行環境における動作検証は Controller にて実施
Private Sub Validate()
    If mConfig.IsTagFilteringEnabled Then
        Dim tagCount As Long
        tagCount = UBound(mConfig.ExcludedTags) - LBound(mConfig.ExcludedTags) + 1
        If tagCount <= 0 Then
            Err.Raise vbObjectError + 513, "ConfigBuilder", _
                "タグフィルタリングが有効ですが、除外タグが指定されていません"
        End If
    End If
    
    If mConfig.IsWriteToExcelSheet Then
        If mConfig.OutputExcelSheet Is Nothing Then
            Err.Raise vbObjectError + 513, "ConfigBuilder", _
                "Excel出力が有効ですが、出力先シートが設定されていません"
        End If
    End If
End Sub

Private Sub Class_Terminate()
    If Not mIsBuilt Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
                  LOGGER_NAMESPACE & "." & CLASS_NAME & ".Class_Terminate", _
                  "Build() が呼ばれていません、ロガーの設定が適用されていないため正常に動作しません"
    End If
End Sub
