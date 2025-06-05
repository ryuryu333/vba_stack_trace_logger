VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_Writer_Excel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Logger_ExcelWriter.cls
Option Explicit
Implements Logger_ILogWriter

Private mOutputSheet As Worksheet
Private mCurrentRow As Long

' 出力先シートを設定
Public Sub SetTargetSheet(sheet As Worksheet)
    Set mOutputSheet = sheet
End Sub

Private Sub Logger_ILogWriter_Initialize()
    If mOutputSheet Is Nothing Then
        MsgBox "Logger_ExcelWriter: 出力先のシートが設定されていません。", vbExclamation
        Exit Sub
    End If

    ' 過去のログ内容と背景色をクリア (A列からD列)
    With mOutputSheet.Columns("A:D")
        .ClearContents ' 内容を消去
        .Interior.ColorIndex = xlNone ' 背景色をクリア
    End With

    With mOutputSheet
        .Cells(1, 1).value = "Timestamp"
        .Cells(1, 2).value = "Tag"
        .Cells(1, 3).value = "呼び出し元情報" ' Changed from Module
        .Cells(1, 4).value = "Message"       ' Changed from Procedure (and effectively was Message)
    End With
    
    With mOutputSheet.Range("A1:D1") ' Adjusted range to D1
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
    End With

    mCurrentRow = 2  ' ログ出力開始行

    Debug.Print "エクセル用のロガーを初期化しました"
End Sub

Private Sub Logger_ILogWriter_WriteLog(logInfo As LoggerLogInfoStruct)
    If mOutputSheet Is Nothing Then Exit Sub

    With mOutputSheet
        .Cells(mCurrentRow, 1).value = logInfo.Timestamp
        .Cells(mCurrentRow, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss.000"
        .Cells(mCurrentRow, 2).value = logInfo.TagName

        If logInfo.isIncludeCallerInfo Then
            If Len(logInfo.CallPath) > 0 Then
                .Cells(mCurrentRow, 3).value = logInfo.CallPath
            ElseIf Len(logInfo.ModuleName) > 0 And Len(logInfo.ProcedureName) > 0 Then
                .Cells(mCurrentRow, 3).value = logInfo.ModuleName & "." & logInfo.ProcedureName
            Else
                .Cells(mCurrentRow, 3).value = "" ' No specific caller info
            End If
        Else
            .Cells(mCurrentRow, 3).value = "" ' Caller info not requested
        End If

        .Cells(mCurrentRow, 4).value = logInfo.Message ' Message now in column D
        
        .Range(.Cells(mCurrentRow, 1), .Cells(mCurrentRow, 4)).VerticalAlignment = xlCenter ' Adjusted range to column D
        
        ' 最初の出力（2行目）のときだけ列幅を AutoFit
        If mCurrentRow = 2 Then
            .Columns("A:D").AutoFit ' Adjusted range to column D
        End If
        
        ' 背景色の適用
        Dim bgColor As Long
        Select Case logInfo.tagType
            Case LogTag_Warning
                bgColor = RGB(255, 255, 153)  ' 薄い黄色
            Case LogTag_Error, LogTag_Critical
                bgColor = RGB(255, 199, 206)  ' 薄い赤色
            Case Else
                bgColor = xlNone  ' デフォルト（塗らない）
        End Select
    
        If bgColor <> xlNone Then
            .Range(.Cells(mCurrentRow, 1), .Cells(mCurrentRow, 4)).Interior.Color = bgColor ' Adjusted range to column D
        End If
            
    End With
    
    mCurrentRow = mCurrentRow + 1
End Sub

Private Sub Logger_ILogWriter_Terminate()
    If Not mOutputSheet Is Nothing Then
        On Error Resume Next ' Handle cases where sheet might not be available
        mOutputSheet.Columns("A:D").AutoFit ' Adjusted range to column D
        On Error GoTo 0
    End If
End Sub

