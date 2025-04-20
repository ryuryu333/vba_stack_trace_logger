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

Private pOutputSheet As Worksheet
Private pCurrentRow As Long

' 出力先シートを設定
Public Sub SetTargetSheet(sheet As Worksheet)
    Set pOutputSheet = sheet
End Sub

Private Sub Logger_ILogWriter_Initialize()
    If pOutputSheet Is Nothing Then
        MsgBox "Logger_ExcelWriter: 出力先のシートが設定されていません。", vbExclamation
        Exit Sub
    End If

    ' ヘッダー行を出力（1行目）
    With pOutputSheet
        .Cells(1, 1).value = "Timestamp"
        .Cells(1, 2).value = "Tag"
        .Cells(1, 3).value = "Module"
        .Cells(1, 4).value = "Procedure"
        .Cells(1, 5).value = "Message"
    End With
    
    With pOutputSheet.Range("A1:E1")
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Font.Bold = True
    End With

    pCurrentRow = 2  ' ログ出力開始行

    Debug.Print "エクセル用のロガーを初期化しました"
End Sub

Private Sub Logger_ILogWriter_WriteLog(logInfo As LoggerLogInfoStruct)
    If pOutputSheet Is Nothing Then Exit Sub

    With pOutputSheet
        .Cells(pCurrentRow, 1).value = logInfo.Timestamp
        .Cells(pCurrentRow, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss.000"
        .Cells(pCurrentRow, 2).value = logInfo.TagName

        If logInfo.IsIncludeCallerInfo Then
            .Cells(pCurrentRow, 3).value = logInfo.ModuleName
            .Cells(pCurrentRow, 4).value = logInfo.ProcedureName
        Else
            .Cells(pCurrentRow, 3).value = ""
            .Cells(pCurrentRow, 4).value = ""
        End If

        .Cells(pCurrentRow, 5).value = logInfo.Message
        
        .Range(.Cells(pCurrentRow, 1), .Cells(pCurrentRow, 5)).VerticalAlignment = xlCenter
        
        ' 最初の出力（2行目）のときだけ列幅を AutoFit、毎回処理するのは重そうなので、最初の1回だけ実施
        If pCurrentRow = 2 Then
            .Columns("A:E").AutoFit
        End If
        
        ' 背景色の適用
        Dim bgColor As Long
        Select Case logInfo.TagType
            Case LogTag_Warning
                bgColor = RGB(255, 255, 153)  ' 薄い黄色
            Case LogTag_Error, LogTag_Critical
                bgColor = RGB(255, 199, 206)  ' 薄い赤色
            Case Else
                bgColor = xlNone  ' デフォルト（塗らない）
        End Select
    
        If bgColor <> xlNone Then
            .Range(.Cells(pCurrentRow, 1), .Cells(pCurrentRow, 5)).Interior.Color = bgColor
        End If
            
        End With
    
    pCurrentRow = pCurrentRow + 1
End Sub

Private Sub Logger_ILogWriter_Terminate()
    pOutputSheet.Columns("A:E").AutoFit
End Sub

