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
'@Folder("VbaStackTraceLogger.Core.Writer")

Option Explicit
Implements Logger_ILogWriter

Private Const CLASS_NAME As String = "Logger_ExcelWriter"

' If logger functions are changed, update thins Eunum
Private Enum ColumnType
    ColTimestamp = 1
    ColTag = 2
    ColCallPath = 3
    ColMessage = 4
End Enum

Private mOutputSheet As Worksheet
Private mCurrentRow As Long


Public Sub SetTargetSheet(sheet As Worksheet)
    Set mOutputSheet = sheet
End Sub

Private Sub Logger_ILogWriter_Initialize()
    Debug.Print "Initializing Excel writer..."

    If mOutputSheet Is Nothing Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".Logger_ILogWriter_Initialize", _
            "Unexpected error. mOutputSheet is Nothing."
    End If

    ' Clean up previous cell contents and background colors
    With mOutputSheet.Columns("A:D")
        .ClearContents
        .Interior.ColorIndex = xlNone
    End With

    ' Set up the header row
    With mOutputSheet
        .Cells(1, ColumnType.ColTimestamp).value = "Timestamp"
        .Cells(1, ColumnType.ColTag).value = "Tag"
        .Cells(1, ColumnType.ColCallPath).value = "CallPath"
        .Cells(1, ColumnType.ColMessage).value = "Message"
    End With
    
    ' Set header styles
    With mOutputSheet.Range("A1:D1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
    End With

    ' Start writing from the second row
    mCurrentRow = 2

    Debug.Print "Finished initializing Excel writer."
End Sub

Private Sub Logger_ILogWriter_WriteLog(ByVal logRequest As Logger_LogWriteRequestDTO)
    With mOutputSheet
        ' Write log data to the current row
        .Cells(mCurrentRow, ColumnType.ColTimestamp).value = logRequest.Timestamp
        .Cells(mCurrentRow, ColumnType.ColTimestamp).NumberFormat = "yyyy-mm-dd hh:mm:ss.000"
        .Cells(mCurrentRow, ColumnType.ColTag).value = logRequest.TagName
        .Cells(mCurrentRow, ColumnType.ColCallPath).value = logRequest.CallPath
        .Cells(mCurrentRow, ColumnType.ColMessage).value = logRequest.Message
        
        ' Set background color based on tag type
        Dim bgColor As Long
        Select Case logRequest.TagType
            Case LogTag_Warning
                bgColor = RGB(255, 255, 153)  ' Yellow
            Case LogTag_Error, LogTag_Critical
                bgColor = RGB(255, 199, 206)  ' Red
            Case Else
                bgColor = xlNone  ' No background color
        End Select
        If bgColor <> xlNone Then
            .Range(.Cells(mCurrentRow, ColumnType.ColTimestamp), .Cells(mCurrentRow, ColumnType.ColMessage)).Interior.Color = bgColor
        End If
    End With
    
    ' Move to the next row for the next log entry
    mCurrentRow = mCurrentRow + 1
End Sub

Private Sub Logger_ILogWriter_Terminate()
    mOutputSheet.Columns("A:D").AutoFit
End Sub

