VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_WriterManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Logger_WriterManager.cls
'@Folder("VbaStackTraceLogger.Core")

Option Explicit
Private Const CLASS_NAME As String = "Logger_WriterManager"

Private mWriters As Collection

Private Sub Class_Initialize()
    Set mWriters = New Collection
End Sub

' === Writer Creation and Management ===
' Create Writer group based on configuration
Public Sub CreateWriters(ByRef config As Logger_ConfigStruct)
    ' Clear collection
    Set mWriters = New Collection
    
    ' Immediate Window (VBE)
    If config.IsWriteToImmediate Then
        Dim immediateWriter As Logger_Writer_Immediate
        Set immediateWriter = New Logger_Writer_Immediate
        mWriters.Add immediateWriter
    End If
    
    ' Excel Sheet (with target sheet configuration)
    If config.IsWriteToExcelSheet Then
        Dim excelWriter As Logger_Writer_Excel
        Set excelWriter = New Logger_Writer_Excel
        excelWriter.SetTargetSheet config.OutputExcelSheet
        mWriters.Add excelWriter
    End If
End Sub

' Initialize all Writers
Public Sub InitializeWriters()
    If mWriters.Count = 0 Then Exit Sub
    
    Dim writer As Logger_ILogWriter
    For Each writer In mWriters
        writer.Initialize
    Next
End Sub

' Send log output command to all Writers
Public Sub WriteLogToAll(ByRef logInfo As Logger_LogInfoStruct)
    If mWriters Is Nothing Then Exit Sub
    If mWriters.Count = 0 Then Exit Sub
    
    Dim writer As Logger_ILogWriter
    For Each writer In mWriters
        writer.WriteLog logInfo
    Next
End Sub

' Terminate all Writers
Public Sub TerminateWriters()
    If mWriters Is Nothing Then Exit Sub
    If mWriters.Count = 0 Then Exit Sub
    
    Dim writer As Logger_ILogWriter
    For Each writer In mWriters
        writer.Terminate
    Next
End Sub

' Get Writer count (for testing)
Public Property Get WriterCount() As Long
    If mWriters Is Nothing Then
        WriterCount = 0
    Else
        WriterCount = mWriters.Count
    End If
End Property

Private Sub Class_Terminate()
    ' Fail-safe
    TerminateWriters
    Set mWriters = Nothing
End Sub
