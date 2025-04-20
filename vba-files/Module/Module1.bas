Attribute VB_Name = "Module1"
Option Explicit
Private Const MODULE_NAME As String = "Module1"

Sub SampleLogUsage()
    Dim myLogger As Logger_Facade
    Set myLogger = GetMyLogger

    Dim config As LoggerConfigStruct
    With config
        .IsIncludeCallerInfo = True
        .IsWriteToImmediate = True
        .IsWriteToExcelSheet = True
        Set .OutputExcelSheet = ActiveSheet
    End With

    myLogger.Initialize config

    myLogger.SetCallerInfo MODULE_NAME, "SampleLogUsage"
    myLogger.Log "始まり"
    myLogger.Log "警告", LogTag_Warning
    myLogger.Log "エラー", LogTag_Error
    myLogger.Log "致命的なエラー", LogTag_Critical
    
    Dim myClass As ClassHoge
    Set myClass = New ClassHoge
    myClass.ClassHogeFunction
    
    myLogger.SetCallerInfo MODULE_NAME, "SampleLogUsage"
    myLogger.Log "終わり" & vbCrLf & "改行あり"
    
    myLogger.Terminate
End Sub


