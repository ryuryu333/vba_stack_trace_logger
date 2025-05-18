Attribute VB_Name = "Logger_SingletonManager"
' Logger_SingletonManager.bas
Option Explicit

Private pLogger As Logger_Facade

' グローバルからアクセスするための関数
Public Function MyLogger() As Logger_Facade
    If pLogger Is Nothing Then
        Err.Raise vbObjectError, "Logger_SingletonManager.MyLogger", "Loggerが初期化されていません。先にNewMyLoggerを実行してください。"
    End If
    Set MyLogger = pLogger
End Function

Public Function NewMyLogger() As Logger_Facade
    Set pLogger = New Logger_Facade
    Set NewMyLogger = pLogger
End Function

Public Sub ReleaseMyLogger()
    Set pLogger = Nothing
End Sub
