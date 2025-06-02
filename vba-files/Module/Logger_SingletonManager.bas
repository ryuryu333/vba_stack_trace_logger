Attribute VB_Name = "Logger_SingletonManager"
' Logger_SingletonManager.bas
Option Explicit

Private mLogger As Logger_Facade

' グローバルからアクセスするための関数
Public Function MyLogger() As Logger_Facade
    If mLogger Is Nothing Then
        Set mLogger = New Logger_Facade
    End If
    Set MyLogger = mLogger
End Function

Public Function NewMyLogger() As Logger_Facade
    Set mLogger = New Logger_Facade
    Set NewMyLogger = mLogger
End Function

Public Sub ReleaseMyLogger()
    Set mLogger = Nothing
End Sub
