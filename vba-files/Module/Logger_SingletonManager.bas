Attribute VB_Name = "Logger_SingletonManager"
' Logger_SingletonManager.bas
Option Explicit

Private pLogger As Logger_Facade

' グローバルからアクセスするための関数
Public Function GetMyLogger() As Logger_Facade
    If pLogger Is Nothing Then
        Set pLogger = New Logger_Facade
    End If
    Set GetMyLogger = pLogger
End Function
