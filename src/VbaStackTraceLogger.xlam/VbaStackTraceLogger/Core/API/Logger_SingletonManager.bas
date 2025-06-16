Attribute VB_Name = "Logger_SingletonManager"
' Logger_SingletonManager.bas
'@Folder("VbaStackTraceLogger.Core.API")

Option Private Module
Option Explicit

Public mLogger As Logger_Facade

Public Function GetMyLogger() As Logger_Facade
    If mLogger Is Nothing Then
        Set mLogger = New Logger_Facade
    End If
    Set GetMyLogger = mLogger
End Function

' To Debug
' releases the singleton logger instance
Public Sub ReleaseMyLogger()
    Set mLogger = Nothing
End Sub
