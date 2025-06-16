Attribute VB_Name = "Logger_EntryPoint"
' Logger_EntryPoint.bas
'@Folder("VbaStackTraceLogger.Core.API")

Option Explicit

' Private instance for singleton pattern
Private mLogger As Logger_Facade

' Get singleton instance of Logger_Facade
' This module is PUBLIC access level and uses SINGLETON pattern
' User calls MyLogger() DIRECTLY, then always obtains the SAME INSTANCE
' User uses logger by calling MyLogger's methods
' e.g. MyLogger.Log "Message"
Public Function MyLogger() As Logger_Facade
    If mLogger Is Nothing Then
        Set mLogger = Logger_SingletonManager.GetMyLogger
    End If
    Set MyLogger = mLogger
End Function

