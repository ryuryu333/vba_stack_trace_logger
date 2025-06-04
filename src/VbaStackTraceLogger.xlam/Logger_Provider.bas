Attribute VB_Name = "Logger_Provider"
' Logger_Provider.bas
Option Explicit

Public Function MyLogger() As Logger_Facade
    Set MyLogger = Logger_SingletonManager.GetMyLogger
End Function

