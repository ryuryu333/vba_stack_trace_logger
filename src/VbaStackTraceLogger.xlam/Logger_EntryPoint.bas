Attribute VB_Name = "Logger_EntryPoint"
' Logger_EntryPoint.bas
Option Explicit

Public Function MyLogger() As Logger_Facade
    Set MyLogger = Logger_SingletonManager.GetMyLogger
End Function

