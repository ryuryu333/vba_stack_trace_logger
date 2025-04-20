VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassHoge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const MODULE_NAME As String = "ClassHoge"

Sub ClassHogeFunction()
    Dim myLogger As Logger_Facade
    Set myLogger = GetMyLogger
    myLogger.SetCallerInfo MODULE_NAME, "Class1Function"
    
    myLogger.Log "ÉçÉOÇ≈Ç∑"
    
End Sub

