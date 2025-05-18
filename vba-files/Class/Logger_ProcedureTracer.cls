VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_ProcedureTracer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Logger_ProcedureTracer.cls
Option Explicit

Private mModuleName As String
Private mProcedureName As String
Private pLogger As Logger_Facade

Friend Sub Init(ByVal ModuleName As String, _
                ByVal ProcName As String)
    mModuleName = ModuleName
    mProcedureName = ProcName

    Set pLogger = MyLogger
    If Not pLogger Is Nothing Then
        pLogger.SetCallerInfo mModuleName, mProcedureName
        pLogger.Log ">> Enter " & mProcedureName, LogTag_Trace
    End If
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    If Not pLogger Is Nothing Then
        pLogger.SetCallerInfo mModuleName, mProcedureName
        pLogger.Log "<< Exit  " & mProcedureName, LogTag_Trace
    End If
End Sub

