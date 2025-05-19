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
Private mPreviousCallPath As String
Private mLogger As Logger_Facade

Friend Sub Init(ByVal ModuleName As String, ByVal procName As String)
    mModuleName = ModuleName
    mProcedureName = procName

    Set mLogger = MyLogger
    If Not mLogger Is Nothing Then
        mPreviousCallPath = mLogger.EnterProcedureContext(mModuleName, mProcedureName)
        mLogger.Log ">> Enter " & mProcedureName, LogTag_Trace
    End If
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    If Not mLogger Is Nothing Then
        mLogger.Log "<< Exit  " & mProcedureName, LogTag_Trace
        mLogger.ExitProcedureContext mPreviousCallPath
    End If
End Sub

