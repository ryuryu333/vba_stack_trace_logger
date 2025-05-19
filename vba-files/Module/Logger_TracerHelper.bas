Attribute VB_Name = "Logger_TracerHelper"
' Logger_TracerHelper.bas
Option Explicit

Public Function UsingTracer(ByVal currentModuleName As String, _
                            ByVal currentProcName As String) As Logger_ProcedureTracer
    Dim tracer As New Logger_ProcedureTracer
    tracer.Init currentModuleName, currentProcName
    Set UsingTracer = tracer
End Function
