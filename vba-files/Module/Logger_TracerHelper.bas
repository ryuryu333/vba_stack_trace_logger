Attribute VB_Name = "Logger_TracerHelper"
' Logger_TracerHelper.bas
Option Explicit

Public Function UsingTracer(ByVal ModuleName As String, _
                            ByVal ProcName As String) As Logger_ProcedureTracer
    Dim tracer As New Logger_ProcedureTracer
    tracer.Init ModuleName, ProcName
    Set UsingTracer = tracer
End Function
