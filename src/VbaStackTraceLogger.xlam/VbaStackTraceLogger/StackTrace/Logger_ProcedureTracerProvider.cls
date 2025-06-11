VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_ProcedureTracerProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Logger_ProcedureTracerProvider.cls
'@Folder("VbaStackTraceLogger.StackTrace")

Option Explicit
Private Const CLASS_NAME As String = "Logger_ProcedureTracerProvider"

Private mIsStackTraceEnabled As Boolean
Private mIsInitialized As Boolean

Public Sub Initialize(ByVal IsStackTraceEnabled As Boolean)
    mIsStackTraceEnabled = IsStackTraceEnabled
    mIsInitialized = True
End Sub

Public Function CreateTracer(ByVal currentModuleName As String, ByVal currentProcedureName As String, _
                             ByRef stackTraceController As Logger_StackTraceController) As Logger_ProcedureTracer
    EnsureInitialized
    
    ' Return Logger_ProcedureTracer only when stack trace feature is enabled
    ' From user's perspective, receive Variant type variable via Logger.UsingTracer at the beginning of function
    ' Do not directly manipulate Logger_ProcedureTracer
    ' Return Nothing when stack trace feature is disabled (consider interface if functionality becomes complex)
    Dim tracer As Logger_ProcedureTracer
    If mIsStackTraceEnabled Then
        Set tracer = New Logger_ProcedureTracer
        tracer.Initialize currentModuleName, currentProcedureName, stackTraceController
    Else
        Set tracer = Nothing
    End If
    
    Set CreateTracer = tracer
End Function

Private Sub EnsureInitialized()
    If Not mIsInitialized Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".EnsureInitialized", _
            "ProcedureTracerProvider is not initialized"
    End If
End Sub


