VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_WriterFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Factory class for creating log writers"
'@Folder("VbaStackTraceLogger.Core.Writer")
Option Explicit

Implements Logger_IWriterFactory

Private Const CLASS_NAME As String = "Logger_WriterFactory"

Public Function CreateWriters(ByVal config As Logger_ConfigurationDTO) As Collection
    Dim writers As Collection
    Set writers = New Collection
    
    ' Immediate Window (VBE)
    If config.IsWriteToImmediate Then
        Dim immediateWriter As Logger_Writer_Immediate
        Set immediateWriter = New Logger_Writer_Immediate
        writers.Add immediateWriter
    End If
    
    ' Excel Sheet (with target sheet configuration)
    If config.IsWriteToExcelSheet Then
        Dim excelWriter As Logger_Writer_Excel
        Set excelWriter = New Logger_Writer_Excel
        excelWriter.SetTargetSheet config.OutputExcelSheet
        writers.Add excelWriter
    End If
    
    Set CreateWriters = writers
End Function

' === Interface Implementation ===

' Logger_IWriterFactory implementation
Private Function Logger_IWriterFactory_CreateWriters(ByVal config As Logger_ConfigurationDTO) As Collection
    Set Logger_IWriterFactory_CreateWriters = CreateWriters(config)
End Function

