VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_IWriterFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for Writer Factory"
'@Folder("VbaStackTraceLogger.Core.Writer")
'@Interface
'@ModuleDescription("Interface for Writer Factory")
Option Explicit

'@Description("Create collection of writers based on configuration")
Public Function CreateWriters(ByVal config As Logger_ConfigurationDTO) As Collection
End Function
