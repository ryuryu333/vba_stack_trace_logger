VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_TagArrayWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Logger_TagArrayWrapper.cls
Private tags() As LoggerLogTag

Public Sub Init(ByRef source() As LoggerLogTag)
    tags = source
End Sub

Public Function GetArray() As LoggerLogTag()
    GetArray = tags
End Function
