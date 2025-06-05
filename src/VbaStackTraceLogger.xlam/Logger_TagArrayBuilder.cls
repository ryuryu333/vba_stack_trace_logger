VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_TagArrayBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Logger_TagArrayBuilder.cls
Private tags() As LoggerLogTag
Private count As Long

Private Sub Class_Initialize()
    ReDim tags(0)
    count = 0
End Sub

Public Function Add(tag As LoggerLogTag) As Logger_TagArrayBuilder
    ReDim Preserve tags(0 To count)
    tags(count) = tag
    count = count + 1
    Set Add = Me
End Function

Public Function ToArray() As Logger_TagArrayWrapper
    Dim wrapper As Logger_TagArrayWrapper: Set wrapper = New Logger_TagArrayWrapper
    wrapper.Init tags
    Set ToArray = wrapper
End Function
