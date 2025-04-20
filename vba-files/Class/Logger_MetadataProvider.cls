VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_MetadataProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Logger_MetadataProvider.cls
Option Explicit

' LogInfo構造体の Timestamp へ記入
Public Sub AttachMetadata(ByRef logInfo As LoggerLogInfoStruct)
    Dim dt As Date
    dt = Now

    ' ミリ秒の計算、Timer の精度の問題でミリ秒単位の正確な時刻を表すものではないのに注意
    Dim milliseconds As Long
    milliseconds = (Timer - Int(Timer)) * 1000

    logInfo.Timestamp = Format$(dt, "yyyy-mm-dd HH:nn:ss") & "." & Format$(milliseconds, "000")
End Sub
