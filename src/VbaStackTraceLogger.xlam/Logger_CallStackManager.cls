VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger_CallStackManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Logger_CallStackManager.cls
Option Explicit
Private Const CLASS_NAME As String = "Logger_CallStackManager"

' コールスタック
Private mCallNestLevel As Long
Private mCallSite As Collection
Private mCallPath As Collection
' Ex. HogeCls.HogeProc にて FugaCls.FugaProc が呼ばれた場合
' mCallNestLevel = 2
' | mCallSite(1) = HogeCls.HogeProc | mCallPath(1) = HogeCls.HogeProc |
' | mCallSite(2) = FugaCls.FugaProc | mCallPath(2) = FugaCls.FugaProc < HogeCls.HogeProc |

Private Sub Class_Initialize()
    Set mCallSite = New Collection
    Set mCallPath = New Collection
    mCallNestLevel = 0
End Sub

' === 新しいプロシージャが呼び出された際の処理 ===
Public Sub EnterProcedure(ByVal callSite As String)
    ' === コールスタック情報を更新 ===
    ' スタックする情報を作成
    Dim newCallSite As String
    newCallSite = callSite
    Dim newCallPath As String
    If mCallNestLevel > 0 Then
        newCallPath = newCallSite & " < " & mCallPath(mCallNestLevel)
    Else
        newCallPath = newCallSite
    End If
    ' 情報を反映
    mCallSite.Add newCallSite
    mCallPath.Add newCallPath
    mCallNestLevel = mCallNestLevel + 1
    ' 整合性の確認
    'ValidateCallStack
End Sub

' === プロシージャが終了した際の処理 ===
Public Sub ExitProcedure()
    ' === コールスタック情報を更新 ===
    If mCallNestLevel = 0 Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".ExitProcedure", _
            "EnterProcedure が呼ばれる前に ExitProcedure が呼び出されました"
    End If
    
    mCallSite.Remove mCallNestLevel
    mCallPath.Remove mCallNestLevel
    mCallNestLevel = mCallNestLevel - 1
    'ValidateCallStack
End Sub

' === スタックトレース情報を返す ===
Public Function GetCurrentCallPath() As String
    If mCallNestLevel = 0 Then
        GetCurrentCallPath = ""
    Else
        GetCurrentCallPath = mCallPath(mCallNestLevel)
    End If
End Function

' === スタックトレース情報の整合性を検証 ===
Private Sub ValidateCallStack()
    Dim isValid As Boolean: isValid = True
    If mCallNestLevel < 0 Then isValid = False
    If mCallSite.count <> mCallNestLevel Then isValid = False
    If mCallPath.count <> mCallNestLevel Then isValid = False
    
    If isValid = False Then
        Err.Raise ERR_VBA_STACK_TRACE_LOGGER, _
            LOGGER_NAMESPACE & "." & CLASS_NAME & ".ValidateCallStack", _
            "コールスタック情報が不正な値になっています"
    End If
End Sub

Private Sub Class_Terminate()
    Set mCallSite = Nothing
    Set mCallPath = Nothing
End Sub
