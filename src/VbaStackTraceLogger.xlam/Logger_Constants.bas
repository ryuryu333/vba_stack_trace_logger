Attribute VB_Name = "Logger_Constants"
' Logger_Constants.bas
Option Private Module
Option Explicit

' === エラー関連 ===
Public Const ERR_VBA_STACK_TRACE_LOGGER As Long = vbObjectError + 513

' === 名前空間 ===
Public Const LOGGER_NAMESPACE As String = "VbaStackTraceLogger"

' === バージョン情報 ===
Public Const LOGGER_VERSION As String = "1.0.0"
Public Const LOGGER_BUILD_DATE As String = "2025-06-06"
Public Const LOGGER_AUTHOR As String = "ryuryu333"
