Attribute VB_Name = "modStockControl"
Option Private Module
Option Explicit

' =============================================================================
' 自動調整モードの入口
' =============================================================================
' リボンの「自動調整」から呼ばれ、標準設定（手持2.0週、最下限5）で
' 共通エンジン RunStockControlCore を実行します。

Public Sub 在庫自動調整実行()
    Const PROC_NAME As String = "在庫自動調整実行"

    On Error GoTo EH

    ' 自動調整は固定条件で実行します。
    ' 具体的な販売計画・フェイス陳列数の計算は共通エンジン側に集約しています。
    RunStockControlCore AutoAdjustmentModeName(), 2#, DEFAULT_MIN_VAL, AutoTargetCols(), True
    Exit Sub

EH:
    ReportMacroError PROC_NAME, "在庫調整エンジン呼び出し", Err.Number, Err.Description, ActiveWorkbook
End Sub
