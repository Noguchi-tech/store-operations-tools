Attribute VB_Name = "modRibbonCallbacks"
Option Explicit

' =============================================================================
' リボンのコールバック
' =============================================================================
' customUI14.xml の onAction から呼ばれる入口だけを置いています。
' 実処理は各モジュールへ渡し、ここではリボンとマクロ本体の対応を保ちます。

' RibbonOnLoad で受け取ったリボン参照です。将来ボタン表示を更新したい場合に使います。
Public gRibbon As IRibbonUI

Public Sub RibbonOnLoad(ByVal ribbon As IRibbonUI)
    On Error GoTo EH

    ' リボンを後で再描画したい場合に使えるよう、参照を保持します。
    Set gRibbon = ribbon
    Exit Sub

EH:
    ReportMacroError "RibbonOnLoad", "リボン参照の保持", Err.Number, Err.Description, ActiveWorkbook
End Sub

Public Sub RunStockAdjustFromRibbon(ByVal control As IRibbonControl)
    On Error GoTo EH

    ' control は Excel から渡されるボタン情報です。現在は処理分岐に使わず、入口の対応確認用に残しています。
    ' customUI14.xml: btnStockAdjust
    Call 在庫自動調整実行
    Exit Sub

EH:
    ReportMacroError "RunStockAdjustFromRibbon", "自動調整ボタン", Err.Number, Err.Description, ActiveWorkbook
End Sub

Public Sub RunManualAdjustFromRibbon(ByVal control As IRibbonControl)
    On Error GoTo EH

    ' リボン側のボタン名と実行マクロの対応をここで固定します。
    ' customUI14.xml: btnManualAdjust
    Call 在庫手動調整実行
    Exit Sub

EH:
    ReportMacroError "RunManualAdjustFromRibbon", "手動調整ボタン", Err.Number, Err.Description, ActiveWorkbook
End Sub

Public Sub RunHbFoodAdjustFromRibbon(ByVal control As IRibbonControl)
    On Error GoTo EH

    ' HB食品用は、店型2陳列数(本部推奨)を使わない手動調整です。
    ' customUI14.xml: btnHbFoodAdjust
    Call 在庫HB食品用調整実行
    Exit Sub

EH:
    ReportMacroError "RunHbFoodAdjustFromRibbon", "HB食品用調整ボタン", Err.Number, Err.Description, ActiveWorkbook
End Sub

Public Sub RunResetIntentionFromRibbon(ByVal control As IRibbonControl)
    On Error GoTo EH

    ' 列単位脱色は値を書き換えず、指定列範囲の色と変更フラグだけを戻します。
    ' customUI14.xml: btnResetIntention
    Call 列単位脱色
    Exit Sub

EH:
    ReportMacroError "RunResetIntentionFromRibbon", "列単位脱色ボタン", Err.Number, Err.Description, ActiveWorkbook
End Sub

Public Sub RunResetSelectedCellsFromRibbon(ByVal control As IRibbonControl)
    On Error GoTo EH

    ' 局所脱色は、選択中のセルだけを行種別に応じた通常色へ戻します。
    ' customUI14.xml: btnResetSelectedCells
    Call 局所脱色
    Exit Sub

EH:
    ReportMacroError "RunResetSelectedCellsFromRibbon", "局所脱色ボタン", Err.Number, Err.Description, ActiveWorkbook
End Sub

Public Sub RunExportSelectionPdfFromRibbon(ByVal control As IRibbonControl)
    On Error GoTo EH

    ' 選択中のセル範囲を A4片面（1ページ）の PDF として保存します。
    ' customUI14.xml: btnExportSelectionPdf
    Call 選択範囲PDF出力
    Exit Sub

EH:
    ReportMacroError "RunExportSelectionPdfFromRibbon", "選択範囲PDF出力ボタン", Err.Number, Err.Description, ActiveWorkbook
End Sub
