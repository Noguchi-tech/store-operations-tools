Attribute VB_Name = "modManualMode"
Option Private Module
Option Explicit

' =============================================================================
' 手動調整モードの入口
' =============================================================================
' 目標手持週数とフェイス陳列数の最下限値をユーザー入力で受け取り、
' 共通エンジン RunStockControlCore に渡します。
' HB食品用調整も同じ入力フローを使い、店型2陳列数(本部推奨)だけ使わない設定で実行します。

Public Sub 在庫手動調整実行()
    Const PROC_NAME As String = "在庫手動調整実行"

    On Error GoTo EH

    ' 手動調整は、店型2陳列数(本部推奨)の加算を含む標準の計算で実行します。
    RunManualAdjustmentFlow "手動調整", True
    Exit Sub

EH:
    ReportMacroError PROC_NAME, "手動調整フロー呼び出し", Err.Number, Err.Description, ActiveWorkbook
End Sub

Public Sub 在庫HB食品用調整実行()
    Const PROC_NAME As String = "在庫HB食品用調整実行"

    On Error GoTo EH

    ' HB食品用は、店型2陳列数(本部推奨)を一切使わず、
    ' フェイス陳列数(本部推奨)＝本部推奨フェイス1のみで計算します。
    RunManualAdjustmentFlow "HB食品用調整", False
    Exit Sub

EH:
    ReportMacroError PROC_NAME, "HB食品用調整フロー呼び出し", Err.Number, Err.Description, ActiveWorkbook
End Sub

Private Sub RunManualAdjustmentFlow(ByVal modeName As String, ByVal useStoreType2 As Boolean)
    Const PROC_NAME As String = "RunManualAdjustmentFlow"

    On Error GoTo EH

    Dim strWeeks As String
    Dim targetWeeks As Double
    Dim strMin As String
    Dim minVal As Long
    Dim stepName As String

    ' 手動系の調整では、フェイス陳列数だけに効く目標手持週数をユーザーに指定してもらいます。
    stepName = "目標手持週数の入力"
    strWeeks = InputBox( _
        "目標とする手持週数を入力してください。" & vbCrLf & vbCrLf & _
        "例：1.8 → 販売計画の1.8週分に近いフェイス陳列数へ調整" & vbCrLf & _
        "※販売計画は自動調整と同じロジックで計算し、この数値の影響は受けません。", _
        "手持週数の設定（" & modeName & "）", "2")

    If StrPtr(strWeeks) = 0 Then Exit Sub
    If Not IsNumeric(strWeeks) Then
        MsgBox "数値を入力してください。", vbExclamation, "入力エラー"
        Exit Sub
    End If

    ' 入力値は 0.1 週単位に丸め、共通エンジンへ渡す値をここで確定します。
    targetWeeks = CDbl(strWeeks)
    If targetWeeks <= 0# Then
        MsgBox "0 より大きい値を入力してください。", vbExclamation, "入力エラー"
        Exit Sub
    End If
    targetWeeks = WorksheetFunction.Round(targetWeeks, 1)

    ' シート上の下限値とは別に、運用上の最低フェイス数を指定できるようにします。
    stepName = "フェイス陳列数の最下限値入力"
    strMin = InputBox( _
        "フェイス陳列数の最下限値を入力してください。" & vbCrLf & vbCrLf & _
        "通常は 5 のままで問題ありません。" & vbCrLf & _
        "本部推奨より低く入力済みの値は、本部推奨まで戻します。", _
        "最下限値の設定（" & modeName & "）", CStr(DEFAULT_MIN_VAL))

    If StrPtr(strMin) = 0 Then Exit Sub
    If Not IsNumeric(strMin) Then
        MsgBox "数値を入力してください。", vbExclamation, "入力エラー"
        Exit Sub
    End If

    minVal = CLng(Fix(CDbl(strMin)))
    If minVal < 0 Then
        MsgBox "0 以上の値を入力してください。", vbExclamation, "入力エラー"
        Exit Sub
    End If

    ' 入力チェックが終わったら、以降の計算・書き込みは自動調整と同じ共通処理を使います。
    stepName = "在庫調整エンジン呼び出し"
    RunStockControlCore modeName, targetWeeks, minVal, useStoreType2:=useStoreType2
    Exit Sub

EH:
    ReportMacroError PROC_NAME & "（" & modeName & "）", stepName, Err.Number, Err.Description, ActiveWorkbook
End Sub
