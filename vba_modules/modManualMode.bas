Attribute VB_Name = "modManualMode"
Option Private Module
Option Explicit

' =============================================================================
' 手動調整モードの入口
' =============================================================================
' 目標手持週数とフェイス陳列数の最下限値をユーザー入力で受け取り、
' 共通エンジン RunStockControlCore に渡します。

Public Sub 在庫手動調整実行()
    Const PROC_NAME As String = "在庫手動調整実行"

    On Error GoTo EH

    Dim strWeeks As String
    Dim targetWeeks As Double
    Dim strMin As String
    Dim minVal As Long
    Dim stepName As String

    ' 手動調整では、フェイス陳列数だけに効く目標手持週数をユーザーに指定してもらいます。
    stepName = "目標手持週数の入力"
    strWeeks = InputBox( _
        "目標とする手持週数を入力してください。" & vbCrLf & vbCrLf & _
        "例：1.8 → 販売計画の1.8週分に近いフェイス陳列数へ調整" & vbCrLf & _
        "※販売計画は自動調整と同じロジックで計算し、この数値の影響は受けません。", _
        "手持週数の設定", "2")

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
        "最下限値の設定", CStr(DEFAULT_MIN_VAL))

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
    RunStockControlCore "手動調整", targetWeeks, minVal
    Exit Sub

EH:
    ReportMacroError PROC_NAME, stepName, Err.Number, Err.Description, ActiveWorkbook
End Sub
