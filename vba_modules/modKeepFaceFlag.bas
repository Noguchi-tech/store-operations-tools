Attribute VB_Name = "modKeepFaceFlag"
Option Private Module
Option Explicit

' =============================================================================
' AM列フェイス維持フラグ退避
' =============================================================================
' MDシステムへ渡せない AM 列の維持フラグを履歴保存し、対象セルを空白化します。

Public Sub AM列フェイス維持フラグ退避()
    Const PROC_NAME As String = "AM列フェイス維持フラグ退避"

    On Error GoTo EH

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsFlag As Worksheet
    Dim appState As ApplicationState
    Dim ans As VbMsgBoxResult
    Dim lastInputRow As Long
    Dim planRow As Long
    Dim stockRow As Long
    Dim colNo As Long
    Dim targetCell As Range
    Dim flagRows() As Long
    Dim janValues() As String
    Dim flagCount As Long
    Dim clearedCount As Long
    Dim skippedFormulaCount As Long
    Dim saveFolder As String
    Dim historyPath As String
    Dim i As Long
    Dim appendedHistoryCount As Long
    Dim duplicateSkippedCount As Long
    Dim duplicateDeletedCount As Long
    Dim completeMsg As String
    Dim stepName As String

    stepName = "対象ブックと list シートの取得"
    Set wb = GetActiveTargetWorkbookOrNothing()
    If wb Is Nothing Then
        MsgBox "対象ブックが見つかりません。処理したいブックを前面にしてから実行してください。", vbExclamation
        Exit Sub
    End If

    Set ws = GetWorksheetOrNothing(wb, SHEET_NAME)
    If ws Is Nothing Then
        MsgBox "Sheet '" & SHEET_NAME & "' が見つかりません。", vbExclamation
        Exit Sub
    End If
    Set wsFlag = GetRequiredFlagWorksheet(wb)
    If wsFlag Is Nothing Then
        MsgBox "change_flag_sheet が見つかりません。MDシステム取込用の変更フラグを記録できないため、処理を中止します。", vbExclamation
        Exit Sub
    End If

    ' いきなり消去せず、履歴ブックを作ることを実行者に確認してから進めます。
    ans = MsgBox( _
        "AM列のフェイス陳列数行にある維持フラグ「2」のJANを、別ブックに保存してから空白にします。" & vbCrLf & vbCrLf & _
        "MDシステムへインポートする前の前処理として実行してください。" & vbCrLf & _
        "続行しますか？", _
        vbOKCancel + vbQuestion, _
        "AM列維持フラグ退避")
    If ans = vbCancel Then Exit Sub

    colNo = ws.Columns(KEEP_FACE_FLAG_COL).Column
    lastInputRow = ws.Cells(ws.Rows.Count, ITEM_CHECK_COL).End(xlUp).Row

    ' 先に対象行と JAN を集め、履歴保存に成功してから実セルを消します。
    stepName = "AM列維持フラグ対象の収集"
    For planRow = START_ROW To lastInputRow Step STEP_ROW
        If ShouldExitSkuLoop(ws.Cells(planRow, ITEM_CHECK_COL).Value) Then Exit For

        stockRow = planRow + STOCKROW_OFFSET
        Set targetCell = ws.Cells(stockRow, colNo)

        If IsKeepFaceFlagValue(targetCell.Value) Then
            If targetCell.HasFormula Then
                skippedFormulaCount = skippedFormulaCount + 1
            Else
                flagCount = flagCount + 1
                ReDim Preserve flagRows(1 To flagCount)
                ReDim Preserve janValues(1 To flagCount)

                flagRows(flagCount) = stockRow
                janValues(flagCount) = GetKeepFaceHistoryJan(ws, planRow, stockRow)
            End If
        End If
    Next planRow

    If flagCount > 0 Then
        stepName = "AM維持フラグ履歴ブックの保存先取得"
        saveFolder = ResolveKeepFaceHistoryFolder()
        If Len(saveFolder) = 0 Then Exit Sub

        stepName = "AM維持フラグ履歴ブックの保存"
        historyPath = BuildKeepFaceHistoryFilePath(saveFolder)
        SaveKeepFaceHistoryWorkbook janValues, flagCount, historyPath, appendedHistoryCount, duplicateSkippedCount, duplicateDeletedCount

        ' 履歴保存後にだけ対象セルを空白化し、変更印を残します。
        stepName = "Excel の再計算と画面更新を一時停止"
        BeginApplicationQuietMode appState

        stepName = "AM列維持フラグの消去"
        For i = 1 To flagCount
            Set targetCell = ws.Cells(flagRows(i), colNo)
            If IsKeepFaceFlagValue(targetCell.Value) And Not targetCell.HasFormula Then
                targetCell.ClearContents
                targetCell.Interior.Color = WRITTEN_COLOR()
                MarkChanged ws, wsFlag, flagRows(i), colNo
                clearedCount = clearedCount + 1
            End If
        Next i
    End If

    stepName = "完了メッセージ表示"
    completeMsg = "AM列の維持フラグ退避が完了しました。" & vbCrLf & _
                  "退避して消去した件数: " & CStr(clearedCount) & vbCrLf & _
                  "履歴へ新規保存したJAN件数: " & CStr(appendedHistoryCount) & vbCrLf & _
                  "重複のため追記しなかったJAN件数: " & CStr(duplicateSkippedCount) & vbCrLf & _
                  "数式セルのため消去しなかった件数: " & CStr(skippedFormulaCount)
    If duplicateDeletedCount > 0 Then
        completeMsg = completeMsg & vbCrLf & "履歴ブック内の重複削除件数: " & CStr(duplicateDeletedCount)
    End If
    If clearedCount > 0 Then
        completeMsg = completeMsg & vbCrLf & "保存先: " & historyPath
    Else
        completeMsg = completeMsg & vbCrLf & "退避・消去できる2はありませんでした。"
    End If
    MsgBox completeMsg, vbInformation

CLEANUP:
    RestoreApplicationState appState
    Exit Sub

EH:
    ReportMacroError PROC_NAME, stepName, Err.Number, Err.Description, wb, ws, planRow, stockRow, colNo
    Resume CLEANUP
End Sub

Public Function IsKeepFaceFlagValue(ByVal v As Variant) As Boolean
    ' AM列維持フラグとして扱う値かを、数値誤差込みで判定します。
    If IsError(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function

    IsKeepFaceFlagValue = (Abs(CDbl(v) - CDbl(KEEP_FACE_FLAG_VALUE)) <= EPS)
End Function

Public Function GetKeepFaceHistoryJan(ByVal ws As Worksheet, ByVal planRow As Long, ByVal stockRow As Long) As String
    ' 履歴ブックに残す JAN は、フェイス行を優先し、空なら販売計画行から拾います。
    Dim janText As String

    janText = CellValueAsText(ws.Cells(stockRow, ITEM_CHECK_COL))
    If Len(janText) = 0 Then janText = CellValueAsText(ws.Cells(planRow, ITEM_CHECK_COL))

    GetKeepFaceHistoryJan = janText
End Function

Public Function CellValueAsText(ByVal cell As Range) As String
    ' JAN が指数表記にならないよう、セル値を保存用の文字列へ整えます。
    Dim v As Variant
    Dim s As String

    v = cell.Value2
    If IsError(v) Then Exit Function

    s = Trim$(CStr(v))
    If Len(s) = 0 Then Exit Function

    If IsNumeric(v) Then
        If InStr(1, s, "E", vbTextCompare) > 0 Or InStr(1, s, ".") > 0 Then
            s = Format$(CDbl(v), "0")
        End If
    End If

    CellValueAsText = s
End Function
