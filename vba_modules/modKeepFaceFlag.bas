Attribute VB_Name = "modKeepFaceFlag"
Option Private Module
Option Explicit

' =============================================================================
' AM列フェイス維持フラグ退避
' =============================================================================
' MDシステムへ渡せない AM 列の維持フラグを履歴保存し、対象セルを空白化します。
' 自動調整・手動調整の前処理として、RunStockControlCore から既定で呼び出されます。

Public Function ExecuteKeepFaceFlagEvacuation(ByVal ws As Worksheet, ByVal wsFlag As Worksheet, ByRef resultNote As String) As Boolean
    ' AM列のフェイス陳列数行にある維持フラグ「2」を、履歴ブックへ保存してから空白化します。
    ' 戻り値が False の場合は履歴保存先が確定しなかったため、呼び出し側で処理を中止してください。
    ' エラーは呼び出し側（RunStockControlCore）のエラー処理へそのまま伝えます。
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

    resultNote = "対象なし"

    colNo = ws.Columns(KEEP_FACE_FLAG_COL).Column
    lastInputRow = ws.Cells(ws.Rows.Count, ITEM_CHECK_COL).End(xlUp).Row

    ' 先に対象行と JAN を集め、履歴保存に成功してから実セルを消します。
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

    If flagCount = 0 Then
        If skippedFormulaCount > 0 Then
            resultNote = "対象なし（数式セルのため消去できない2が " & CStr(skippedFormulaCount) & " 件）"
        End If
        ExecuteKeepFaceFlagEvacuation = True
        Exit Function
    End If

    ' 履歴保存先が確定しない場合は、対象セルを消さずに中止扱いにします。
    saveFolder = ResolveKeepFaceHistoryFolder()
    If Len(saveFolder) = 0 Then
        resultNote = "保存先が確定しなかったため中止"
        Exit Function
    End If

    historyPath = BuildKeepFaceHistoryFilePath(saveFolder)
    SaveKeepFaceHistoryWorkbook janValues, flagCount, historyPath, appendedHistoryCount, duplicateSkippedCount, duplicateDeletedCount

    ' 履歴保存後にだけ対象セルを空白化し、変更印を残します。
    For i = 1 To flagCount
        Set targetCell = ws.Cells(flagRows(i), colNo)
        If IsKeepFaceFlagValue(targetCell.Value) And Not targetCell.HasFormula Then
            targetCell.ClearContents
            targetCell.Interior.Color = WRITTEN_COLOR()
            MarkChanged ws, wsFlag, flagRows(i), colNo
            clearedCount = clearedCount + 1
        End If
    Next i

    resultNote = "退避して消去 " & CStr(clearedCount) & " 件（履歴へ新規保存 " & CStr(appendedHistoryCount) & " 件）"
    If duplicateSkippedCount > 0 Then resultNote = resultNote & "、重複のため追記なし " & CStr(duplicateSkippedCount) & " 件"
    If duplicateDeletedCount > 0 Then resultNote = resultNote & "、履歴内の重複削除 " & CStr(duplicateDeletedCount) & " 件"
    If skippedFormulaCount > 0 Then resultNote = resultNote & "、数式のため未消去 " & CStr(skippedFormulaCount) & " 件"
    resultNote = resultNote & vbCrLf & "　履歴保存先：" & historyPath

    ExecuteKeepFaceFlagEvacuation = True
End Function

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
