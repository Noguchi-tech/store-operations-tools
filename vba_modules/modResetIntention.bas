Attribute VB_Name = "modResetIntention"
Option Explicit

' =============================================================================
' 脱色
' =============================================================================
' 列単位脱色と局所脱色で、販売計画行とフェイス陳列数行の色・変更フラグを
' 通常状態へ戻します。再エクスポート時の起動マクロによる再着色フラグも消します。

Public Sub 列単位脱色()
    Const PROC_NAME As String = "列単位脱色"

    On Error GoTo EH

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim flagWs As Worksheet
    Dim bfWs As Worksheet
    Dim startColText As String
    Dim endColText As String
    Dim startCol As Long
    Dim endCol As Long
    Dim lastInputRow As Long
    Dim planRow As Long
    Dim stockRow As Long
    Dim c As Long
    Dim appState As ApplicationState
    Dim stepName As String

    stepName = "対象ブックと list シートの取得"
    Set wb = GetActiveTargetWorkbookOrNothing()
    If wb Is Nothing Then
        MsgBox "対象ブックが見つかりません。脱色したいブックを前面にしてから実行してください。", vbExclamation
        Exit Sub
    End If

    Set ws = GetWorksheetOrNothing(wb, SHEET_NAME)
    If ws Is Nothing Then
        MsgBox "Sheet '" & SHEET_NAME & "' が見つかりません。", vbExclamation
        Exit Sub
    End If

    Set flagWs = GetRequiredFlagWorksheet(wb)
    If flagWs Is Nothing Then
        MsgBox "change_flag_sheet が見つかりません。変更フラグを整理できないため、処理を中止します。", vbExclamation
        Exit Sub
    End If
    Set bfWs = GetBfWorksheet(wb)

    ' 開始列・終了列は列記号で受け取り、後続処理では列番号に変換して扱います。
    stepName = "脱色開始列の入力"
    startColText = InputBox("脱色開始列を入力してください。Z～AKの範囲で指定してください。例：Z", "脱色開始列", "Z")
    If Len(Trim$(startColText)) = 0 Then Exit Sub

    stepName = "脱色終了列の入力"
    endColText = InputBox("脱色終了列を入力してください。Z～AKの範囲で指定してください。例：AK", "脱色終了列", "AK")
    If Len(Trim$(endColText)) = 0 Then Exit Sub

    stepName = "列範囲の検証"
    If Not TryGetColumnNumber(ws, startColText, startCol) Then
        MsgBox "開始列の指定が正しくありません。例：Z", vbExclamation, "入力エラー"
        Exit Sub
    End If
    If Not TryGetColumnNumber(ws, endColText, endCol) Then
        MsgBox "終了列の指定が正しくありません。例：AK", vbExclamation, "入力エラー"
        Exit Sub
    End If
    If Not IsResetAllowedColumn(startCol) Or Not IsResetAllowedColumn(endCol) Then
        MsgBox "脱色可能列はZ列からAK列までです。範囲内で指定してください。", vbExclamation, "入力エラー"
        Exit Sub
    End If
    If startCol > endCol Then
        MsgBox "開始列と終了列が逆です。", vbExclamation
        Exit Sub
    End If

    stepName = "Excel の再計算と画面更新を一時停止"
    BeginApplicationQuietMode appState

    ' F列の最終入力行まで、22行単位の商品ブロックとして脱色します。
    ' 各商品ブロック内では、販売計画行とフェイス陳列数行だけを対象にします。
    lastInputRow = ws.Cells(ws.Rows.Count, ITEM_CHECK_COL).End(xlUp).Row

    For planRow = START_ROW To lastInputRow Step STEP_ROW
        stepName = "SKU ブロック判定"
        If ShouldExitSkuLoop(ws.Cells(planRow, ITEM_CHECK_COL).Value) Then Exit For

        stockRow = planRow + STOCKROW_OFFSET

        stepName = "販売計画行とフェイス行の脱色"
        ' list / change_bf_data の色を通常色へ戻し、change_flag_sheet のセル単位フラグを消します。
        For c = startCol To endCol
            ResetColorOneCell ws, flagWs, bfWs, planRow, c, PLAN_EDITABLE_COLOR()
            ResetColorOneCell ws, flagWs, bfWs, stockRow, c, FACE_EDITABLE_COLOR()
        Next c

        stepName = "行単位の変更印を整理"
        ' 行内にフラグが残っていない場合だけ、B列の変更印も消します。
        If Not HasAnyFlagInRow(flagWs, planRow) Then ws.Cells(planRow, CHANGE_ROW_MARK_COL).ClearContents
        If Not HasAnyFlagInRow(flagWs, stockRow) Then ws.Cells(stockRow, CHANGE_ROW_MARK_COL).ClearContents

    Next planRow

    stepName = "完了メッセージ表示"
    MsgBox "脱色が完了しました。" & vbCrLf & _
           "販売計画：RGB(217,225,242)" & vbCrLf & _
           "フェイス陳列数：RGB(255,255,153)" & vbCrLf & _
           "脱色可能列：Z～AK" & vbCrLf & _
           "※フラグ有無に関係なく、オレンジも脱色対象です。", vbInformation

CLEANUP:
    RestoreApplicationState appState
    Exit Sub

EH:
    ReportMacroError PROC_NAME, stepName, Err.Number, Err.Description, wb, ws, planRow, stockRow, c
    Resume CLEANUP
End Sub

Public Sub 局所脱色()
    Const PROC_NAME As String = "局所脱色"

    On Error GoTo EH

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim flagWs As Worksheet
    Dim bfWs As Worksheet
    Dim selectedRange As Range
    Dim area As Range
    Dim cell As Range
    Dim normalColor As Long
    Dim isStockRow As Boolean
    Dim resetCount As Long
    Dim skippedCount As Long
    Dim targetRow As Long
    Dim targetCol As Long
    Dim errorPlanRow As Long
    Dim errorStockRow As Long
    Dim appState As ApplicationState
    Dim stepName As String

    stepName = "対象ブックと list シートの取得"
    Set wb = GetActiveTargetWorkbookOrNothing()
    If wb Is Nothing Then
        MsgBox "対象ブックが見つかりません。脱色したいブックを前面にしてから実行してください。", vbExclamation
        Exit Sub
    End If

    Set ws = GetWorksheetOrNothing(wb, SHEET_NAME)
    If ws Is Nothing Then
        MsgBox "Sheet '" & SHEET_NAME & "' が見つかりません。", vbExclamation
        Exit Sub
    End If

    Set flagWs = GetRequiredFlagWorksheet(wb)
    If flagWs Is Nothing Then
        MsgBox "change_flag_sheet が見つかりません。変更フラグを整理できないため、処理を中止します。", vbExclamation
        Exit Sub
    End If
    Set bfWs = GetBfWorksheet(wb)

    stepName = "選択範囲の確認"
    If Not ActiveSheet Is ws Then
        MsgBox "list シート上で脱色したいセルを選択してから実行してください。", vbExclamation
        Exit Sub
    End If
    If TypeName(Selection) <> "Range" Then
        MsgBox "脱色したいセルを選択してから実行してください。", vbExclamation
        Exit Sub
    End If
    Set selectedRange = Selection

    stepName = "Excel の再計算と画面更新を一時停止"
    BeginApplicationQuietMode appState

    stepName = "選択セルの脱色"
    For Each area In selectedRange.Areas
        For Each cell In area.Cells
            targetRow = cell.Row
            targetCol = cell.Column
            errorPlanRow = 0
            errorStockRow = 0

            If IsResetAllowedColumn(targetCol) And TryGetResetColorForRow(targetRow, normalColor, isStockRow) Then
                If isStockRow Then
                    errorStockRow = targetRow
                Else
                    errorPlanRow = targetRow
                End If

                ResetColorOneCell ws, flagWs, bfWs, targetRow, targetCol, normalColor
                If Not HasAnyFlagInRow(flagWs, targetRow) Then ws.Cells(targetRow, CHANGE_ROW_MARK_COL).ClearContents
                resetCount = resetCount + 1
            Else
                skippedCount = skippedCount + 1
            End If
        Next cell
    Next area

    stepName = "完了メッセージ表示"
    MsgBox "選択セルの脱色が完了しました。" & vbCrLf & _
           "脱色セル数：" & CStr(resetCount) & " 件" & vbCrLf & _
           "対象外セル数：" & CStr(skippedCount) & " 件" & vbCrLf & _
           "脱色可能列：Z～AK" & vbCrLf & _
           "販売計画行は販売計画用の基本色、フェイス陳列数行はフェイス用の基本色へ戻しました。", vbInformation

CLEANUP:
    RestoreApplicationState appState
    Exit Sub

EH:
    ReportMacroError PROC_NAME, stepName, Err.Number, Err.Description, wb, ws, errorPlanRow, errorStockRow, targetCol
    Resume CLEANUP
End Sub

Private Function IsResetAllowedColumn(ByVal colNo As Long) As Boolean
    IsResetAllowedColumn = (colNo >= RESET_ALLOWED_START_COL_NO And colNo <= RESET_ALLOWED_END_COL_NO)
End Function

Private Function TryGetResetColorForRow(ByVal rowNo As Long, ByRef normalColor As Long, ByRef isStockRow As Boolean) As Boolean
    ' 商品ブロック内の販売計画行とフェイス陳列数行だけを、局所脱色の対象にします。
    Dim rowOffset As Long

    If rowNo < START_ROW Then Exit Function

    rowOffset = (rowNo - START_ROW) Mod STEP_ROW
    Select Case rowOffset
        Case 0
            normalColor = PLAN_EDITABLE_COLOR()
            isStockRow = False
            TryGetResetColorForRow = True
        Case STOCKROW_OFFSET
            normalColor = FACE_EDITABLE_COLOR()
            isStockRow = True
            TryGetResetColorForRow = True
    End Select
End Function

Private Sub ResetColorOneCell(ByVal ws As Worksheet, ByVal flagWs As Worksheet, ByVal bfWs As Worksheet, ByVal r As Long, ByVal c As Long, ByVal normalColor As Long)
    ' list と change_bf_data は通常色へ戻し、change_flag_sheet はフラグを消します。
    ws.Cells(r, c).Interior.Color = normalColor

    If Not bfWs Is Nothing Then
        bfWs.Cells(r, c).Interior.Color = normalColor
    End If

    flagWs.Cells(r, c).ClearContents
    flagWs.Cells(r, c).Interior.ColorIndex = xlColorIndexNone

    ClearReexportColorFlag ws, r, c
End Sub

Private Sub ClearReexportColorFlag(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long)
    ' 在庫管理ブック側の起動マクロは、対象列から15列右の 1 / 2 を見て再着色します。
    ' 脱色後にMDシステムへ取り込み、再エクスポートしたときも着色されないよう、このフラグを消します。
    Dim flagCol As Long

    flagCol = c + REEXPORT_COLOR_FLAG_COL_OFFSET
    If flagCol > ws.Columns.Count Then Exit Sub

    If Trim$(CStr(ws.Cells(r, flagCol).Value)) = "1" Or Trim$(CStr(ws.Cells(r, flagCol).Value)) = "2" Then
        ws.Cells(r, flagCol).ClearContents
    End If
End Sub
