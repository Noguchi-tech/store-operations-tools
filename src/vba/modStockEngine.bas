Attribute VB_Name = "modStockEngine"
Option Private Module
Option Explicit

' =============================================================================
' 在庫調整エンジン
' =============================================================================
' 自動調整・手動調整・HB食品用調整の共通本体です。
' 入口側（modStockControl / modManualMode）はモード名・目標手持週数・最下限値などを渡し、
' ここで SKU ブロックごとの販売計画とフェイス陳列数を計算します。

Public Sub RunStockControlCore(ByVal modeName As String, ByVal targetWeeks As Double, ByVal minVal As Long, Optional ByVal targetColsOverride As Variant, Optional ByVal appendTwoWeekSuffix As Boolean = False, Optional ByVal useStoreType2 As Boolean = True)
    Const PROC_NAME As String = "RunStockControlCore"

    On Error GoTo EH

    ' このプロシージャは、自動調整・手動調整・HB食品用調整の共通本体です。
    ' modeName は画面表示と保存名、targetWeeks / minVal はフェイス陳列数の算出条件に使います。
    ' useStoreType2 が False の場合（HB食品用）は、店型2陳列数(本部推奨)を加算にも
    ' 販売予定なし判定にも使わず、本部推奨フェイス1のみで計算します。
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsFlag As Worksheet
    Dim includeIntentioned As Boolean
    Dim appState As ApplicationState
    Dim cols As Variant
    Dim lastInputRow As Long
    Dim planRow As Long
    Dim stockRow As Long
    Dim hqRow As Long
    Dim ratio As Double
    Dim c As Variant
    Dim colLetter As String
    Dim colNo As Long
    Dim planVal As Long
    Dim faceVal As Long
    Dim currentPlan As Double
    Dim planChangedCount As Long
    Dim faceChangedCount As Long
    Dim operatorName As String
    Dim adjustedPath As String
    Dim saveNote As String
    Dim completeMsg As String
    Dim scopeLabel As String
    Dim keepFaceNote As String
    Dim stepName As String

    stepName = "対象ブックと list シートの取得"
    Set wb = GetActiveTargetWorkbookOrNothing()
    If wb Is Nothing Then
        MsgBox "対象ブックが見つかりません。調整したいブックを前面にしてから実行してください。", vbExclamation
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

    stepName = "意思入れ済みセルを含めるか確認"
    If Not ConfirmAdjustmentScope(modeName, includeIntentioned, scopeLabel) Then Exit Sub

    ' 既定値を list シートのデパ名から作るため、ws を渡します。
    stepName = "あなたの名前の入力"
    operatorName = PromptOperatorName(modeName, ws)
    If Len(operatorName) = 0 Then Exit Sub

    stepName = "Excel の再計算と画面更新を一時停止"
    BeginApplicationQuietMode appState

    ' MDシステムへ渡せない AM列の維持フラグは、調整の前処理として既定で退避・消去します。
    stepName = "AM列維持フラグの退避"
    If Not ExecuteKeepFaceFlagEvacuation(ws, wsFlag, keepFaceNote) Then
        MsgBox "AM維持フラグ履歴の保存先が確定しなかったため、" & modeName & "を中止しました。", vbExclamation
        GoTo CLEANUP
    End If

    ' F列の最終入力行まで、22行単位の商品ブロックとして処理します。
    ' planRow が販売計画行、stockRow が同じ商品のフェイス陳列数行を表します。
    If IsMissing(targetColsOverride) Then
        cols = TargetCols()
    Else
        cols = targetColsOverride
    End If
    lastInputRow = ws.Cells(ws.Rows.Count, ITEM_CHECK_COL).End(xlUp).Row

    For planRow = START_ROW To lastInputRow Step STEP_ROW
        stepName = "SKU ブロック判定"
        If ShouldExitSkuLoop(ws.Cells(planRow, ITEM_CHECK_COL).Value) Then Exit For

        stockRow = planRow + STOCKROW_OFFSET
        hqRow = planRow + PLAN_HQ_OFFSET

        ' 直近実績がない新商品相当は、過去実績倍率を使えないため SKU ごと一切触りません。
        If IsLikelyNewProduct(ws, planRow) Then GoTo NEXT_SKU

        ' 本部推奨フェイスが0の対象週は、販売予定なしとして販売計画・フェイスとも一切触りません。
        ' 全対象列が販売予定なしなら、過去実績倍率の計算も行わず SKU ごとスキップします。
        If Not HasAnySalesPlanTargetWeek(ws, planRow, cols, useStoreType2) Then GoTo NEXT_SKU

        stepName = "直近3週の倍率計算"
        ratio = CalcRatioByTermAverage(ws, planRow, hqRow)
        If ratio < 0# Then ratio = 0#

        ' 先に販売計画を書き換え、シート計算後にフェイス陳列数を決めます。
        ' フェイス計算は販売計画の値を参照するため、この順序を変えないでください。
        For Each c In cols
            colLetter = CStr(c)
            colNo = ws.Columns(colLetter).Column

            If Not IsNoSalesPlanWeek(ws, planRow, colNo, useStoreType2) Then
                stepName = "販売計画の自動計算"
                If ShouldAdjustByHqFace(ws, planRow, colNo, minVal) Then
                    If CanWriteTargetCell(ws.Cells(planRow, colNo), includeIntentioned, PLAN_EDITABLE_COLOR()) Then
                        planVal = CalcAutoPlanValue(ws, planRow, hqRow, colNo, ratio)
                        If WriteIfChanged(ws, wsFlag, planRow, colLetter, CDbl(planVal), planRow, stockRow) Then
                            planChangedCount = planChangedCount + 1
                        End If
                    End If
                End If
            End If
        Next c

        stepName = "販売計画反映後のシート再計算"
        ws.Calculate

        ' 販売計画更新後の最新下限・上限を使って、フェイス陳列数を調整します。
        For Each c In cols
            colLetter = CStr(c)
            colNo = ws.Columns(colLetter).Column

            If Not IsNoSalesPlanWeek(ws, planRow, colNo, useStoreType2) Then
                stepName = "フェイス陳列数の自動計算"
                If ShouldAdjustByHqFace(ws, planRow, colNo, minVal) Then
                    If CanWriteTargetCell(ws.Cells(stockRow, colNo), includeIntentioned, FACE_EDITABLE_COLOR()) Then
                        currentPlan = ToDouble(ws.Cells(planRow, colNo).Value)
                        faceVal = GetPriorityFaceQty(ws, planRow, colNo, currentPlan, minVal, targetWeeks, useStoreType2)
                        If WriteIfChanged(ws, wsFlag, stockRow, colLetter, CDbl(faceVal), planRow, stockRow) Then
                            faceChangedCount = faceChangedCount + 1
                        End If
                    ElseIf Not includeIntentioned Then
                        ' 意思入れ済みセルでも、MDシステム取込エラーを避けるためフェイス陳列数上下限だけは守ります。
                        stepName = "意思入れ済みフェイス陳列数の上下限補正"
                        If ClampExistingFaceIfNeeded(ws, wsFlag, stockRow, colLetter, planRow, minVal) Then
                            faceChangedCount = faceChangedCount + 1
                        End If
                    End If
                ElseIf CanWriteTargetCell(ws.Cells(stockRow, colNo), includeIntentioned, FACE_EDITABLE_COLOR()) Then
                    If HqFaceBase(ws, planRow, colNo) > 0# Then
                        stepName = "本部推奨優先フェイス陳列数の調整"
                        If AdjustFaceForLowHqIfNeeded(ws, wsFlag, stockRow, colLetter, planRow, minVal, targetWeeks, useStoreType2) Then
                            faceChangedCount = faceChangedCount + 1
                        End If
                    Else
                        ' 本部推奨フェイス1が0で店型2陳列数(本部推奨)だけに数値がある週は、
                        ' 計算・店型2加算を行わず、シート上下限から外れた既存値の補正だけ行います。
                        ' （HB食品用調整ではこの週自体が販売予定なしとしてスキップされます）
                        stepName = "フェイス1が0の週の上下限補正"
                        If ClampExistingFaceIfNeeded(ws, wsFlag, stockRow, colLetter, planRow, minVal) Then
                            faceChangedCount = faceChangedCount + 1
                        End If
                    End If
                ElseIf Not includeIntentioned Then
                    stepName = "意思入れ済みフェイス陳列数の上下限補正"
                    If ClampExistingFaceIfNeeded(ws, wsFlag, stockRow, colLetter, planRow, minVal) Then
                        faceChangedCount = faceChangedCount + 1
                    End If
                End If
            End If
        Next c

NEXT_SKU:
    Next planRow

    stepName = "ブック全体の再計算"
    Application.CalculateFull

    ' 調整後ブックは条件と操作者が分かる名前で保存し直し、保存成功後に元ファイルを削除します。
    stepName = "調整後ブックの名前変更保存"
    adjustedPath = SaveAdjustedWorkbookWithName(wb, ws, modeName, targetWeeks, minVal, includeIntentioned, operatorName, saveNote, appendTwoWeekSuffix)

    stepName = "完了メッセージ作成"
    completeMsg = modeName & "が完了しました。" & vbCrLf & vbCrLf & _
                  "販売計画修正数：" & CStr(planChangedCount) & " 件" & vbCrLf & _
                  "フェイス陳列修正数：" & CStr(faceChangedCount) & " 件" & vbCrLf & _
                  "修正対象商品：" & scopeLabel & vbCrLf & _
                  "手持週数：" & Format$(targetWeeks, "0.0") & "週" & vbCrLf & _
                  "運用側最下限値：" & CStr(minVal) & vbCrLf & _
                  "保存名：" & IIf(Len(adjustedPath) > 0, adjustedPath, "保存なし") & vbCrLf & _
                  "元ファイル削除結果：" & saveNote & vbCrLf & _
                  "AM維持フラグ退避：" & keepFaceNote

    MsgBox completeMsg, vbInformation

CLEANUP:
    RestoreApplicationState appState
    Exit Sub

EH:
    ReportMacroError PROC_NAME & "（" & modeName & "）", stepName, Err.Number, Err.Description, wb, ws, planRow, stockRow, colNo
    Resume CLEANUP
End Sub
