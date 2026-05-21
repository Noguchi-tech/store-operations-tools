Attribute VB_Name = "modCalculations"
Option Private Module
Option Explicit

' =============================================================================
' 計算ヘルパー
' =============================================================================
' 販売計画・フェイス陳列数の算出、フェイス陳列数上下限補正、対象判定をまとめます。

Public Function ToDouble(ByVal v As Variant) As Double
    ' 空白・文字列・エラー値を安全に 0 扱いへ寄せます。
    If IsError(v) Then
        ToDouble = 0#
    ElseIf IsNumeric(v) Then
        ToDouble = CDbl(v)
    ElseIf VarType(v) = vbString Then
        Dim s As String
        s = Trim$(CStr(v))
        If IsNumeric(s) Then
            ToDouble = CDbl(s)
        Else
            ToDouble = 0#
        End If
    Else
        ToDouble = 0#
    End If
End Function

Public Function CeilToLong(ByVal x As Double) As Long
    ' Excel の端数誤差を考慮して、正の数だけ切り上げます。
    If x <= 0# Then
        CeilToLong = 0
    ElseIf Abs(x - Fix(x)) <= EPS Then
        CeilToLong = CLng(Fix(x))
    Else
        CeilToLong = CLng(Fix(x) + 1)
    End If
End Function

Public Function ClampLong(ByVal v As Long, ByVal lo As Double, ByVal hi As Double) As Long
    ' 下限と上限の間に値を収めます。上限が下限を下回る場合は下限を優先します。
    Dim lowerLimit As Double
    Dim upperLimit As Double

    lowerLimit = lo
    upperLimit = hi
    If upperLimit < lowerLimit Then upperLimit = lowerLimit

    If CDbl(v) < lowerLimit Then
        ClampLong = CeilToLong(lowerLimit)
    ElseIf CDbl(v) > upperLimit Then
        ClampLong = CLng(Fix(upperLimit))
    Else
        ClampLong = v
    End If
End Function

Public Function GetPlanLowerLimit(ByVal ws As Worksheet, ByVal planRow As Long, ByVal colNo As Long) As Double
    ' 販売計画の下限は planRow + 4 行にあります。
    GetPlanLowerLimit = ToDouble(ws.Cells(planRow + PLAN_LOWER_OFFSET, colNo).Value)
End Function

Public Function GetPlanUpperLimit(ByVal ws As Worksheet, ByVal planRow As Long, ByVal colNo As Long) As Double
    ' 販売計画の上限は planRow + 3 行です。下限未満なら下限まで引き上げます。
    Dim lowerLimit As Double
    Dim upperLimit As Double

    lowerLimit = GetPlanLowerLimit(ws, planRow, colNo)
    upperLimit = ToDouble(ws.Cells(planRow + PLAN_UPPER_OFFSET, colNo).Value)
    If upperLimit < lowerLimit Then upperLimit = lowerLimit

    GetPlanUpperLimit = upperLimit
End Function

Public Function ClampPlanValue(ByVal planVal As Long, ByVal ws As Worksheet, ByVal planRow As Long, ByVal colNo As Long) As Long
    ' 算出した販売計画を、シート上の販売計画下限・上限に収めます。
    Dim lowerLimit As Double
    Dim upperLimit As Double

    lowerLimit = GetPlanLowerLimit(ws, planRow, colNo)
    upperLimit = GetPlanUpperLimit(ws, planRow, colNo)

    ClampPlanValue = ClampLong(planVal, lowerLimit, upperLimit)
End Function

Public Function CalcRatioByTermAverage(ByVal ws As Worksheet, ByVal planRow As Long, ByVal hqRow As Long) As Double
    ' 直近3週（V/W/X列）の自店実績 ÷ 本部計画を平均し、販売計画の倍率にします。
    Dim vAct As Double
    Dim wAct As Double
    Dim xAct As Double
    Dim vHq As Double
    Dim wHq As Double
    Dim xHq As Double
    Dim sumVal As Double
    Dim cnt As Long

    vAct = ToDouble(ws.Cells(planRow, "V").Value)
    wAct = ToDouble(ws.Cells(planRow, "W").Value)
    xAct = ToDouble(ws.Cells(planRow, "X").Value)

    vHq = ToDouble(ws.Cells(hqRow, "V").Value)
    wHq = ToDouble(ws.Cells(hqRow, "W").Value)
    xHq = ToDouble(ws.Cells(hqRow, "X").Value)

    If vHq > 0# Then
        sumVal = sumVal + (vAct / vHq)
        cnt = cnt + 1
    End If
    If wHq > 0# Then
        sumVal = sumVal + (wAct / wHq)
        cnt = cnt + 1
    End If
    If xHq > 0# Then
        sumVal = sumVal + (xAct / xHq)
        cnt = cnt + 1
    End If

    If cnt = 0 Then
        CalcRatioByTermAverage = 0#
    Else
        CalcRatioByTermAverage = sumVal / CDbl(cnt)
    End If
End Function

Public Function CalcAutoPlanValue(ByVal ws As Worksheet, ByVal planRow As Long, ByVal hqRow As Long, ByVal colNo As Long, ByVal ratio As Double) As Long
    ' 本部計画に直近倍率を掛け、販売計画の上下限内に丸めた値を返します。
    Dim hqPlan As Double
    Dim rawPlan As Double
    Dim planVal As Long

    hqPlan = ToDouble(ws.Cells(hqRow, colNo).Value)

    If hqPlan > 0# And ratio > 0# Then
        rawPlan = hqPlan * ratio
    Else
        rawPlan = 0#
    End If

    planVal = CeilToLong(rawPlan)
    CalcAutoPlanValue = ClampPlanValue(planVal, ws, planRow, colNo)
End Function

Public Function GetEffectiveFaceLower(ByVal lowerFace As Double, Optional ByVal minVal As Long = DEFAULT_MIN_VAL) As Long
    ' 通常調整では、シート下限値と実行時の最下限の大きい方をフェイス陳列数の下限にします。
    Dim lo As Long
    lo = CeilToLong(lowerFace)

    If minVal < 0 Then minVal = 0

    If lo > minVal Then
        GetEffectiveFaceLower = lo
    Else
        GetEffectiveFaceLower = minVal
    End If
End Function

Public Function GetSheetFaceLowerLimit(ByVal ws As Worksheet, ByVal planRow As Long, ByVal colNo As Long) As Long
    ' MDシステム取込用の実下限として、シート上のフェイス陳列数下限行だけを読みます。
    GetSheetFaceLowerLimit = CeilToLong(ToDouble(ws.Cells(planRow + DISP_LOWER_OFFSET, colNo).Value))
End Function

Public Function GetFaceUpperLimit(ByVal ws As Worksheet, ByVal planRow As Long, ByVal colNo As Long) As Long
    ' フェイス陳列数の上限が空白・0 の場合は MAX_COUNT を上限として扱います。
    Dim upperFace As Double
    upperFace = ToDouble(ws.Cells(planRow + FACE_UPPER_OFFSET, colNo).Value)

    If upperFace <= 0# Then
        GetFaceUpperLimit = MAX_COUNT
    ElseIf upperFace > MAX_COUNT Then
        GetFaceUpperLimit = MAX_COUNT
    Else
        GetFaceUpperLimit = CLng(Fix(upperFace))
    End If
End Function

Public Function GetBestFaceQty(ByVal planVal As Double, ByVal lowerFace As Double, ByVal upperFace As Double, Optional ByVal minVal As Long = DEFAULT_MIN_VAL, Optional ByVal targetWeeks As Double = 2#) As Long
    ' 販売計画 × 目標手持週数に最も近い整数を、フェイス陳列数上下限内で選びます。
    Dim lo As Long
    Dim hi As Long
    Dim targetQty As Double
    Dim lowCandidate As Long
    Dim highCandidate As Long
    Dim lowDiff As Double
    Dim highDiff As Double

    lo = GetEffectiveFaceLower(lowerFace, minVal)
    hi = CLng(Fix(upperFace))
    If hi < lo Then hi = lo

    If planVal <= 0# Then
        GetBestFaceQty = lo
        Exit Function
    End If

    If targetWeeks <= 0# Then targetWeeks = 2#
    targetQty = targetWeeks * planVal

    If targetQty <= CDbl(lo) Then
        GetBestFaceQty = lo
        Exit Function
    End If

    If targetQty >= CDbl(hi) Then
        GetBestFaceQty = hi
        Exit Function
    End If

    lowCandidate = CLng(Fix(targetQty))
    highCandidate = lowCandidate + 1

    If lowCandidate < lo Then lowCandidate = lo
    If lowCandidate > hi Then lowCandidate = hi
    If highCandidate < lo Then highCandidate = lo
    If highCandidate > hi Then highCandidate = hi

    lowDiff = Abs((CDbl(lowCandidate) / planVal) - targetWeeks)
    highDiff = Abs((CDbl(highCandidate) / planVal) - targetWeeks)

    If lowDiff <= (highDiff + EPS) Then
        GetBestFaceQty = lowCandidate
    Else
        GetBestFaceQty = highCandidate
    End If
End Function

Public Function GetPriorityFaceQty(ByVal ws As Worksheet, ByVal planRow As Long, ByVal colNo As Long, ByVal planVal As Double, Optional ByVal minVal As Long = DEFAULT_MIN_VAL, Optional ByVal targetWeeks As Double = 2#) As Long
    ' 優先順位は「シート上下限 > 手持週数 > 運用下限値/HQ関係」。売れ筋を機械的に下げないため、手持週数側の必要数を先に見ます。
    Dim lowerFace As Double
    Dim upperFace As Long
    Dim hqFace As Double
    Dim handWeeksVal As Long
    Dim relationFloor As Long
    Dim candidateVal As Long

    If minVal < 0 Then minVal = 0

    lowerFace = ToDouble(ws.Cells(planRow + DISP_LOWER_OFFSET, colNo).Value)
    upperFace = GetFaceUpperLimit(ws, planRow, colNo)
    hqFace = HqFaceTotal(ws, planRow, colNo)

    handWeeksVal = GetBestFaceQty(planVal, lowerFace, CDbl(upperFace), 0, targetWeeks)

    If hqFace > CDbl(minVal) Then
        relationFloor = minVal
    ElseIf hqFace > 0# Then
        relationFloor = CeilToLong(hqFace)
    Else
        relationFloor = 0
    End If

    If handWeeksVal >= relationFloor Then
        candidateVal = handWeeksVal
    Else
        candidateVal = relationFloor
    End If

    GetPriorityFaceQty = ClampFaceValueToSheetLimits(candidateVal, ws, planRow, colNo)
End Function

Public Function ClampFaceValueToCurrentLimits(ByVal faceVal As Long, ByVal ws As Worksheet, ByVal planRow As Long, ByVal colNo As Long, Optional ByVal minVal As Long = DEFAULT_MIN_VAL) As Long
    ' 既存のフェイス陳列数を、再計算後の上下限に合わせて補正します。
    Dim lowerFace As Double
    Dim lowerLimit As Long
    Dim upperLimit As Long
    Dim candidateVal As Long

    lowerFace = ToDouble(ws.Cells(planRow + DISP_LOWER_OFFSET, colNo).Value)
    lowerLimit = GetEffectiveFaceLower(lowerFace, minVal)
    upperLimit = GetFaceUpperLimit(ws, planRow, colNo)
    If upperLimit < lowerLimit Then upperLimit = lowerLimit

    candidateVal = ClampLong(faceVal, CDbl(lowerLimit), CDbl(upperLimit))
    ClampFaceValueToCurrentLimits = ClampFaceValueToSheetLimits(candidateVal, ws, planRow, colNo)
End Function

Public Function ClampFaceValueToSheetLimits(ByVal faceVal As Long, ByVal ws As Worksheet, ByVal planRow As Long, ByVal colNo As Long) As Long
    ' 本部推奨を運用側下限より優先する場合でも、MDシステム取込用のシート上下限は必ず守ります。
    Dim lowerLimit As Long
    Dim upperLimit As Long

    lowerLimit = GetSheetFaceLowerLimit(ws, planRow, colNo)
    upperLimit = GetFaceUpperLimit(ws, planRow, colNo)
    If upperLimit < lowerLimit Then upperLimit = lowerLimit

    ClampFaceValueToSheetLimits = ClampLong(faceVal, CDbl(lowerLimit), CDbl(upperLimit))
End Function

Public Function ClampExistingFaceIfNeeded(ByVal ws As Worksheet, ByVal wsFlag As Worksheet, ByVal stockRow As Long, ByVal colLetter As String, ByVal planRow As Long, Optional ByVal minVal As Long = DEFAULT_MIN_VAL) As Boolean
    ' 意思入れ済みセルを書き換えない設定でも、MD取込用シート上下限から外れた既存値だけは補正します。
    Dim colNo As Long
    Dim stockCell As Range
    Dim currentRaw As Variant
    Dim currentVal As Long
    Dim clampedVal As Long

    colNo = ws.Columns(colLetter).Column
    Set stockCell = ws.Cells(stockRow, colNo)
    If stockCell.HasFormula Then Exit Function

    currentRaw = stockCell.Value
    If IsError(currentRaw) Then Exit Function
    If Not IsNumeric(currentRaw) Then Exit Function

    currentVal = CLng(Fix(CDbl(currentRaw)))
    clampedVal = ClampFaceValueToSheetLimits(currentVal, ws, planRow, colNo)

    If Abs(CDbl(currentRaw) - CDbl(clampedVal)) <= EPS Then Exit Function

    WriteIfChanged ws, wsFlag, stockRow, colLetter, CDbl(clampedVal), planRow, stockRow
    ClampExistingFaceIfNeeded = True
End Function

Public Function AdjustFaceForLowHqIfNeeded(ByVal ws As Worksheet, ByVal wsFlag As Worksheet, ByVal stockRow As Long, ByVal colLetter As String, ByVal planRow As Long, Optional ByVal minVal As Long = DEFAULT_MIN_VAL, Optional ByVal targetWeeks As Double = 2#) As Boolean
    ' 本部推奨が運用側下限以下の場合も、手持週数を優先しつつ本部推奨側へ寄せます。
    Dim colNo As Long
    Dim stockCell As Range
    Dim currentRaw As Variant
    Dim currentVal As Double
    Dim currentPlan As Double
    Dim hqFace As Double
    Dim targetVal As Long

    If minVal < 0 Then minVal = 0

    colNo = ws.Columns(colLetter).Column
    Set stockCell = ws.Cells(stockRow, colNo)
    If stockCell.HasFormula Then Exit Function

    currentRaw = stockCell.Value
    If IsError(currentRaw) Then Exit Function
    If Not IsNumeric(currentRaw) Then Exit Function

    currentVal = CDbl(currentRaw)
    hqFace = HqFaceTotal(ws, planRow, colNo)
    If hqFace <= 0# Then Exit Function
    If hqFace > CDbl(minVal) Then Exit Function

    currentPlan = ToDouble(ws.Cells(planRow, colNo).Value)
    targetVal = GetPriorityFaceQty(ws, planRow, colNo, currentPlan, minVal, targetWeeks)
    If Abs(currentVal - CDbl(targetVal)) <= EPS Then Exit Function

    WriteIfChanged ws, wsFlag, stockRow, colLetter, CDbl(targetVal), planRow, stockRow
    AdjustFaceForLowHqIfNeeded = True
End Function

Public Function IsNoSalesPlanWeek(ByVal ws As Worksheet, ByVal planRow As Long, ByVal colNo As Long) As Boolean
    ' 本部推奨フェイス2行がどちらも0なら、その週は販売予定なしとして一切書き換えません。
    IsNoSalesPlanWeek = _
        Abs(ToDouble(ws.Cells(planRow + FACE_HQ_1_OFFSET, colNo).Value)) <= EPS And _
        Abs(ToDouble(ws.Cells(planRow + FACE_HQ_2_OFFSET, colNo).Value)) <= EPS
End Function

Public Function HasAnySalesPlanTargetWeek(ByVal ws As Worksheet, ByVal planRow As Long, ByVal cols As Variant) As Boolean
    ' 対象列の全週が販売予定なしなら、過去実績倍率の計算も含めて SKU 全体を触りません。
    Dim c As Variant
    Dim colNo As Long

    For Each c In cols
        colNo = ws.Columns(CStr(c)).Column
        If Not IsNoSalesPlanWeek(ws, planRow, colNo) Then
            HasAnySalesPlanTargetWeek = True
            Exit Function
        End If
    Next c
End Function

Public Function IsZero3Weeks(ByVal ws As Worksheet, ByVal targetRow As Long) As Boolean
    ' V/W/X列の3週分がすべてゼロなら True。
    IsZero3Weeks = _
        Abs(ToDouble(ws.Cells(targetRow, "V").Value)) <= EPS And _
        Abs(ToDouble(ws.Cells(targetRow, "W").Value)) <= EPS And _
        Abs(ToDouble(ws.Cells(targetRow, "X").Value)) <= EPS
End Function

Public Function IsLikelyNewProduct(ByVal ws As Worksheet, ByVal planRow As Long) As Boolean
    ' 直近の販売計画と実績がない新商品相当は、過去実績倍率を使えないため SKU ごと触りません。
    IsLikelyNewProduct = _
        IsZero3Weeks(ws, planRow) And _
        IsZero3Weeks(ws, planRow + STOCK_ACTUAL_OFFSET)
End Function

Public Function IsWritableRow(ByVal r As Long, ByVal planRow As Long, ByVal stockRow As Long) As Boolean
    ' このマクロが値を書いてよいのは販売計画行とフェイス陳列数行だけです。
    IsWritableRow = (r = planRow Or r = stockRow)
End Function

Public Function CanWriteTargetCell(ByVal targetCell As Range, ByVal includeIntentioned As Boolean, ByVal requiredColor As Long) As Boolean
    ' 数式セルは保護します。全対象モードでは色を問わず、未意思入れモードでは通常色とマクロ書込済みセルだけを書き換えます。
    If targetCell.HasFormula Then Exit Function

    If includeIntentioned Then
        CanWriteTargetCell = True
    Else
        CanWriteTargetCell = _
            targetCell.Interior.Color = requiredColor Or _
            targetCell.Interior.Color = WRITTEN_COLOR()
    End If
End Function

Public Function HqFaceTotal(ByVal ws As Worksheet, ByVal planRow As Long, ByVal colNo As Long) As Double
    ' 本部フェイスは2行に分かれているため合算して判定します。
    HqFaceTotal = ToDouble(ws.Cells(planRow + FACE_HQ_1_OFFSET, colNo).Value) + _
                  ToDouble(ws.Cells(planRow + FACE_HQ_2_OFFSET, colNo).Value)
End Function

Public Function ShouldAdjustByHqFace(ByVal ws As Worksheet, ByVal planRow As Long, ByVal colNo As Long, ByVal minVal As Long) As Boolean
    ' 通常の自動計算対象は、本部推奨フェイスが実行時の下限値を超える商品です。
    ' 販売予定なしや下限値以下の本部推奨は、呼び出し側で別ルールとして扱います。
    If minVal < 0 Then minVal = 0
    ShouldAdjustByHqFace = (HqFaceTotal(ws, planRow, colNo) > CDbl(minVal))
End Function
