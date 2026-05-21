Attribute VB_Name = "modCommon"
Option Private Module
Option Explicit

' =============================================================================
' 共通設定・対象ブック・変更記録ヘルパー
' =============================================================================
' 全モジュールから参照する定数、対象ブック取得、Excel状態復元、変更印をまとめます。
' 計算、ファイル保存、設定、AM退避の処理本体は専用モジュールへ分割しています。

' 対象ブックに必要なシート名です。
Public Const SHEET_NAME As String = "list"
Public Const FLAG_SHEET_NAME As String = "change_flag_sheet"
Public Const BF_SHEET_NAME As String = "change_bf_data"
Public Const SETTING_SHEET_NAME As String = "setting"

' list シートは 1 商品が 22 行ブロックで並ぶ前提です。
Public Const START_ROW As Long = 8
Public Const STEP_ROW As Long = 22
Public Const MAX_COUNT As Long = 300
Public Const ITEM_CHECK_COL As String = "F"
Public Const CHANGE_ROW_MARK_COL As Long = 2
Public Const EPS As Double = 0.0000001

' AM列のフェイス陳列数行に入る「2」はMDシステムに渡せない維持フラグです。
Public Const KEEP_FACE_FLAG_COL As String = "AM"
Public Const KEEP_FACE_FLAG_VALUE As Long = 2
Public Const KEEP_FACE_HISTORY_FILE_SUFFIX As String = "AM維持フラグ商品履歴"
Public Const KEEP_FACE_HISTORY_SETTING_KEY As String = "AM維持フラグ履歴保存先"
Public Const ADJUSTED_BOOK_EXCLUDE_INTENTIONED_SUFFIX As String = "_未意思入のみ"
' 再エクスポート時の起動マクロが読む着色フラグは、対象列から15列右にあります。
Public Const REEXPORT_COLOR_FLAG_COL_OFFSET As Long = 15
' 脱色できる列範囲です。Z列からAK列までに限定します。
Public Const RESET_ALLOWED_START_COL_NO As Long = 26
Public Const RESET_ALLOWED_END_COL_NO As Long = 37

' フェイス陳列数の運用側最下限です。
' 手動調整ではユーザー入力値で上書きできます。
Public Const DEFAULT_MIN_VAL As Long = 5

' planRow（販売計画行）から見た各参照行の相対位置です。
Public Const PLAN_HQ_OFFSET As Long = -2
Public Const PLAN_UPPER_OFFSET As Long = 3
Public Const PLAN_LOWER_OFFSET As Long = 4
Public Const FACE_UPPER_OFFSET As Long = 6
Public Const DISP_LOWER_OFFSET As Long = 7
Public Const FACE_HQ_1_OFFSET As Long = 12
Public Const FACE_HQ_2_OFFSET As Long = 13
Public Const STOCKROW_OFFSET As Long = 14
Public Const STOCK_ACTUAL_OFFSET As Long = 16

' 画面更新停止などを安全に戻すため、変更前の Excel 状態をまとめて保持します。
Public Type ApplicationState
    EnableEvents As Boolean
    ScreenUpdating As Boolean
    Calculation As XlCalculation
    Captured As Boolean
End Type

Public Function TargetCols() As Variant
    ' 手動調整と既定処理の対象列です。自動調整は AutoTargetCols() を使います。
    TargetCols = Array("Z")
End Function

Public Function AutoTargetCols() As Variant
    ' 自動調整対象の列です。2週間分として Z列とAA列を処理します。
    AutoTargetCols = Array("Z", "AA")
End Function

Public Function AutoAdjustmentModeName() As String
    AutoAdjustmentModeName = ChrW$(&H81EA) & ChrW$(&H52D5) & ChrW$(&H8ABF) & ChrW$(&H6574)
End Function

Public Function PLAN_EDITABLE_COLOR() As Long
    ' 販売計画の「マクロが触ってよい」通常色。
    PLAN_EDITABLE_COLOR = RGB(217, 225, 242)
End Function

Public Function FACE_EDITABLE_COLOR() As Long
    ' フェイス陳列数の「マクロが触ってよい」通常色。
    FACE_EDITABLE_COLOR = RGB(255, 255, 153)
End Function

Public Function WRITTEN_COLOR() As Long
    ' マクロが値を書き換えたセルを後から目視確認するための色。
    WRITTEN_COLOR = RGB(255, 128, 128)
End Function

Public Function ORANGE_COLOR() As Long
    ' 手入力や意思入れ済みを示す運用色です。
    ORANGE_COLOR = RGB(255, 153, 0)
End Function

Public Function GetActiveTargetWorkbookOrNothing() As Workbook
    ' アドイン自身ではなく、ユーザーが操作しているブックを対象にします。
    If Application.Workbooks.Count = 0 Then Exit Function
    If ActiveWorkbook Is Nothing Then Exit Function
    If ActiveWorkbook Is ThisWorkbook Then Exit Function

    Set GetActiveTargetWorkbookOrNothing = ActiveWorkbook
End Function

Public Function GetWorksheetOrNothing(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    ' シート未存在をエラーにせず Nothing で返す共通入口です。
    If wb Is Nothing Then Exit Function

    On Error Resume Next
    Set GetWorksheetOrNothing = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Public Function GetRequiredFlagWorksheet(ByVal wb As Workbook) As Worksheet
    ' change_flag_sheet は運用上必須です。呼び出し側で Nothing を検知したら処理を止めます。
    Set GetRequiredFlagWorksheet = GetWorksheetOrNothing(wb, FLAG_SHEET_NAME)
End Function

Public Function GetBfWorksheet(ByVal wb As Workbook) As Worksheet
    ' 調整前データを保持する補助シートを取得します。
    Set GetBfWorksheet = GetWorksheetOrNothing(wb, BF_SHEET_NAME)
End Function

Public Sub CaptureApplicationState(ByRef state As ApplicationState)
    ' 後で必ず元に戻せるよう、Excel の実行状態を一度だけ退避します。
    If state.Captured Then Exit Sub

    state.EnableEvents = Application.EnableEvents
    state.ScreenUpdating = Application.ScreenUpdating
    state.Calculation = Application.Calculation
    state.Captured = True
End Sub

Public Sub BeginApplicationQuietMode(ByRef state As ApplicationState)
    ' 大量セル更新中のちらつきと途中再計算を抑え、処理速度を安定させます。
    CaptureApplicationState state

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
End Sub

Public Sub RestoreApplicationState(ByRef state As ApplicationState)
    ' 復元中の例外で本来のエラー表示を邪魔しないよう、この中だけ握ります。
    If Not state.Captured Then Exit Sub

    On Error Resume Next
    Application.EnableEvents = state.EnableEvents
    Application.ScreenUpdating = state.ScreenUpdating
    Application.Calculation = state.Calculation
    state.Captured = False
    On Error GoTo 0
End Sub

Public Function ColumnLetterFromNumber(ByVal colNo As Long) As String
    ' 1 -> A、27 -> AA のように、列番号を Excel の列記号へ変換します。
    Dim n As Long
    Dim result As String

    If colNo <= 0 Then Exit Function

    n = colNo
    Do While n > 0
        result = Chr$(((n - 1) Mod 26) + Asc("A")) & result
        n = (n - 1) \ 26
    Loop

    ColumnLetterFromNumber = result
End Function

Public Function TryGetColumnNumber(ByVal ws As Worksheet, ByVal columnText As String, ByRef colNo As Long) As Boolean
    ' Range("列1") に直接渡す前に、列名だけを安全に検証します。
    Dim normalized As String
    Dim i As Long
    Dim charCode As Long
    Dim value As Long

    colNo = 0
    If ws Is Nothing Then Exit Function

    normalized = UCase$(Trim$(columnText))
    If Len(normalized) = 0 Then Exit Function

    For i = 1 To Len(normalized)
        charCode = AscW(Mid$(normalized, i, 1))
        If charCode < AscW("A") Or charCode > AscW("Z") Then Exit Function

        value = value * 26 + (charCode - AscW("A") + 1)
        If value > ws.Columns.Count Then Exit Function
    Next i

    colNo = value
    TryGetColumnNumber = (colNo > 0)
End Function

Public Function ShouldExitSkuLoop(ByVal vCheck As Variant) As Boolean
    ' F列が空になった商品ブロック以降は処理対象外と判断します。
    If IsError(vCheck) Then
        ShouldExitSkuLoop = True
    Else
        ShouldExitSkuLoop = (Len(Trim$(CStr(vCheck))) = 0)
    End If
End Function

Public Function WriteIfChanged(ByVal ws As Worksheet, ByVal wsFlag As Worksheet, ByVal r As Long, ByVal colLetter As String, ByVal newVal As Double, ByVal planRow As Long, ByVal stockRow As Long) As Boolean
    ' 値が実際に変わる場合だけ書き込み、セル色と変更フラグを更新します。
    Dim c As Long
    Dim oldVal As Variant
    Dim changed As Boolean

    If Not IsWritableRow(r, planRow, stockRow) Then Exit Function

    c = ws.Columns(colLetter).Column
    If ws.Cells(r, c).HasFormula Then Exit Function

    oldVal = ws.Cells(r, c).Value
    If IsError(oldVal) Then oldVal = Empty

    If IsNumeric(oldVal) Then
        changed = (Abs(CDbl(oldVal) - newVal) > EPS)
    Else
        changed = True
    End If

    If Not changed Then Exit Function

    ws.Cells(r, c).Value = newVal
    ws.Cells(r, c).Interior.Color = WRITTEN_COLOR()

    MarkChanged ws, wsFlag, r, c
    WriteIfChanged = True
End Function

Public Sub MarkChanged(ByVal ws As Worksheet, ByVal wsFlag As Worksheet, ByVal r As Long, ByVal c As Long)
    ' B列の変更印と change_flag_sheet のセル単位フラグをそろえて立てます。
    ws.Cells(r, CHANGE_ROW_MARK_COL).Value = "V"
    wsFlag.Cells(r, c).Value = "V"
End Sub

Public Function HasAnyFlagInRow(ByVal wsFlag As Worksheet, ByVal r As Long) As Boolean
    ' 指定行に V フラグが1つでも残っているか確認します。
    Dim lastCol As Long
    Dim c As Long
    Dim values As Variant

    lastCol = wsFlag.Cells(r, wsFlag.Columns.Count).End(xlToLeft).Column
    If lastCol = 1 Then
        HasAnyFlagInRow = (UCase$(Trim$(CStr(wsFlag.Cells(r, 1).Value))) = "V")
        Exit Function
    End If

    values = wsFlag.Range(wsFlag.Cells(r, 1), wsFlag.Cells(r, lastCol)).Value2

    For c = 1 To lastCol
        If UCase$(Trim$(CStr(values(1, c)))) = "V" Then
            HasAnyFlagInRow = True
            Exit Function
        End If
    Next c
End Function
