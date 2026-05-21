Attribute VB_Name = "modDebug"
Option Private Module
Option Explicit

' =============================================================================
' デバッグ・エラー報告
' =============================================================================
' マクロ実行中にエラーが起きたとき、後任者が「どの処理・どの行・どの列」で
' 止まったかを追えるようにするための共通モジュールです。
'
' エラー時は以下の3か所へ同じ情報を出します。
'   1. 画面のメッセージボックス
'   2. VBE の Immediate ウィンドウ（Debug.Print）
'   3. 実行対象ブック内の debug_log シート
'
' debug_log シートは、エラー発生時に存在しなければ自動作成します。
' ブック保護などで作成できない場合でも、マクロ本体のエラー処理を止めないよう、
' このモジュール内のログ書き込みエラーは握りつぶします。

Private Const DEBUG_LOG_SHEET_NAME As String = "debug_log"

Public Sub ReportMacroError( _
    ByVal procName As String, _
    ByVal stepName As String, _
    ByVal errNumber As Long, _
    ByVal errDescription As String, _
    Optional ByVal wb As Workbook = Nothing, _
    Optional ByVal ws As Worksheet = Nothing, _
    Optional ByVal planRow As Long = 0, _
    Optional ByVal stockRow As Long = 0, _
    Optional ByVal colNo As Long = 0)

    ' 各マクロから渡された現在位置を、人が読めるメッセージに整形します。
    Dim locationText As String
    Dim msg As String

    locationText = BuildLocationText(wb, ws, planRow, stockRow, colNo)

    msg = "エラーが発生しました。" & vbCrLf & _
          "処理: " & procName & vbCrLf & _
          "段階: " & IIf(Len(stepName) > 0, stepName, "(未設定)") & vbCrLf & _
          locationText & _
          "Error " & errNumber & ": " & errDescription

    Debug.Print Format$(Now, "yyyy/mm/dd hh:nn:ss"); " "; Replace(msg, vbCrLf, " / ")
    AppendDebugLog wb, procName, stepName, errNumber, errDescription, wb, ws, planRow, stockRow, colNo

    MsgBox msg, vbExclamation, "在庫自動調整 エラー"
End Sub

Public Sub デバッグログ動作確認()
    ' 後任者向けの確認用です。
    ' VBE の Immediate ウィンドウから「デバッグログ動作確認」と実行すると、
    ' ActiveWorkbook に debug_log シートを作れるか確認できます。
    On Error GoTo EH

    Dim wb As Workbook

    Set wb = ActiveWorkbook
    AppendDebugLog wb, "デバッグログ動作確認", "ログ書き込みテスト", 0, "テスト行です。", wb
    MsgBox "debug_log シートへの書き込み確認が完了しました。", vbInformation
    Exit Sub

EH:
    MsgBox "デバッグログ動作確認 Error " & Err.Number & vbCrLf & Err.Description, vbExclamation
End Sub

Public Function DebugCompileSmokeTest() As Boolean
    ' Excel COM からの検証用です。画面表示やシート更新をせず、
    ' VBA プロジェクト全体がコンパイル可能かを確認するときに呼びます。
    DebugCompileSmokeTest = (ColumnLetterFromNumber(26) = "Z")
End Function

Private Function BuildLocationText( _
    Optional ByVal wb As Workbook = Nothing, _
    Optional ByVal ws As Worksheet = Nothing, _
    Optional ByVal planRow As Long = 0, _
    Optional ByVal stockRow As Long = 0, _
    Optional ByVal colNo As Long = 0) As String

    ' 未指定の情報は出さず、分かっている範囲だけをログに載せます。
    Dim text As String

    If Not wb Is Nothing Then text = text & "ブック: " & wb.Name & vbCrLf
    If Not ws Is Nothing Then text = text & "シート: " & ws.Name & vbCrLf
    If planRow > 0 Then text = text & "販売計画行: " & CStr(planRow) & vbCrLf
    If stockRow > 0 Then text = text & "フェイス陳列数行: " & CStr(stockRow) & vbCrLf
    If colNo > 0 Then text = text & "列: " & ColumnLetterFromNumber(colNo) & " (" & CStr(colNo) & ")" & vbCrLf

    BuildLocationText = text
End Function

Private Sub AppendDebugLog( _
    Optional ByVal logWb As Workbook = Nothing, _
    Optional ByVal procName As String = "", _
    Optional ByVal stepName As String = "", _
    Optional ByVal errNumber As Long = 0, _
    Optional ByVal errDescription As String = "", _
    Optional ByVal targetWb As Workbook = Nothing, _
    Optional ByVal targetWs As Worksheet = Nothing, _
    Optional ByVal planRow As Long = 0, _
    Optional ByVal stockRow As Long = 0, _
    Optional ByVal colNo As Long = 0)

    ' ログ出力の失敗で本来のエラー処理を止めないよう、この中だけは Resume Next にします。
    On Error Resume Next

    Dim wsLog As Worksheet
    Dim nextRow As Long
    Dim colText As String

    If logWb Is Nothing Then Set logWb = ActiveWorkbook
    If logWb Is Nothing Then Exit Sub

    ' debug_log がなければ末尾に作成し、同じ列構成で追記できるようにします。
    Set wsLog = GetWorksheetOrNothing(logWb, DEBUG_LOG_SHEET_NAME)
    If wsLog Is Nothing Then
        Set wsLog = logWb.Worksheets.Add(After:=logWb.Worksheets(logWb.Worksheets.Count))
        wsLog.Name = DEBUG_LOG_SHEET_NAME
        wsLog.Range("A1:J1").Value = Array("日時", "処理", "段階", "対象ブック", "対象シート", "販売計画行", "フェイス行", "列", "Err.Number", "Err.Description")
        wsLog.Rows(1).Font.Bold = True
    End If

    nextRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
    If colNo > 0 Then colText = ColumnLetterFromNumber(colNo) & " (" & CStr(colNo) & ")"

    wsLog.Cells(nextRow, 1).Value = Now
    wsLog.Cells(nextRow, 2).Value = procName
    wsLog.Cells(nextRow, 3).Value = stepName
    If Not targetWb Is Nothing Then wsLog.Cells(nextRow, 4).Value = targetWb.Name
    If Not targetWs Is Nothing Then wsLog.Cells(nextRow, 5).Value = targetWs.Name
    If planRow > 0 Then wsLog.Cells(nextRow, 6).Value = planRow
    If stockRow > 0 Then wsLog.Cells(nextRow, 7).Value = stockRow
    wsLog.Cells(nextRow, 8).Value = colText
    wsLog.Cells(nextRow, 9).Value = errNumber
    wsLog.Cells(nextRow, 10).Value = errDescription
    wsLog.Columns("A:J").AutoFit

    On Error GoTo 0
End Sub
