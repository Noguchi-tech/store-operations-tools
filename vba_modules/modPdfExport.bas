Attribute VB_Name = "modPdfExport"
Option Private Module
Option Explicit

' =============================================================================
' 選択範囲のPDF出力
' =============================================================================
' 選択中のセル範囲を、A4片面（1ページ）に収めて PDF ファイルへ出力します。
' シートの印刷設定は一時的に変更し、出力後に元へ戻します。

Public Sub 選択範囲PDF出力()
    Const PROC_NAME As String = "選択範囲PDF出力"

    Dim targetRange As Range
    Dim ws As Worksheet
    Dim initialFolder As String
    Dim defaultName As String
    Dim savePath As Variant
    Dim oldPaperSize As XlPaperSize
    Dim oldZoom As Variant
    Dim oldFitWide As Variant
    Dim oldFitTall As Variant
    Dim pageSetupCaptured As Boolean
    Dim stepName As String

    On Error GoTo EH

    stepName = "選択範囲の確認"
    If ActiveWorkbook Is Nothing Then
        MsgBox "対象ブックが見つかりません。出力したいブックを前面にしてから実行してください。", vbExclamation
        Exit Sub
    End If
    If TypeName(Selection) <> "Range" Then
        MsgBox "セル範囲を選択してから実行してください。", vbExclamation
        Exit Sub
    End If

    Set targetRange = Selection
    If targetRange.Areas.Count > 1 Then
        MsgBox "複数の範囲が選択されています。" & vbCrLf & _
               "連続した1つの範囲を選択してから実行してください。", vbExclamation
        Exit Sub
    End If

    Set ws = targetRange.Worksheet

    ' 保存先はダイアログで指定します。初期位置は Downloads、初期名は日付とシート名です。
    stepName = "保存先の指定"
    defaultName = Format$(Date, "mmdd") & "_" & SanitizeFileNamePart(ws.Name) & "_選択範囲.pdf"
    initialFolder = GetDefaultDownloadFolder()
    If Len(initialFolder) > 0 Then defaultName = CombinePath(initialFolder, defaultName)

    savePath = Application.GetSaveAsFilename( _
        InitialFileName:=defaultName, _
        FileFilter:="PDFファイル (*.pdf), *.pdf", _
        Title:="選択範囲PDFの保存先を指定してください")
    If VarType(savePath) = vbBoolean Then Exit Sub
    If LCase$(Right$(CStr(savePath), 4)) <> ".pdf" Then savePath = CStr(savePath) & ".pdf"

    ' A4用紙1ページ（片面）に収まるよう、対象シートの印刷設定を一時変更します。
    stepName = "A4片面のページ設定"
    With ws.PageSetup
        oldPaperSize = .PaperSize
        oldZoom = .Zoom
        oldFitWide = .FitToPagesWide
        oldFitTall = .FitToPagesTall
        pageSetupCaptured = True

        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With

    stepName = "PDFの出力"
    targetRange.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=CStr(savePath), _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        OpenAfterPublish:=True

CLEANUP:
    ' 対象シートの印刷設定を実行前の状態へ戻します。
    If pageSetupCaptured Then
        On Error Resume Next
        With ws.PageSetup
            .PaperSize = oldPaperSize
            If VarType(oldZoom) = vbBoolean And oldZoom = False Then
                .Zoom = False
            Else
                .Zoom = oldZoom
            End If
            .FitToPagesWide = oldFitWide
            .FitToPagesTall = oldFitTall
        End With
        On Error GoTo 0
    End If
    Exit Sub

EH:
    ReportMacroError PROC_NAME, stepName, Err.Number, Err.Description, ActiveWorkbook, ws
    Resume CLEANUP
End Sub
