Attribute VB_Name = "modFileSave"
Option Private Module
Option Explicit

' =============================================================================
' ファイル保存・ファイル名ヘルパー
' =============================================================================
' 履歴ブック、調整後ブックの保存と、ファイル名の組み立てをまとめます。

Public Function GetTextAfterColon(ByVal v As Variant) As String
    ' 「コード:名称」のような文字列から、名称部分を取り出します。
    ' 半角・全角コロンの早い方で区切り、コロンが無い場合は全体を返します。
    ' 調整後ブック名の商品名部分と、編集者名の既定値の両方で使う共通処理です。
    Dim s As String
    Dim p1 As Long
    Dim p2 As Long
    Dim p As Long

    If IsError(v) Then Exit Function

    s = Trim$(CStr(v))
    If Len(s) = 0 Then Exit Function

    p1 = InStr(1, s, ":", vbTextCompare)
    p2 = InStr(1, s, "：", vbTextCompare)

    If p1 > 0 And p2 > 0 Then
        If p1 < p2 Then p = p1 Else p = p2
    ElseIf p1 > 0 Then
        p = p1
    Else
        p = p2
    End If

    If p > 0 Then s = Mid$(s, p + 1)

    GetTextAfterColon = Trim$(s)
End Function

Public Function GetKeepFaceHistoryNamePart(ByVal v As Variant) As String
    ' D列の「コード:名称」のような文字列から、保存名に使う名称部分を取り出します。
    ' 調整後ブック名では、コロン右側の先頭4文字だけを使います。
    Dim s As String

    s = GetTextAfterColon(v)
    If Len(s) > 4 Then s = Left$(s, 4)

    GetKeepFaceHistoryNamePart = s
End Function

Public Function BuildKeepFaceHistoryFilePath(ByVal saveFolder As String) As String
    ' 日付・固定 suffix から履歴ブック名を作ります。同名ファイルがあれば追記します。
    BuildKeepFaceHistoryFilePath = CombinePath(saveFolder, Format$(Date, "mmdd") & KEEP_FACE_HISTORY_FILE_SUFFIX & ".xlsx")
End Function

Public Function SanitizeFileNamePart(ByVal s As String) As String
    ' Windows のファイル名に使えない文字を置換し、末尾の空白やピリオドを削ります。
    Dim badChars As Variant
    Dim i As Long

    s = Trim$(s)
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(badChars) To UBound(badChars)
        s = Replace$(s, CStr(badChars(i)), "_")
    Next i

    Do While Len(s) > 0 And (Right$(s, 1) = "." Or Right$(s, 1) = " ")
        s = Left$(s, Len(s) - 1)
    Loop

    SanitizeFileNamePart = s
End Function

Public Function CombinePath(ByVal folderPath As String, ByVal itemName As String) As String
    ' フォルダ末尾の \ 有無に関係なく、正しいフルパスを組み立てます。
    ' 引数名を fileName にすると、VBE が SaveAs の名前付き引数 Filename:= を
    ' fileName:= へ自動リケースしてしまうため、itemName という名前にしています。
    If Right$(folderPath, 1) = "\" Or Right$(folderPath, 1) = "/" Then
        CombinePath = folderPath & itemName
    Else
        CombinePath = folderPath & "\" & itemName
    End If
End Function

Public Function FolderExists(ByVal folderPath As String) As Boolean
    ' ローカルフォルダだけを対象に存在確認します。URL 形式のパスは対象外です。
    If Len(folderPath) = 0 Then Exit Function
    If InStr(1, folderPath, "://", vbTextCompare) > 0 Then Exit Function

    On Error Resume Next
    FolderExists = (Len(Dir$(folderPath, vbDirectory)) > 0)
    On Error GoTo 0
End Function

Private Function NormalizeHistoryJanKey(ByVal v As Variant) As String
    ' JAN の比較用キーです。数値化で指数表記になった場合だけ、整数文字列へ戻します。
    Dim s As String

    If IsError(v) Then Exit Function

    s = Trim$(CStr(v))
    If Len(s) = 0 Then Exit Function

    If IsNumeric(v) Then
        If InStr(1, s, "E", vbTextCompare) > 0 Or InStr(1, s, ".") > 0 Then
            s = Format$(CDbl(v), "0")
        End If
    End If

    NormalizeHistoryJanKey = s
End Function

Public Sub SaveKeepFaceHistoryWorkbook(ByRef janValues() As String, ByVal itemCount As Long, ByVal fullPath As String, ByRef appendedCount As Long, ByRef skippedDuplicateCount As Long, ByRef deletedDuplicateCount As Long)
    ' JAN のみを1列にした xlsx を作成または開き、重複を除いて追記します。
    Dim histWb As Workbook
    Dim histWs As Worksheet
    Dim i As Long
    Dim r As Long
    Dim lastRow As Long
    Dim nextRow As Long
    Dim oldDisplayAlerts As Boolean
    Dim displayAlertsCaptured As Boolean
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    Dim existingKeys As Object
    Dim newKeys As Object
    Dim janKey As String
    Dim duplicateRows() As Long
    Dim duplicateRowCount As Long
    Dim changedWorkbook As Boolean

    On Error GoTo EH

    appendedCount = 0
    skippedDuplicateCount = 0
    deletedDuplicateCount = 0
    Set existingKeys = CreateObject("Scripting.Dictionary")
    existingKeys.CompareMode = vbTextCompare
    Set newKeys = CreateObject("Scripting.Dictionary")
    newKeys.CompareMode = vbTextCompare

    If Len(Dir$(fullPath)) > 0 Then
        Set histWb = Application.Workbooks.Open(Filename:=fullPath, UpdateLinks:=0, ReadOnly:=False)
        Set histWs = histWb.Worksheets(1)
    Else
        Set histWb = Application.Workbooks.Add(xlWBATWorksheet)
        Set histWs = histWb.Worksheets(1)
    End If

    histWs.Columns(1).NumberFormat = "@"

    lastRow = histWs.Cells(histWs.Rows.Count, 1).End(xlUp).Row
    If Not (lastRow = 1 And Len(Trim$(CStr(histWs.Cells(1, 1).Value))) = 0) Then
        For r = 1 To lastRow
            janKey = NormalizeHistoryJanKey(histWs.Cells(r, 1).Value2)
            If Len(janKey) > 0 Then
                If existingKeys.Exists(janKey) Then
                    duplicateRowCount = duplicateRowCount + 1
                    ReDim Preserve duplicateRows(1 To duplicateRowCount)
                    duplicateRows(duplicateRowCount) = r
                Else
                    existingKeys.Add janKey, True
                End If
            End If
        Next r
    End If

    For i = duplicateRowCount To 1 Step -1
        histWs.Rows(duplicateRows(i)).Delete
        deletedDuplicateCount = deletedDuplicateCount + 1
        changedWorkbook = True
    Next i

    nextRow = histWs.Cells(histWs.Rows.Count, 1).End(xlUp).Row
    If nextRow = 1 And Len(Trim$(CStr(histWs.Cells(1, 1).Value))) = 0 Then
        nextRow = 1
    Else
        nextRow = nextRow + 1
    End If

    For i = 1 To itemCount
        janKey = NormalizeHistoryJanKey(janValues(i))
        If Len(janKey) > 0 And (existingKeys.Exists(janKey) Or newKeys.Exists(janKey)) Then
            skippedDuplicateCount = skippedDuplicateCount + 1
        Else
            histWs.Cells(nextRow + appendedCount, 1).Value = janValues(i)
            appendedCount = appendedCount + 1
            changedWorkbook = True
            If Len(janKey) > 0 Then newKeys.Add janKey, True
        End If
    Next i
    histWs.Columns(1).AutoFit

    If changedWorkbook Then
        oldDisplayAlerts = Application.DisplayAlerts
        displayAlertsCaptured = True
        Application.DisplayAlerts = False
        If Len(Dir$(fullPath)) > 0 Then
            histWb.Save
        Else
            histWb.SaveAs Filename:=fullPath, FileFormat:=xlOpenXMLWorkbook
        End If
        Application.DisplayAlerts = oldDisplayAlerts
    End If

    histWb.Close SaveChanges:=False
    Exit Sub

EH:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description

    On Error Resume Next
    If displayAlertsCaptured Then Application.DisplayAlerts = oldDisplayAlerts
    If Not histWb Is Nothing Then histWb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise errNumber, errSource, errDescription
End Sub

Public Function GetListNamePartForAdjustedWorkbook(ByVal ws As Worksheet) As String
    ' 調整後ファイル名に入れる商品名部分を、最初に見つかる商品ブロックから取得します。
    Dim lastInputRow As Long
    Dim planRow As Long
    Dim stockRow As Long
    Dim namePart As String

    If ws Is Nothing Then Exit Function

    lastInputRow = ws.Cells(ws.Rows.Count, ITEM_CHECK_COL).End(xlUp).Row

    For planRow = START_ROW To lastInputRow Step STEP_ROW
        If ShouldExitSkuLoop(ws.Cells(planRow, ITEM_CHECK_COL).Value) Then Exit For

        stockRow = planRow + STOCKROW_OFFSET

        namePart = GetKeepFaceHistoryNamePart(ws.Cells(stockRow, "D").Value)
        If Len(namePart) = 0 Then namePart = GetKeepFaceHistoryNamePart(ws.Cells(planRow, "D").Value)
        If Len(namePart) > 0 Then
            GetListNamePartForAdjustedWorkbook = namePart
            Exit Function
        End If
    Next planRow
End Function

Public Function BuildAdjustedWorkbookBaseName(ByVal ws As Worksheet, ByVal modeName As String, ByVal targetWeeks As Double, ByVal minVal As Long, ByVal includeIntentioned As Boolean, ByVal operatorName As String, Optional ByVal appendTwoWeekSuffix As Boolean = False) As String
    ' 日付・商品名・操作者・調整条件を組み合わせ、保存時のベース名を作ります。
    Dim namePart As String
    Dim safeNamePart As String
    Dim safeOperatorName As String
    Dim baseName As String

    namePart = GetListNamePartForAdjustedWorkbook(ws)
    safeNamePart = SanitizeFileNamePart(namePart)
    If Len(safeNamePart) = 0 Then safeNamePart = "未設定"

    safeOperatorName = SanitizeFileNamePart(operatorName)
    If Len(safeOperatorName) = 0 Then safeOperatorName = "未入力"

    baseName = Format$(Date, "mmdd") & safeNamePart & safeOperatorName

    baseName = baseName & "_" & FormatTargetWeeksForFileName(targetWeeks) & "週_下限値" & CStr(minVal)
    If Not includeIntentioned Then baseName = baseName & ADJUSTED_BOOK_EXCLUDE_INTENTIONED_SUFFIX
    If appendTwoWeekSuffix Then baseName = baseName & TwoWeekFileNameSuffix()

    BuildAdjustedWorkbookBaseName = baseName
End Function

Public Function FormatTargetWeeksForFileName(ByVal targetWeeks As Double) As String
    ' ファイル名内の手持週数表記を 1 桁小数へそろえます。
    FormatTargetWeeksForFileName = Format$(targetWeeks, "0.0")
End Function

Private Function TwoWeekFileNameSuffix() As String
    TwoWeekFileNameSuffix = "_2" & ChrW$(&H9031) & ChrW$(&H9593) & ChrW$(&H5206)
End Function

Public Function GetWorkbookExtension(ByVal wb As Workbook) As String
    ' 既存ブックの拡張子を維持して保存できるよう、名前または FileFormat から推定します。
    Dim ext As String
    Dim p As Long

    If wb Is Nothing Then Exit Function

    p = InStrRev(wb.Name, ".")
    If p > 0 Then ext = Mid$(wb.Name, p)

    If Len(ext) = 0 Then
        Select Case wb.FileFormat
            Case xlOpenXMLWorkbookMacroEnabled
                ext = ".xlsm"
            Case xlExcel12
                ext = ".xlsb"
            Case xlExcel8
                ext = ".xls"
            Case Else
                ext = ".xlsx"
        End Select
    End If

    GetWorkbookExtension = ext
End Function

Public Function SaveAdjustedWorkbookWithName(ByVal wb As Workbook, ByVal ws As Worksheet, ByVal modeName As String, ByVal targetWeeks As Double, ByVal minVal As Long, ByVal includeIntentioned As Boolean, ByVal operatorName As String, ByRef saveNote As String, Optional ByVal appendTwoWeekSuffix As Boolean = False) As String
    ' 調整した対象ブックを条件が分かる名前で保存し、保存成功後に元ファイルを削除します。
    Dim saveFolder As String
    Dim baseName As String
    Dim extensionText As String
    Dim firstChoicePath As String
    Dim fullPath As String
    Dim oldFullPath As String
    Dim oldDisplayAlerts As Boolean
    Dim duplicateDeleteNote As String
    Dim originalDeleteNote As String
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    saveNote = "対象なし"
    If wb Is Nothing Or ws Is Nothing Then Exit Function

    If Len(wb.Path) > 0 And InStr(1, wb.FullName, "://", vbTextCompare) = 0 Then
        oldFullPath = wb.FullName
    End If

    ' 保存先は常に操作PCの Downloads フォルダです（modStorageFolders 参照）。
    ' 元ブックと同じフォルダへは保存しません。ブラウザから直接開いたブックだと
    ' 元フォルダがキャッシュ等の分かりにくい場所になり、MDシステムへの取込時に
    ' ファイルを見失うためです。
    saveFolder = ResolveAdjustedWorkbookSaveFolder()
    If Len(saveFolder) = 0 Then
        Err.Raise vbObjectError + 5201, "SaveAdjustedWorkbookWithName", "調整後ブックの保存先が選択されなかったため、保存を中止しました。"
    End If

    baseName = BuildAdjustedWorkbookBaseName(ws, modeName, targetWeeks, minVal, includeIntentioned, operatorName, appendTwoWeekSuffix)
    extensionText = GetWorkbookExtension(wb)
    If Len(extensionText) > 0 And Left$(extensionText, 1) <> "." Then extensionText = "." & extensionText

    firstChoicePath = CombinePath(saveFolder, baseName & extensionText)
    fullPath = firstChoicePath
    duplicateDeleteNote = DeleteExistingAdjustedWorkbookIfNeeded(fullPath, oldFullPath)

    On Error GoTo EH
    oldDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=fullPath, FileFormat:=wb.FileFormat
    Application.DisplayAlerts = oldDisplayAlerts

    saveNote = duplicateDeleteNote
    If Len(oldFullPath) > 0 And StrComp(oldFullPath, fullPath, vbTextCompare) <> 0 Then
        ' 元ファイル削除は「Downloads・一時フォルダ・ブラウザキャッシュ」にある場合だけです。
        ' 共有フォルダやデスクトップ等に置かれた原本まで消さないための安全弁です。
        If IsSafeOriginalDeleteLocation(oldFullPath) Then
            originalDeleteNote = DeleteOriginalWorkbookFile(oldFullPath)
        Else
            originalDeleteNote = "保持（Downloads・一時フォルダ以外のため削除しません）：" & oldFullPath
        End If
        If Len(saveNote) > 0 And Len(originalDeleteNote) > 0 Then
            saveNote = saveNote & " / " & originalDeleteNote
        ElseIf Len(originalDeleteNote) > 0 Then
            saveNote = originalDeleteNote
        End If
    End If
    If Len(saveNote) = 0 Then saveNote = "対象なし"

    SaveAdjustedWorkbookWithName = fullPath
    Exit Function

EH:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    On Error Resume Next
    Application.DisplayAlerts = oldDisplayAlerts
    On Error GoTo 0
    Err.Raise errNumber, errSource, errDescription
End Function

Private Function DeleteExistingAdjustedWorkbookIfNeeded(ByVal adjustedPath As String, ByVal oldFullPath As String) As String
    If Len(adjustedPath) = 0 Then Exit Function
    If InStr(1, adjustedPath, "://", vbTextCompare) > 0 Then Exit Function
    If Len(oldFullPath) > 0 Then
        If StrComp(adjustedPath, oldFullPath, vbTextCompare) = 0 Then Exit Function
    End If
    If Len(Dir$(adjustedPath)) = 0 Then Exit Function

    On Error GoTo EH
    SetAttr adjustedPath, vbNormal
    Kill adjustedPath
    DeleteExistingAdjustedWorkbookIfNeeded = "同名旧ファイル削除済み：" & adjustedPath
    Exit Function

EH:
    Err.Raise vbObjectError + 5202, "DeleteExistingAdjustedWorkbookIfNeeded", "同名の旧ファイルを削除できませんでした：" & Err.Description & "（" & adjustedPath & "）"
End Function

Public Function IsSafeOriginalDeleteLocation(ByVal oldFullPath As String) As Boolean
    ' 元ファイルを削除してよい場所かを判定します。
    ' 保存先が Downloads 固定になったため、元ファイルは任意の場所にあり得ます。
    ' 削除してよいのは「置きっぱなしでも困らない場所」だけに限定します。
    '   - Downloads フォルダ配下（ブラウザの通常ダウンロード先）
    '   - TEMP / TMP 配下
    '   - ブラウザキャッシュ（INetCache / Temporary Internet Files）
    ' 共有フォルダ（UNC）やそれ以外のフォルダの原本は削除せず保持します。
    Dim folderPath As String
    Dim candidates(1 To 3) As String
    Dim i As Long

    If Len(oldFullPath) = 0 Then Exit Function
    If InStr(1, oldFullPath, "://", vbTextCompare) > 0 Then Exit Function
    If Left$(oldFullPath, 2) = "\\" Then Exit Function

    If InStr(1, oldFullPath, "\INetCache\", vbTextCompare) > 0 Then
        IsSafeOriginalDeleteLocation = True
        Exit Function
    End If
    If InStr(1, oldFullPath, "\Temporary Internet Files\", vbTextCompare) > 0 Then
        IsSafeOriginalDeleteLocation = True
        Exit Function
    End If

    candidates(1) = GetDefaultDownloadFolder()
    candidates(2) = Trim$(Environ$("TEMP"))
    candidates(3) = Trim$(Environ$("TMP"))

    For i = 1 To 3
        folderPath = candidates(i)
        If Len(folderPath) > 0 Then
            If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
            If StrComp(Left$(oldFullPath, Len(folderPath)), folderPath, vbTextCompare) = 0 Then
                IsSafeOriginalDeleteLocation = True
                Exit Function
            End If
        End If
    Next i
End Function

Public Function DeleteOriginalWorkbookFile(ByVal oldFullPath As String) As String
    ' SaveAs 成功後にだけ呼び、フォルダに旧ブックと新ブックが混在しないようにします。
    If Len(oldFullPath) = 0 Then Exit Function
    If InStr(1, oldFullPath, "://", vbTextCompare) > 0 Then Exit Function

    On Error GoTo EH
    If Len(Dir$(oldFullPath)) = 0 Then
        DeleteOriginalWorkbookFile = "対象なし（元ファイルなし）"
        Exit Function
    End If

    SetAttr oldFullPath, vbNormal
    Kill oldFullPath
    DeleteOriginalWorkbookFile = "実行済み：" & oldFullPath
    Exit Function

EH:
    DeleteOriginalWorkbookFile = "失敗：" & Err.Description & "（" & oldFullPath & "）"
End Function
