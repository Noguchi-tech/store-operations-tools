Attribute VB_Name = "modStorageFolders"
Option Private Module
Option Explicit

' =============================================================================
' 保存先フォルダ選択・記憶
' =============================================================================
' 操作PCごとに変わる保存先の選択、前回保存先の再利用、フォルダピッカーの初期位置を扱います。

Public Function GetWorkbookSaveFolder(ByVal wb As Workbook) As String
    ' 元ブックがローカルに保存済みなら同じフォルダ、未保存やURLなら選択ダイアログを使います。
    If wb Is Nothing Then Exit Function
    If Len(wb.Path) > 0 And InStr(1, wb.Path, "://", vbTextCompare) = 0 Then
        GetWorkbookSaveFolder = wb.Path
        Exit Function
    End If

    GetWorkbookSaveFolder = PickAdjustedWorkbookSaveFolder()
End Function

Public Function PickAdjustedWorkbookSaveFolder() As String
    ' 調整後ブックの保存は必須です。キャンセルされた場合は選び直しを促します。
    Dim fd As FileDialog
    Dim ans As VbMsgBoxResult

    Do
        Set fd = Application.FileDialog(msoFileDialogFolderPicker)
        With fd
            .Title = "調整後ブックの保存先フォルダを選択してください"
            .AllowMultiSelect = False
            SetFolderPickerInitialPath fd
            If .Show = -1 Then
                PickAdjustedWorkbookSaveFolder = CStr(.SelectedItems(1))
                Exit Function
            End If
        End With

        ans = MsgBox( _
            "調整後ブックの保存は必須です。" & vbCrLf & _
            "保存先フォルダを選び直しますか？" & vbCrLf & vbCrLf & _
            "再試行：保存先を選ぶ" & vbCrLf & _
            "キャンセル：処理を中止", _
            vbRetryCancel + vbExclamation + vbDefaultButton1, _
            "調整後ブック保存先")
        If ans = vbCancel Then Exit Function
    Loop
End Function

Public Function ResolveKeepFaceHistoryFolder() As String
    ' 前回保存先が使える場合は再利用し、使えない場合は選び直して設定へ保存します。
    Dim savedFolder As String
    Dim ans As VbMsgBoxResult
    Dim pickedFolder As String

    savedFolder = ReadAddInSetting(KEEP_FACE_HISTORY_SETTING_KEY)
    If Len(savedFolder) > 0 And FolderExists(savedFolder) Then
        ans = MsgBox( _
            "AM維持フラグ履歴ブックの保存先は現在こちらです。" & vbCrLf & vbCrLf & _
            savedFolder & vbCrLf & vbCrLf & _
            "この保存先を使いますか？" & vbCrLf & _
            "はい：この保存先を使う" & vbCrLf & _
            "いいえ：保存先を選び直す" & vbCrLf & _
            "キャンセル：処理を中止", _
            vbYesNoCancel + vbQuestion, _
            "履歴ブック保存先")

        If ans = vbYes Then
            ResolveKeepFaceHistoryFolder = savedFolder
            Exit Function
        ElseIf ans = vbCancel Then
            Exit Function
        End If
    ElseIf Len(savedFolder) > 0 Then
        MsgBox "前回の保存先が見つかりません。" & vbCrLf & _
               savedFolder & vbCrLf & _
               "保存先を選び直してください。", vbInformation, "履歴ブック保存先"
    End If

    pickedFolder = PickKeepFaceHistoryFolder()
    If Len(pickedFolder) = 0 Then Exit Function

    WriteAddInSetting KEEP_FACE_HISTORY_SETTING_KEY, pickedFolder
    ResolveKeepFaceHistoryFolder = pickedFolder
End Function

Public Function PickKeepFaceHistoryFolder() As String
    ' AM維持フラグの履歴ブックを保存するフォルダをユーザーに選択してもらいます。
    Dim fd As FileDialog

    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "AM維持フラグ商品履歴ブックの保存先フォルダを選択してください"
        .AllowMultiSelect = False
        SetFolderPickerInitialPath fd
        If .Show <> -1 Then Exit Function
        PickKeepFaceHistoryFolder = CStr(.SelectedItems(1))
    End With
End Function

Public Function GetDefaultDownloadFolder() As String
    ' 保存先選択の初期位置は、操作PCの Downloads フォルダを優先します。
    Dim folderPath As String

    folderPath = Trim$(Environ$("USERPROFILE"))
    If Len(folderPath) > 0 Then
        folderPath = CombinePath(folderPath, "Downloads")
        If FolderExists(folderPath) Then
            GetDefaultDownloadFolder = folderPath
            Exit Function
        End If
    End If

    folderPath = Trim$(Environ$("HOMEDRIVE") & Environ$("HOMEPATH"))
    If Len(folderPath) > 0 Then
        folderPath = CombinePath(folderPath, "Downloads")
        If FolderExists(folderPath) Then GetDefaultDownloadFolder = folderPath
    End If
End Function

Private Sub SetFolderPickerInitialPath(ByVal fd As FileDialog)
    ' FileDialog は存在しない初期パスでエラーになるため、使える場合だけ設定します。
    Dim defaultFolder As String

    defaultFolder = GetDefaultDownloadFolder()
    If Len(defaultFolder) = 0 Then Exit Sub

    On Error Resume Next
    fd.InitialFileName = CombinePath(defaultFolder, "")
    On Error GoTo 0
End Sub
