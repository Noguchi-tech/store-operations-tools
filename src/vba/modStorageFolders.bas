Attribute VB_Name = "modStorageFolders"
Option Private Module
Option Explicit

' =============================================================================
' 保存先フォルダ選択
' =============================================================================
' 調整後ブック・AM維持フラグ履歴ブックの保存先（いずれも Downloads 固定）と、
' フォルダピッカーの初期位置を扱います。

Public Function ResolveAdjustedWorkbookSaveFolder() As String
    ' 調整後ブックの保存先は、操作PCの Downloads フォルダに固定します。
    ' ブラウザからダウンロードして直接開いたブックは、元フォルダがキャッシュ等の
    ' 分かりにくい場所を指すため、元ブックと同じフォルダには保存しません。
    ' Downloads が見つからないPCに限り、フォルダ選択ダイアログへ切り替えます。
    Dim folderPath As String

    folderPath = GetDefaultDownloadFolder()
    If Len(folderPath) > 0 Then
        ResolveAdjustedWorkbookSaveFolder = folderPath
        Exit Function
    End If

    ResolveAdjustedWorkbookSaveFolder = PickAdjustedWorkbookSaveFolder()
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
    ' AM維持フラグ履歴ブックの保存先も、調整後ブックと同じく Downloads フォルダに固定します。
    ' 同名の履歴ブック（mmddAM維持フラグ商品履歴.xlsx）へ追記していく仕様のため、
    ' 保存先を毎回そろえてダイアログなしで確定させます。
    ' Downloads が見つからないPCに限り、フォルダ選択ダイアログへ切り替えます。
    ' 空文字を返した場合、呼び出し側（ExecuteKeepFaceFlagEvacuation）は調整全体を中止します。
    Dim folderPath As String

    folderPath = GetDefaultDownloadFolder()
    If Len(folderPath) > 0 Then
        ResolveKeepFaceHistoryFolder = folderPath
        Exit Function
    End If

    ResolveKeepFaceHistoryFolder = PickKeepFaceHistoryFolder()
End Function

Public Function PickKeepFaceHistoryFolder() As String
    ' Downloads を解決できないPC向けのフォールバックとして、
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
    ' 操作PCの Downloads フォルダを返します。調整後ブックの保存先と、
    ' フォルダピッカーの初期位置の両方で使います。
    ' ブラウザの保存先と一致させるため、Windows の既知フォルダ設定を最優先で
    ' 解決します（Downloads を別ドライブ等へ移動しているPCでも追従できます）。
    Dim folderPath As String

    On Error Resume Next
    folderPath = CreateObject("Shell.Application").NameSpace("shell:Downloads").Self.Path
    On Error GoTo 0
    folderPath = Trim$(folderPath)
    If Len(folderPath) > 0 Then
        If FolderExists(folderPath) Then
            GetDefaultDownloadFolder = folderPath
            Exit Function
        End If
    End If

    ' 既知フォルダを解決できない場合は、従来のユーザープロファイル配下を使います。
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
