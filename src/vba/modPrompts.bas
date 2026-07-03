Attribute VB_Name = "modPrompts"
Option Private Module
Option Explicit

' =============================================================================
' 入力・確認ダイアログ
' =============================================================================
' 操作者に判断してもらう文言をここへ集約し、処理本体から MsgBox / InputBox の文言を分離します。

Public Function ConfirmAdjustmentScope(ByVal modeName As String, ByRef includeIntentioned As Boolean, ByRef scopeLabel As String) As Boolean
    ' 自動調整・手動調整で共通の修正範囲を確認します。
    ' はいは全て修正、いいえは未意思入れ商品のみ修正です。
    Dim ans As VbMsgBoxResult

    ans = MsgBox( _
        "修正範囲を選択してください。" & vbCrLf & vbCrLf & _
        "はい：全て修正" & vbCrLf & _
        "いいえ：未意思入れ商品のみ修正" & vbCrLf & _
        "キャンセル：中止", _
        vbYesNoCancel + vbQuestion + vbDefaultButton1, _
        "在庫コントロール " & modeName)

    If ans = vbCancel Then Exit Function

    includeIntentioned = (ans = vbYes)
    If includeIntentioned Then
        scopeLabel = "全て修正"
    Else
        scopeLabel = "未意思入れ商品のみ修正"
    End If

    ConfirmAdjustmentScope = True
End Function

Public Function PromptOperatorName(ByVal modeName As String, ByVal ws As Worksheet) As String
    ' 調整後ファイル名に入れるあなたの名前を受け取り、空欄なら再入力を促します。
    ' 既定値は list シートのデパ名から作る「◯◯◯担当」（例「ファニ担当」）です。
    ' デパ名から作れない場合だけ、Excel のユーザー名を既定値にします。
    Dim defaultName As String
    Dim inputText As String

    defaultName = BuildDefaultOperatorName(ws)
    If Len(defaultName) = 0 Then defaultName = Application.UserName

    Do
        inputText = InputBox( _
            modeName & "後のファイル名に入れるあなたの名前を入力してください。", _
            "あなたの名前の入力", _
            defaultName)

        If StrPtr(inputText) = 0 Then Exit Function

        inputText = Trim$(inputText)
        If Len(inputText) > 0 Then
            PromptOperatorName = inputText
            Exit Function
        End If

        MsgBox "あなたの名前を入力してください。", vbExclamation, "入力エラー"
    Loop
End Function

Public Function BuildDefaultOperatorName(ByVal ws As Worksheet) As String
    ' list シートの C6 以下にあるデパ名から、編集者名の既定値を作ります。
    ' 例「110:ファニチャー」→ コロン右側の先頭3文字＋担当 →「ファニ担当」
    ' C6 から下へ見て、コロンを含む最初のセルだけをデパ名として使います。
    ' コロンの無いメモや数値をうっかり既定値へ採用しないための条件です。
    ' デパ名が見つからない場合は空文字を返し、呼び出し側でユーザー名へフォールバックします。
    Dim lastRow As Long
    Dim r As Long
    Dim cellValue As Variant
    Dim cellText As String
    Dim namePart As String

    If ws Is Nothing Then Exit Function

    lastRow = ws.Cells(ws.Rows.Count, OPERATOR_DEPT_COL).End(xlUp).Row

    For r = OPERATOR_DEPT_START_ROW To lastRow
        cellValue = ws.Cells(r, OPERATOR_DEPT_COL).Value
        If Not IsError(cellValue) Then
            cellText = Trim$(CStr(cellValue))
            If InStr(1, cellText, ":", vbTextCompare) > 0 Or InStr(1, cellText, "：", vbTextCompare) > 0 Then
                namePart = GetTextAfterColon(cellText)
                If Len(namePart) > 0 Then
                    If Len(namePart) > OPERATOR_NAME_CHAR_COUNT Then
                        namePart = Left$(namePart, OPERATOR_NAME_CHAR_COUNT)
                    End If
                    BuildDefaultOperatorName = namePart & OPERATOR_NAME_SUFFIX
                    Exit Function
                End If
            End If
        End If
    Next r
End Function
