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

Public Function PromptOperatorName(ByVal modeName As String) As String
    ' 調整後ファイル名に入れるあなたの名前を受け取り、空欄なら再入力を促します。
    Dim inputText As String

    Do
        inputText = InputBox( _
            modeName & "後のファイル名に入れるあなたの名前を入力してください。", _
            "あなたの名前の入力", _
            Application.UserName)

        If StrPtr(inputText) = 0 Then Exit Function

        inputText = Trim$(inputText)
        If Len(inputText) > 0 Then
            PromptOperatorName = inputText
            Exit Function
        End If

        MsgBox "あなたの名前を入力してください。", vbExclamation, "入力エラー"
    Loop
End Function
