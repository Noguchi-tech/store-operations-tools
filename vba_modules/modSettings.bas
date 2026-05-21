Attribute VB_Name = "modSettings"
Option Private Module
Option Explicit

' =============================================================================
' アドイン設定保存
' =============================================================================
' アドイン内の setting シートへ前回保存先などの設定を保存・取得します。

Public Function ReadAddInSetting(ByVal settingKey As String) As String
    ' アドイン内の setting シートから、指定キーの保存値を読み取ります。
    Dim wsSetting As Worksheet
    Dim foundCell As Range

    Set wsSetting = GetAddInSettingWorksheet(False)
    If wsSetting Is Nothing Then Exit Function

    Set foundCell = wsSetting.Columns(1).Find( _
        What:=settingKey, _
        After:=wsSetting.Cells(wsSetting.Rows.Count, 1), _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False, _
        SearchFormat:=False)
    If foundCell Is Nothing Then Exit Function

    ReadAddInSetting = Trim$(CStr(foundCell.Offset(0, 1).Value))
End Function

Public Sub WriteAddInSetting(ByVal settingKey As String, ByVal settingValue As String)
    ' アドイン内の setting シートへキーと値を保存します。保存先などの前回設定に使います。
    Const PROC_NAME As String = "WriteAddInSetting"

    Dim wsSetting As Worksheet
    Dim foundCell As Range
    Dim nextRow As Long

    On Error GoTo EH

    Set wsSetting = GetAddInSettingWorksheet(True)
    If wsSetting Is Nothing Then
        Err.Raise vbObjectError + 5101, PROC_NAME, "setting シートを作成または取得できませんでした。"
    End If

    Set foundCell = wsSetting.Columns(1).Find( _
        What:=settingKey, _
        After:=wsSetting.Cells(wsSetting.Rows.Count, 1), _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False, _
        SearchFormat:=False)
    If foundCell Is Nothing Then
        nextRow = wsSetting.Cells(wsSetting.Rows.Count, 1).End(xlUp).Row + 1
        If nextRow < 2 Then nextRow = 2
        wsSetting.Cells(nextRow, 1).Value = settingKey
        wsSetting.Cells(nextRow, 2).Value = settingValue
    Else
        foundCell.Offset(0, 1).Value = settingValue
    End If

    wsSetting.Visible = xlSheetVeryHidden
    ThisWorkbook.Save
    Exit Sub

EH:
    ReportMacroError PROC_NAME, "アドイン設定の保存", Err.Number, Err.Description, ThisWorkbook, wsSetting
End Sub

Public Function GetAddInSettingWorksheet(ByVal createIfMissing As Boolean) As Worksheet
    ' 設定保存用の VeryHidden シートを取得します。必要な場合だけ自動作成します。
    Dim wsSetting As Worksheet

    On Error Resume Next
    Set wsSetting = ThisWorkbook.Worksheets(SETTING_SHEET_NAME)
    On Error GoTo 0

    If wsSetting Is Nothing And createIfMissing Then
        Set wsSetting = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsSetting.Name = SETTING_SHEET_NAME
        wsSetting.Range("A1:B1").Value = Array("キー", "値")
        wsSetting.Rows(1).Font.Bold = True
        wsSetting.Visible = xlSheetVeryHidden
    End If

    Set GetAddInSettingWorksheet = wsSetting
End Function
