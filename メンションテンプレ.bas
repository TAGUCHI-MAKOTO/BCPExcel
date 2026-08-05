Attribute VB_Name = "modTemplateFormBuilder"
Option Explicit

' ============================================================
' テンプレート選択フォーム 自動作成マクロ
'
' 【前提シート】
'   原本       A列:ID / B列:件名 / C列:本文 / D列:タグ（1行目は見出し）
'   休日マスタ A列:休日の日付（見出しの有無は問いません）
'
' 【使い方】
'   1. このbasを標準モジュールとしてインポート
'   2. Excelの「VBAプロジェクト オブジェクト モデルへのアクセスを信頼する」をON
'   3. メンションテンプレを作成 を実行（初回のみ）
'   4. メンションテンプレを開く を実行してフォームを表示
' ============================================================

Private Const FORM_NAME As String = "frmTemplateSelector"

Public Sub メンションテンプレを作成()
    Dim targetBook As Workbook
    Dim vbProj As Object, vbComp As Object, frm As Object
    Dim stage As String, errNo As Long, errText As String

    On Error GoTo CreateError

    stage = "作成先ブックの確認"
    Set targetBook = ActiveWorkbook
    If targetBook Is Nothing Then Err.Raise vbObjectError + 1000, , "作成先のブックを取得できません。"

    If Len(targetBook.Path) = 0 Then
        Err.Raise vbObjectError + 1001, , _
                  "先にブックを「Excel マクロ有効ブック（*.xlsm）」として保存してください。"
    End If

    stage = "一時フォルダーの確認"
    PrepareFormTempFolder

    stage = "VBAプロジェクトの取得"
    Set vbProj = targetBook.VBProject

    stage = "既存フォームの確認"
    On Error Resume Next
    vbProj.VBComponents.Remove vbProj.VBComponents(FORM_NAME)
    Err.Clear
    On Error GoTo CreateError

    stage = "UserFormコンポーネントの追加"
    Set vbComp = vbProj.VBComponents.Add(3) '3 = MSForm

    stage = "UserForm名の設定"
    vbComp.Name = FORM_NAME

    stage = "UserFormデザイナーの取得"
    Set frm = vbComp.Designer

    stage = "フォーム本体のプロパティ設定"
    With vbComp
        .Properties("Caption").Value = "テンプレート選択"
        .Properties("Width").Value = 1110
        .Properties("Height").Value = 630
        .Properties("BackColor").Value = RGB(242, 246, 243)
        .Properties("StartUpPosition").Value = 1
    End With

    stage = "件名・本文欄の作成"
    AddLabel frm, "lblSubject", "件名", 18, 18, 80, 18, True
    AddTextBox frm, "txtSubject", 18, 40, 510, 24, False
    AddLabel frm, "lblBody", "本文", 18, 72, 80, 18, True
    AddTextBox frm, "txtBody", 18, 94, 510, 486, True

    stage = "対象日欄の作成"
    AddLabel frm, "lblDate", "対象日", 895, 15, 55, 18, True
    AddTextBox frm, "txtTargetDate", 954, 12, 120, 24, False

    stage = "テンプレートリストの作成"
    AddLabel frm, "lblList", "テンプレート一覧（ID順）", 548, 52, 250, 18, True
    AddLabel frm, "lblCount", "0件", 1004, 52, 70, 18, False
    frm.Controls("lblCount").TextAlign = 3
    AddHeaderBar frm, 548, 74, 526, 20
    AddHeaderText frm, "lblHeadID", "ID", 554, 76, 34, 16
    AddHeaderText frm, "lblHeadSubject", "件名", 599, 76, 264, 16
    AddHeaderText frm, "lblHeadTag", "タグ", 874, 76, 194, 16
    AddHeaderDivider frm, "lblHeadSep1", 593, 74, 20
    AddHeaderDivider frm, "lblHeadSep2", 868, 74, 20
    AddListBox frm, "lstTemplate", 548, 93, 526, 398

    stage = "検索欄の作成"
    AddTextBox frm, "txtSearch", 548, 505, 526, 28, False
    AddSearchHintLabel frm, "lblSearchHint", "キーワードを入力", 558, 510, 220, 18

    stage = "絞り込み操作欄の作成"
    AddCheckBox frm, "chkMark1", "★出社", 548, 545, 84, 36
    AddCheckBox frm, "chkMark2", "☆公休", 644, 545, 84, 36
    AddCheckBox frm, "chkMark4", "至急", 740, 545, 70, 36
    AddCheckBox frm, "chkSVOnly", "LD・SV限", 822, 545, 108, 36

    AddButton frm, "btnClear", "クリア", 944, 545, 130, 36, RGB(218, 229, 222)

    stage = "フォーム処理コードの登録"
    vbComp.CodeModule.AddFromString BuildFormCode()

    MsgBox "テンプレート選択フォームを作成しました。" & vbCrLf & _
           "次回から「メンションテンプレを開く」を実行してください。", vbInformation
    Exit Sub

CreateError:
    errNo = Err.Number
    errText = Err.Description

    On Error Resume Next
    If Not vbComp Is Nothing Then
        If vbComp.Name <> FORM_NAME Or stage <> "既存フォームの確認" Then
            vbProj.VBComponents.Remove vbComp
        End If
    End If
    On Error GoTo 0

    MsgBox "フォームを作成できませんでした。" & vbCrLf & vbCrLf & _
           "発生工程: " & stage & vbCrLf & _
           "エラー番号: " & CStr(errNo) & vbCrLf & _
           "詳細: " & errText & vbCrLf & vbCrLf & _
           BuildErrorAdvice(stage, errNo), vbExclamation
End Sub

Private Sub PrepareFormTempFolder()
    Dim fso As Object, shellObj As Object
    Dim tempPath As String, localTemp As String, shortPath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    tempPath = Trim$(Environ$("TEMP"))

    If Len(tempPath) = 0 Or Not fso.FolderExists(tempPath) Then
        localTemp = Trim$(Environ$("LOCALAPPDATA"))
        If Len(localTemp) > 0 Then localTemp = fso.BuildPath(localTemp, "Temp")

        If Len(localTemp) = 0 Or Not fso.FolderExists(localTemp) Then
            Err.Raise 76, , _
                      "Windowsの一時フォルダーが見つかりません。" & vbCrLf & _
                      "TEMP: " & tempPath & vbCrLf & "LOCALAPPDATA\Temp: " & localTemp
        End If
        tempPath = localTemp
    End If

    On Error Resume Next
    shortPath = fso.GetFolder(tempPath).ShortPath
    On Error GoTo 0
    If Len(shortPath) > 0 Then tempPath = shortPath

    Set shellObj = CreateObject("WScript.Shell")
    shellObj.Environment("PROCESS")("TEMP") = tempPath
    shellObj.Environment("PROCESS")("TMP") = tempPath
End Sub

Private Function BuildErrorAdvice(ByVal stage As String, ByVal errNo As Long) As String
    If stage = "VBAプロジェクトの取得" Then
        BuildErrorAdvice = _
            "「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」をONにし、" & _
            "Excelを完全に閉じてから開き直してください。"
    ElseIf errNo = 76 Then
        BuildErrorAdvice = _
            "パス関連のエラーです。ブックをPC内の短いパス（例: C:\ExcelWork）へ保存し、" & _
            "Excelを開き直してから再実行してください。"
    Else
        BuildErrorAdvice = _
            "ブックが*.xlsmで保存されていることと、読み取り専用になっていないことを確認してください。"
    End If
End Function

Public Sub メンションテンプレを開く()
    On Error GoTo NotCreated
    VBA.UserForms.Add(FORM_NAME).Show
    Exit Sub
NotCreated:
    MsgBox "フォームがまだ作成されていません。" & vbCrLf & _
           "先に「メンションテンプレを作成」を実行してください。", vbExclamation
End Sub

Private Sub AddLabel(ByVal frm As Object, ByVal nm As String, ByVal cap As String, _
                     ByVal x As Single, ByVal y As Single, ByVal w As Single, _
                     ByVal h As Single, ByVal isBold As Boolean)
    Dim c As Object
    Set c = frm.Controls.Add("Forms.Label.1", nm, True)
    With c
        .Caption = cap: .Left = x: .Top = y: .Width = w: .Height = h
        .Font.Name = "Yu Gothic UI": .Font.Size = 10: .Font.Bold = isBold
        .BackStyle = 0: .ForeColor = RGB(48, 64, 54)
    End With
End Sub

Private Sub AddHeaderBar(ByVal frm As Object, ByVal x As Single, ByVal y As Single, _
                         ByVal w As Single, ByVal h As Single)
    Dim c As Object
    Set c = frm.Controls.Add("Forms.Label.1", "lblHeaderBar", True)
    With c
        .Caption = "": .Left = x: .Top = y: .Width = w: .Height = h
        .BackStyle = 1: .BackColor = RGB(218, 229, 222): .BorderStyle = 0
    End With
End Sub

Private Sub AddHeaderText(ByVal frm As Object, ByVal nm As String, ByVal cap As String, _
                          ByVal x As Single, ByVal y As Single, ByVal w As Single, _
                          ByVal h As Single)
    Dim c As Object
    Set c = frm.Controls.Add("Forms.Label.1", nm, True)
    With c
        .Caption = cap: .Left = x: .Top = y: .Width = w: .Height = h
        .Font.Name = "Yu Gothic UI": .Font.Size = 9: .Font.Bold = True
        .BackStyle = 0: .ForeColor = RGB(40, 58, 47): .BorderStyle = 0
    End With
End Sub

Private Sub AddHeaderDivider(ByVal frm As Object, ByVal nm As String, _
                             ByVal x As Single, ByVal y As Single, ByVal h As Single)
    Dim c As Object
    Set c = frm.Controls.Add("Forms.Label.1", nm, True)
    With c
        .Caption = "": .Left = x: .Top = y + 3: .Width = 1: .Height = h - 6
        .BackStyle = 1: .BackColor = RGB(166, 181, 171): .BorderStyle = 0
    End With
End Sub

Private Sub AddTextBox(ByVal frm As Object, ByVal nm As String, ByVal x As Single, _
                       ByVal y As Single, ByVal w As Single, ByVal h As Single, _
                       ByVal multi As Boolean)
    Dim c As Object
    Set c = frm.Controls.Add("Forms.TextBox.1", nm, True)
    With c
        .Left = x: .Top = y: .Width = w: .Height = h
        .Font.Name = "Yu Gothic UI": .Font.Size = 11
        .BackColor = RGB(255, 255, 252): .ForeColor = RGB(40, 52, 44)
        .BorderStyle = 1: .MultiLine = multi
        If multi Then .EnterKeyBehavior = True: .ScrollBars = 2
    End With
End Sub

Private Sub AddSearchHintLabel(ByVal frm As Object, ByVal nm As String, ByVal cap As String, _
                               ByVal x As Single, ByVal y As Single, ByVal w As Single, _
                               ByVal h As Single)
    Dim c As Object
    Set c = frm.Controls.Add("Forms.Label.1", nm, True)
    With c
        .Caption = cap: .Left = x: .Top = y: .Width = w: .Height = h
        .Font.Name = "Yu Gothic UI": .Font.Size = 11
        .BackStyle = 1: .BackColor = RGB(255, 255, 252)
        .ForeColor = RGB(132, 143, 136)
    End With
End Sub

Private Sub AddCheckBox(ByVal frm As Object, ByVal nm As String, ByVal cap As String, _
                        ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single)
    Dim c As Object
    Set c = frm.Controls.Add("Forms.CheckBox.1", nm, True)
    With c
        .Caption = cap: .Left = x: .Top = y: .Width = w: .Height = h
        .Font.Name = "Yu Gothic UI": .Font.Size = 12: .Font.Bold = True
        .BackColor = RGB(242, 246, 243): .ForeColor = RGB(48, 64, 54)
    End With
End Sub

Private Sub AddButton(ByVal frm As Object, ByVal nm As String, ByVal cap As String, _
                      ByVal x As Single, ByVal y As Single, ByVal w As Single, _
                      ByVal h As Single, ByVal bg As Long)
    Dim c As Object
    Set c = frm.Controls.Add("Forms.CommandButton.1", nm, True)
    With c
        .Caption = cap: .Left = x: .Top = y: .Width = w: .Height = h
        .Font.Name = "Yu Gothic UI": .Font.Size = 11: .Font.Bold = True
        .BackColor = bg: .ForeColor = RGB(40, 58, 47)
        .TakeFocusOnClick = False
    End With
End Sub

Private Sub AddListBox(ByVal frm As Object, ByVal nm As String, ByVal x As Single, _
                       ByVal y As Single, ByVal w As Single, ByVal h As Single)
    Dim c As Object
    Set c = frm.Controls.Add("Forms.ListBox.1", nm, True)
    With c
        .Left = x: .Top = y: .Width = w: .Height = h
        .Font.Name = "Yu Gothic UI": .Font.Size = 11
        .BackColor = RGB(255, 255, 252): .ForeColor = RGB(40, 52, 44)
        .BorderStyle = 1: .ColumnCount = 4
        .ColumnWidths = "45 pt;275 pt;185 pt;0 pt"
        .IntegralHeight = False
    End With
End Sub

Private Function BuildFormCode() As String
    Dim q As String, s As String
    q = Chr$(34)

    AddLine s, "Option Explicit"
    AddLine s, "Private mLoading As Boolean"
    AddLine s, "Private mCopying As Boolean"
    AddLine s, "Private Sub UserForm_Initialize()"
    AddLine s, "  mLoading = True"
    AddLine s, "  txtTargetDate.Value = Format$(Date, " & q & "yyyy/mm/dd" & q & ")"
    AddLine s, "  txtSearch.Value = " & q & q
    AddLine s, "  lblSearchHint.Visible = True"
    AddLine s, "  mLoading = False"
    AddLine s, "  RefreshList"
    AddLine s, "End Sub"

    AddLine s, "Private Sub txtTargetDate_Change()"
    AddLine s, "  If mLoading Then Exit Sub"
    AddLine s, "  If IsDate(txtTargetDate.Value) Then"
    AddLine s, "    RefreshList"
    AddLine s, "  Else"
    AddLine s, "    lstTemplate.Clear"
    AddLine s, "    lblCount.Caption = " & q & "0件" & q
    AddLine s, "    txtSubject.Value = " & q & q
    AddLine s, "    txtBody.Value = " & q & q
    AddLine s, "  End If"
    AddLine s, "End Sub"

    AddLine s, "Private Sub txtSearch_Change()"
    AddLine s, "  If mLoading Then Exit Sub"
    AddLine s, "  If Len(Trim$(CStr(txtSearch.Value))) > 0 Then lblSearchHint.Visible = False"
    AddLine s, "  RefreshList"
    AddLine s, "End Sub"

    AddLine s, "Private Sub txtSearch_Enter()"
    AddLine s, "  lblSearchHint.Visible = False"
    AddLine s, "End Sub"

    AddLine s, "Private Sub txtSearch_Exit(ByVal Cancel As MSForms.ReturnBoolean)"
    AddLine s, "  If Len(Trim$(CStr(txtSearch.Value))) = 0 Then"
    AddLine s, "    lblSearchHint.Visible = True"
    AddLine s, "  End If"
    AddLine s, "End Sub"

    AddLine s, "Private Sub lblSearchHint_Click()"
    AddLine s, "  txtSearch.SetFocus"
    AddLine s, "End Sub"

    AddLine s, "Private Sub txtSubject_DblClick(ByVal Cancel As MSForms.ReturnBoolean)"
    AddLine s, "  Cancel = True"
    AddLine s, "  CopyControlText txtSubject"
    AddLine s, "End Sub"

    AddLine s, "Private Sub txtBody_DblClick(ByVal Cancel As MSForms.ReturnBoolean)"
    AddLine s, "  Cancel = True"
    AddLine s, "  CopyControlText txtBody"
    AddLine s, "End Sub"

    AddLine s, "Private Sub CopyControlText(ByVal targetControl As Object)"
    AddLine s, "  Dim clipboard As MSForms.DataObject, copyText As String"
    AddLine s, "  If mCopying Then Exit Sub"
    AddLine s, "  copyText = CStr(targetControl.Value)"
    AddLine s, "  If Len(copyText) = 0 Then Exit Sub"
    AddLine s, "  On Error GoTo CopyError"
    AddLine s, "  mCopying = True"
    AddLine s, "  Set clipboard = New MSForms.DataObject"
    AddLine s, "  clipboard.SetText copyText"
    AddLine s, "  clipboard.PutInClipboard"
    AddLine s, "  FlashCopyHighlight targetControl"
    AddLine s, "  mCopying = False"
    AddLine s, "  Exit Sub"
    AddLine s, "CopyError:"
    AddLine s, "  mCopying = False"
    AddLine s, "  On Error Resume Next"
    AddLine s, "  targetControl.BackColor = RGB(255, 255, 252)"
    AddLine s, "  On Error GoTo 0"
    AddLine s, "  MsgBox " & q & "コピーできませんでした。もう一度お試しください。" & q & ", vbExclamation"
    AddLine s, "End Sub"

    AddLine s, "Private Sub FlashCopyHighlight(ByVal targetControl As Object)"
    AddLine s, "  Dim started As Single, elapsed As Double"
    AddLine s, "  targetControl.BackColor = RGB(255, 244, 170)"
    AddLine s, "  Me.Repaint"
    AddLine s, "  started = Timer"
    AddLine s, "  Do"
    AddLine s, "    DoEvents"
    AddLine s, "    If Timer >= started Then"
    AddLine s, "      elapsed = Timer - started"
    AddLine s, "    Else"
    AddLine s, "      elapsed = (86400# - CDbl(started)) + CDbl(Timer)"
    AddLine s, "    End If"
    AddLine s, "  Loop While elapsed < 0.3"
    AddLine s, "  targetControl.BackColor = RGB(255, 255, 252)"
    AddLine s, "  Me.Repaint"
    AddLine s, "End Sub"

    AddLine s, "Private Sub chkMark1_Click()"
    AddLine s, "  If mLoading Then Exit Sub"
    AddLine s, "  ApplyCheckRules 1"
    AddLine s, "End Sub"

    AddLine s, "Private Sub chkMark2_Click()"
    AddLine s, "  If mLoading Then Exit Sub"
    AddLine s, "  ApplyCheckRules 2"
    AddLine s, "End Sub"

    AddLine s, "Private Sub chkSVOnly_Click()"
    AddLine s, "  If mLoading Then Exit Sub"
    AddLine s, "  ApplyCheckRules 3"
    AddLine s, "End Sub"

    AddLine s, "Private Sub chkMark4_Click()"
    AddLine s, "  If mLoading Then Exit Sub"
    AddLine s, "  ApplyCheckRules 4"
    AddLine s, "End Sub"

    AddLine s, "Private Sub ApplyCheckRules(ByVal selectedNo As Long)"
    AddLine s, "  mLoading = True"
    AddLine s, "  Select Case selectedNo"
    AddLine s, "    Case 1"
    AddLine s, "      If chkMark1.Value Then"
    AddLine s, "        chkMark2.Value = False"
    AddLine s, "        chkSVOnly.Value = False"
    AddLine s, "      End If"
    AddLine s, "    Case 2"
    AddLine s, "      If chkMark2.Value Then"
    AddLine s, "        chkMark1.Value = False"
    AddLine s, "        chkSVOnly.Value = False"
    AddLine s, "      End If"
    AddLine s, "    Case 3"
    AddLine s, "      If chkSVOnly.Value Then"
    AddLine s, "        chkMark1.Value = False"
    AddLine s, "        chkMark2.Value = False"
    AddLine s, "        chkMark4.Value = False"
    AddLine s, "      End If"
    AddLine s, "    Case 4"
    AddLine s, "      If chkMark4.Value Then"
    AddLine s, "        chkSVOnly.Value = False"
    AddLine s, "      End If"
    AddLine s, "  End Select"
    AddLine s, "  mLoading = False"
    AddLine s, "  RefreshList"
    AddLine s, "End Sub"

    AddLine s, "Private Sub btnClear_Click()"
    AddLine s, "  mLoading = True"
    AddLine s, "  txtSearch.Value = " & q & q
    AddLine s, "  lblSearchHint.Visible = True"
    AddLine s, "  chkMark1.Value = False: chkMark2.Value = False"
    AddLine s, "  chkSVOnly.Value = False: chkMark4.Value = False"
    AddLine s, "  mLoading = False: RefreshList"
    AddLine s, "End Sub"

    AddLine s, "Private Sub lstTemplate_Click()"
    AddLine s, "  Dim ws As Worksheet, r As Long"
    AddLine s, "  If lstTemplate.ListIndex < 0 Then Exit Sub"
    AddLine s, "  Set ws = ThisWorkbook.Worksheets(" & q & "原本" & q & ")"
    AddLine s, "  r = CLng(lstTemplate.List(lstTemplate.ListIndex, 3))"
    AddLine s, "  txtSubject.Value = CStr(ws.Cells(r, " & q & "B" & q & ").Value)"
    AddLine s, "  txtBody.Value = CStr(ws.Cells(r, " & q & "C" & q & ").Value)"
    AddLine s, "End Sub"

    AddLine s, "Private Sub RefreshList()"
    AddLine s, "  Dim ws As Worksheet, lastRow As Long, r As Long, n As Long"
    AddLine s, "  Dim data() As Variant, targetDate As Date, holidayMode As Boolean"
    AddLine s, "  Dim idVal As Variant, subject As String, body As String, tag As String"
    AddLine s, "  Dim keyword As String, markFilter As Boolean, markOK As Boolean"
    AddLine s, "  Dim i As Long, j As Long, tmp As Variant"
    AddLine s, "  On Error GoTo SheetError"
    AddLine s, "  lstTemplate.Clear: txtSubject.Value = " & q & q & ": txtBody.Value = " & q & q
    AddLine s, "  If Not IsDate(txtTargetDate.Value) Then"
    AddLine s, "    lblCount.Caption = " & q & "0件" & q
    AddLine s, "    Exit Sub"
    AddLine s, "  End If"
    AddLine s, "  targetDate = DateValue(CDate(txtTargetDate.Value))"
    AddLine s, "  holidayMode = IsHolidayDate(targetDate)"
    AddLine s, "  Set ws = ThisWorkbook.Worksheets(" & q & "原本" & q & ")"
    AddLine s, "  lastRow = ws.Cells(ws.Rows.Count, " & q & "A" & q & ").End(xlUp).Row"
    AddLine s, "  keyword = Trim$(CStr(txtSearch.Value))"
    AddLine s, "  markFilter = (chkMark1.Value Or chkMark2.Value Or chkMark4.Value)"
    AddLine s, "  ReDim data(1 To Application.Max(1, lastRow - 1), 1 To 4)"
    AddLine s, "  For r = 2 To lastRow"
    AddLine s, "    idVal = ws.Cells(r, " & q & "A" & q & ").Value"
    AddLine s, "    subject = CStr(ws.Cells(r, " & q & "B" & q & ").Value)"
    AddLine s, "    body = CStr(ws.Cells(r, " & q & "C" & q & ").Value)"
    AddLine s, "    tag = CStr(ws.Cells(r, " & q & "D" & q & ").Value)"
    AddLine s, "    If Len(Trim$(CStr(idVal))) = 0 Then GoTo ContinueRow"
    AddLine s, "    If holidayMode Then"
    AddLine s, "      If InStr(1, tag, " & q & "土日" & q & ", vbTextCompare) = 0 Then GoTo ContinueRow"
    AddLine s, "    Else"
    AddLine s, "      If InStr(1, tag, " & q & "土日" & q & ", vbTextCompare) > 0 Then GoTo ContinueRow"
    AddLine s, "    End If"
    AddLine s, "    If chkSVOnly.Value Then"
    AddLine s, "      If InStr(1, tag, " & q & "SV限" & q & ", vbTextCompare) = 0 Then GoTo ContinueRow"
    AddLine s, "    Else"
    AddLine s, "      If InStr(1, tag, " & q & "SV限" & q & ", vbTextCompare) > 0 Then GoTo ContinueRow"
    AddLine s, "    End If"
    AddLine s, "    If keyword <> " & q & q & " Then"
    AddLine s, "      If InStr(1, body, keyword, vbTextCompare) = 0 And InStr(1, tag, keyword, vbTextCompare) = 0 Then GoTo ContinueRow"
    AddLine s, "    End If"
    AddLine s, "    If markFilter Then"
    AddLine s, "      markOK = False"
    AddLine s, "      If chkMark4.Value Then"
    AddLine s, "        If chkMark1.Value Then"
    AddLine s, "          If InStr(1, body, " & q & "★★★" & q & ", vbBinaryCompare) > 0 Then markOK = True"
    AddLine s, "        ElseIf chkMark2.Value Then"
    AddLine s, "          If InStr(1, body, " & q & "☆☆☆" & q & ", vbBinaryCompare) > 0 Then markOK = True"
    AddLine s, "        Else"
    AddLine s, "          If InStr(1, body, " & q & "★★★" & q & ", vbBinaryCompare) > 0 Then markOK = True"
    AddLine s, "          If holidayMode And InStr(1, body, " & q & "☆☆☆" & q & ", vbBinaryCompare) > 0 Then markOK = True"
    AddLine s, "        End If"
    AddLine s, "      ElseIf chkMark1.Value Then"
    AddLine s, "        If InStr(1, body, " & q & "★" & q & ", vbBinaryCompare) > 0 Then markOK = True"
    AddLine s, "      ElseIf chkMark2.Value Then"
    AddLine s, "        If InStr(1, body, " & q & "☆" & q & ", vbBinaryCompare) > 0 Then markOK = True"
    AddLine s, "      End If"
    AddLine s, "      If Not markOK Then GoTo ContinueRow"
    AddLine s, "    End If"
    AddLine s, "    n = n + 1: data(n, 1) = idVal: data(n, 2) = subject: data(n, 3) = tag: data(n, 4) = r"
    AddLine s, "ContinueRow:"
    AddLine s, "  Next r"

    AddLine s, "  For i = 2 To n"
    AddLine s, "    For j = i To 2 Step -1"
    AddLine s, "      If CompareID(data(j - 1, 1), data(j, 1)) <= 0 Then Exit For"
    AddLine s, "      tmp = data(j - 1, 1): data(j - 1, 1) = data(j, 1): data(j, 1) = tmp"
    AddLine s, "      tmp = data(j - 1, 2): data(j - 1, 2) = data(j, 2): data(j, 2) = tmp"
    AddLine s, "      tmp = data(j - 1, 3): data(j - 1, 3) = data(j, 3): data(j, 3) = tmp"
    AddLine s, "      tmp = data(j - 1, 4): data(j - 1, 4) = data(j, 4): data(j, 4) = tmp"
    AddLine s, "    Next j"
    AddLine s, "  Next i"
    AddLine s, "  For i = 1 To n"
    AddLine s, "    lstTemplate.AddItem CStr(data(i, 1))"
    AddLine s, "    lstTemplate.List(lstTemplate.ListCount - 1, 1) = CStr(data(i, 2))"
    AddLine s, "    lstTemplate.List(lstTemplate.ListCount - 1, 2) = CStr(data(i, 3))"
    AddLine s, "    lstTemplate.List(lstTemplate.ListCount - 1, 3) = CLng(data(i, 4))"
    AddLine s, "  Next i"
    AddLine s, "  lblCount.Caption = CStr(n) & " & q & "件" & q
    AddLine s, "  Exit Sub"
    AddLine s, "SheetError:"
    AddLine s, "  lblCount.Caption = " & q & "0件" & q
    AddLine s, "  MsgBox " & q & "シート「原本」または「休日マスタ」を確認してください。" & q & " & vbCrLf & Err.Description, vbExclamation"
    AddLine s, "End Sub"

    AddLine s, "Private Function IsHolidayDate(ByVal d As Date) As Boolean"
    AddLine s, "  Dim ws As Worksheet, lastRow As Long, r As Long, v As Variant"
    AddLine s, "  If Weekday(d, vbMonday) >= 6 Then"
    AddLine s, "    IsHolidayDate = True"
    AddLine s, "    Exit Function"
    AddLine s, "  End If"
    AddLine s, "  Set ws = ThisWorkbook.Worksheets(" & q & "休日マスタ" & q & ")"
    AddLine s, "  lastRow = ws.Cells(ws.Rows.Count, " & q & "A" & q & ").End(xlUp).Row"
    AddLine s, "  For r = 1 To lastRow"
    AddLine s, "    v = ws.Cells(r, " & q & "A" & q & ").Value"
    AddLine s, "    If IsDate(v) Then"
    AddLine s, "      If DateValue(CDate(v)) = DateValue(d) Then"
    AddLine s, "        IsHolidayDate = True"
    AddLine s, "        Exit Function"
    AddLine s, "      End If"
    AddLine s, "    End If"
    AddLine s, "  Next r"
    AddLine s, "End Function"

    AddLine s, "Private Function CompareID(ByVal a As Variant, ByVal b As Variant) As Long"
    AddLine s, "  If IsNumeric(a) And IsNumeric(b) Then"
    AddLine s, "    CompareID = Sgn(CDbl(a) - CDbl(b))"
    AddLine s, "  Else"
    AddLine s, "    CompareID = StrComp(CStr(a), CStr(b), vbTextCompare)"
    AddLine s, "  End If"
    AddLine s, "End Function"

    BuildFormCode = s
End Function

Private Sub AddLine(ByRef target As String, ByVal lineText As String)
    target = target & lineText & vbCrLf
End Sub

