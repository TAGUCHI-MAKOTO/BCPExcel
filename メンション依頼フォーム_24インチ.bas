Attribute VB_Name = "modMentionRequestForm24"
Option Explicit

'============================================================
' 依頼入力ユーザーフォーム 自動生成（24インチモニター版）
' 変更点：
' ・代理CA名3欄の Top を固定値に統一
' ・各依頼セクションで同一座標を使用
' ・通常の TextBox 枠線に戻して余計な補助枠は使わない
' 24インチ FHD（1920×1080 / Windows表示倍率100%前後）を想定
' 通常版のレイアウト・配色はそのままに、全体を82%へ縮小
' ・処理日時の次に「期日」入力欄を追加
' ・オプション／タイプを下段へ移動
'============================================================

Private Const FORM_NAME As String = "frmRequestEntry"
Private Const UI_SCALE As Double = 0.82

Private Const vbext_ct_MSForm As Long = 3

Private Const fmBorderStyleSingle As Long = 1
Private Const fmSpecialEffectFlat As Long = 0
Private Const fmStyleDropDownList As Long = 2
Private Const fmScrollBarsVertical As Long = 2

Private Const CLR_FORM_BG As Long = &HFAF6FB
Private Const CLR_WHITE As Long = &HFFFFFF
Private Const CLR_TEXT As Long = &H333333
Private Const CLR_MUTED As Long = &H666666

Private Const CLR_ACCENT As Long = &H8B5E8E
Private Const CLR_HEADER_BOX As Long = &HF4ECF7

Private Const CLR_SECTION1 As Long = &HFAF1FF
Private Const CLR_SECTION2 As Long = &HF2FBF4
Private Const CLR_SECTION3 As Long = &HFFF6F0

Private Const CLR_FIELD_BG As Long = &HFFFFFF
Private Const CLR_BUTTON As Long = &HB88FC2
Private Const CLR_BUTTON_TEXT As Long = &HFFFFFF

Public Sub BuildMentionRequestForm24()

    Dim vbProj As Object
    Dim vbComp As Object
    Dim designer As Object
    Dim oldComp As Object
    Dim ans As VbMsgBoxResult

    On Error GoTo EH

    Set vbProj = ThisWorkbook.VBProject

    On Error Resume Next
    Set oldComp = vbProj.VBComponents(FORM_NAME)
    On Error GoTo EH

    If Not oldComp Is Nothing Then
        ans = MsgBox(FORM_NAME & " は既に存在します。" & vbCrLf & _
                     "削除して作り直しますか？", _
                     vbYesNo + vbQuestion, "フォーム再作成")
        If ans <> vbYes Then Exit Sub
        vbProj.VBComponents.Remove oldComp
    End If

    Set vbComp = vbProj.VBComponents.Add(vbext_ct_MSForm)
    vbComp.Name = FORM_NAME

    With vbComp.Properties
        .Item("Caption") = "依頼入力"
        .Item("Width") = S(1435)
        .Item("Height") = S(820)
        .Item("StartUpPosition") = 2
        .Item("BackColor") = CLR_FORM_BG
        .Item("ScrollBars") = fmScrollBarsVertical
        .Item("KeepScrollBarsVisible") = 1
    End With

    Set designer = vbComp.Designer

    BuildFormLayout designer
    AddCodeStub vbComp

    MsgBox "24インチモニター用フォームを作成しました。", vbInformation, "作成完了"
    Exit Sub

EH:
    MsgBox "フォーム作成中にエラーが発生しました。" & vbCrLf & _
           "Err " & Err.Number & " : " & Err.Description, _
           vbExclamation, "エラー"
End Sub

Private Sub BuildFormLayout(ByVal designer As Object)

    Dim contentW As Double
    Dim marginX As Double
    Dim y As Double
    Dim secGap As Double

    marginX = 18
    contentW = 1380

    ' タイトル帯は削除
    ' その分、フォーム上部に少し余裕を持たせる
    y = 22

    ' セクション間は少し広め
    secGap = 24

    ' 自動取得エリア
    BuildHeaderArea designer, marginX, y, contentW, 52

    ' ヘッダー下もゆったりめ
    y = y + 78

    ' 入力セクション
    BuildRequestSection designer, 1, marginX, y, contentW, 184, CLR_SECTION1
    y = y + 184 + secGap

    BuildRequestSection designer, 2, marginX, y, contentW, 184, CLR_SECTION2
    y = y + 184 + secGap

    BuildRequestSection designer, 3, marginX, y, contentW, 184, CLR_SECTION3
    y = y + 184 + 28

    ' 送信ボタン
    AddCommandButton designer, "cmdSend", "送信", _
                     marginX + contentW - 170, y, 150, 42, True

End Sub

Private Sub BuildHeaderArea(ByVal designer As Object, _
                            ByVal leftX As Double, ByVal topY As Double, _
                            ByVal areaW As Double, ByVal areaH As Double)

    Dim fra As Object
    Dim cardW As Double
    Dim gapX As Double

    Set fra = AddFrame(designer, "fraHeaderArea", "", _
                       leftX, topY, areaW, areaH, RGB(255, 255, 255))

    AddColorBar fra, "barHeader", 0, 0, 9, areaH, CLR_ACCENT

    ' 横方向も少し余裕を持たせる
    gapX = 28
    cardW = (areaW - 36 - gapX * 2) / 3

    ' 自動取得値なので、入力欄のような高さは使わず
    ' ラベルに近い高さでコンパクト表示
    BuildHeaderCard fra, "hdrBase", "依頼拠点", _
                    18, 7, cardW, 34

    BuildHeaderCard fra, "hdrXU", "XU番号", _
                    18 + cardW + gapX, 7, cardW, 34

    BuildHeaderCard fra, "hdrRequester", "依頼者", _
                    18 + (cardW + gapX) * 2, 7, cardW, 34

End Sub

Private Sub BuildHeaderCard(ByVal parent As Object, _
                            ByVal key As String, ByVal titleText As String, _
                            ByVal leftX As Double, ByVal topY As Double, _
                            ByVal cardW As Double, ByVal cardH As Double)

    Dim box As Object

    ' 項目名
    AddLabel parent, "lbl" & key & "Title", titleText, _
             leftX, topY, cardW, 13, 8.5, True, False, CLR_MUTED

    ' 自動取得値
    ' 入力用TextBoxではなく、薄い色の小さな表示ラベルとして扱う
    Set box = AddLabel(parent, "lbl" & key & "Value", "", _
                       leftX, topY + 15, cardW, 17, _
                       9.5, True, False, CLR_TEXT)

    With box
        .BackStyle = 1
        .BackColor = CLR_HEADER_BOX
        .BorderStyle = 0
        .Caption = " "
        .SpecialEffect = fmSpecialEffectFlat
    End With

End Sub

Private Sub BuildRequestSection(ByVal designer As Object, _
                                ByVal idx As Long, _
                                ByVal leftX As Double, ByVal topY As Double, _
                                ByVal areaW As Double, ByVal areaH As Double, _
                                ByVal backColor As Long)

    Dim fra As Object
    Dim innerLeft As Double
    Dim gapX As Double
    Dim w1 As Double, w2 As Double, w3 As Double
    Dim w4 As Double, w5 As Double, w6 As Double, w7 As Double
    Dim proxyX As Double
    Dim processedX As Double
    Dim dueDateX As Double

    '--------------------------------------------------------
    ' Frame 内の縦位置
    '--------------------------------------------------------
    Const LABEL_TOP As Double = 38
    Const ROW1_TOP As Double = 58
    Const FIELD_H As Double = 28
    Const ROW_GAP As Double = 12

    Const ROW2_TOP As Double = ROW1_TOP + FIELD_H + ROW_GAP
    Const ROW3_TOP As Double = ROW2_TOP + FIELD_H + ROW_GAP

    ' オプション／タイプは1段目より少し下へ
    Const SUB_LABEL_TOP As Double = 94
    Const SUB_ROW_TOP As Double = 112

    Set fra = AddFrame(designer, "fraSection" & idx, "", _
                       leftX, topY, areaW, areaH, backColor)

    AddColorBar fra, "barSection" & idx, 0, 0, 9, areaH, CLR_ACCENT
    AddSectionTitle fra, "lblSectionTitle" & idx, "依頼 " & idx, _
                    18, 10, 74, 20

    innerLeft = 18
    gapX = 10

    w1 = 165     ' 組織
    w2 = 165     ' CA名
    w3 = 165     ' 代理CA組織
    w4 = 165     ' 代理CA名
    w5 = 165     ' メールメモ
    w6 = 165     ' 処理日時
    w7 = 165     ' 期日

    proxyX = innerLeft + _
             (w1 + gapX) + _
             (w2 + gapX) + _
             (w3 + gapX)

    processedX = innerLeft + _
                 (w1 + gapX) + _
                 (w2 + gapX) + _
                 (w3 + gapX) + _
                 (w4 + gapX) + _
                 (w5 + gapX)

    dueDateX = processedX + w6 + gapX

    AddLabel fra, "lblOrg" & idx, "組織", _
             innerLeft, LABEL_TOP, w1, 16, 9, True, False, CLR_MUTED

    AddLabel fra, "lblCA" & idx, "CA名", _
             innerLeft + w1 + gapX, LABEL_TOP, w2, 16, 9, True, False, CLR_MUTED

    AddLabel fra, "lblProxyOrg" & idx, "代理CA組織", _
             innerLeft + (w1 + gapX) + (w2 + gapX), _
             LABEL_TOP, w3, 16, 9, True, False, CLR_MUTED

    AddLabel fra, "lblProxyCA" & idx, "代理CA名", _
             proxyX, LABEL_TOP, w4, 16, 9, True, False, CLR_MUTED

    AddLabel fra, "lblMailMemo" & idx, "メールメモ", _
             innerLeft + (w1 + gapX) + (w2 + gapX) + _
             (w3 + gapX) + (w4 + gapX), _
             LABEL_TOP, w5, 16, 9, True, False, CLR_MUTED

    AddLabel fra, "lblProcessedAt" & idx, "処理日時", _
             processedX, LABEL_TOP, w6, 16, 9, True, False, CLR_MUTED

    AddLabel fra, "lblDueDate" & idx, "期日", _
             dueDateX, LABEL_TOP, w7, 16, 9, True, False, CLR_MUTED

    AddTextBox fra, "txtOrg" & idx, _
               innerLeft, ROW1_TOP, w1, FIELD_H

    AddTextBox fra, "txtCA" & idx, _
               innerLeft + w1 + gapX, ROW1_TOP, w2, FIELD_H

    AddTextBox fra, "txtProxyOrg" & idx, _
               innerLeft + (w1 + gapX) + (w2 + gapX), _
               ROW1_TOP, w3, FIELD_H

    AddTextBox fra, "txtProxyCA" & idx & "_1", _
               proxyX, ROW1_TOP, w4, FIELD_H

    AddTextBox fra, "txtMailMemo" & idx, _
               innerLeft + (w1 + gapX) + (w2 + gapX) + _
               (w3 + gapX) + (w4 + gapX), _
               ROW1_TOP, w5, FIELD_H

    AddTextBox fra, "txtProcessedAt" & idx, _
               processedX, ROW1_TOP, w6, FIELD_H

    AddTextBox fra, "txtDueDate" & idx, _
               dueDateX, ROW1_TOP, w7, FIELD_H

    AddTextBox fra, "txtProxyCA" & idx & "_2", _
               proxyX, ROW2_TOP, w4, FIELD_H

    AddTextBox fra, "txtProxyCA" & idx & "_3", _
               proxyX, ROW3_TOP, w4, FIELD_H

    AddLabel fra, "lblOption" & idx, "オプション", _
             processedX, SUB_LABEL_TOP, w6, 16, _
             9, True, False, CLR_MUTED

    AddLabel fra, "lblType" & idx, "タイプ", _
             dueDateX, SUB_LABEL_TOP, w7, 16, _
             9, True, False, CLR_MUTED

    AddOptionPanel fra, idx, _
                   processedX, SUB_ROW_TOP, w6, 58

    AddComboBox fra, "cmbType" & idx, _
                dueDateX, SUB_ROW_TOP, w7, FIELD_H

End Sub

Private Sub AddOptionPanel(ByVal parent As Object, ByVal idx As Long, _
                           ByVal leftX As Double, ByVal topY As Double, _
                           ByVal panelW As Double, ByVal panelH As Double)

    Dim pan As Object
    Set pan = AddFrame(parent, "fraOption" & idx, "", leftX, topY, panelW, panelH, CLR_WHITE)

    AddCheckBox pan, "chkShort" & idx, "時短", 10, 8, panelW - 20, 18
    AddCheckBox pan, "chkUrgent" & idx, "至急", 10, 30, panelW - 20, 18
End Sub

Private Function AddFrame(ByVal parent As Object, ByVal ctlName As String, ByVal captionText As String, _
                          ByVal leftX As Double, ByVal topY As Double, _
                          ByVal ctlW As Double, ByVal ctlH As Double, _
                          Optional ByVal backColor As Long = -1) As Object
    Dim ctl As Object
    Set ctl = parent.Controls.Add("Forms.Frame.1", ctlName, True)

    With ctl
        .Caption = captionText
        .Left = S(leftX)
        .Top = S(topY)
        .Width = S(ctlW)
        .Height = S(ctlH)
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        If backColor <> -1 Then .BackColor = backColor
        With .Font
            .Name = "Meiryo UI"
            .Size = 9
        End With
    End With

    Set AddFrame = ctl
End Function

Private Sub AddColorBar(ByVal parent As Object, ByVal ctlName As String, _
                        ByVal leftX As Double, ByVal topY As Double, _
                        ByVal ctlW As Double, ByVal ctlH As Double, _
                        ByVal backColor As Long)
    Dim ctl As Object
    Set ctl = parent.Controls.Add("Forms.Label.1", ctlName, True)
    With ctl
        .Caption = ""
        .Left = S(leftX)
        .Top = S(topY)
        .Width = S(ctlW)
        .Height = S(ctlH)
        .BackStyle = 1
        .BackColor = backColor
        .BorderStyle = 0
    End With
End Sub

Private Sub AddSectionTitle(ByVal parent As Object, ByVal ctlName As String, ByVal captionText As String, _
                            ByVal leftX As Double, ByVal topY As Double, _
                            ByVal ctlW As Double, ByVal ctlH As Double)
    Dim ctl As Object
    Set ctl = parent.Controls.Add("Forms.Label.1", ctlName, True)
    With ctl
        .Caption = " " & captionText & " "
        .Left = S(leftX)
        .Top = S(topY)
        .Width = S(ctlW)
        .Height = S(ctlH)
        .BackStyle = 1
        .BackColor = RGB(255, 255, 255)
        .BorderStyle = fmBorderStyleSingle
        .ForeColor = RGB(125, 78, 100)
        With .Font
            .Name = "Meiryo UI"
            .Size = 10
            .Bold = True
        End With
    End With
End Sub

Private Function AddLabel(ByVal parent As Object, _
                          ByVal ctlName As String, ByVal captionText As String, _
                          ByVal leftX As Double, ByVal topY As Double, _
                          ByVal ctlW As Double, ByVal ctlH As Double, _
                          Optional ByVal fontSize As Double = 10, _
                          Optional ByVal boldFlg As Boolean = False, _
                          Optional ByVal backFill As Boolean = False, _
                          Optional ByVal foreColor As Long = -1) As Object
    Dim ctl As Object
    Set ctl = parent.Controls.Add("Forms.Label.1", ctlName, True)

    With ctl
        .Caption = captionText
        .Left = S(leftX)
        .Top = S(topY)
        .Width = S(ctlW)
        .Height = S(ctlH)
        .BackStyle = IIf(backFill, 1, 0)
        If foreColor <> -1 Then .ForeColor = foreColor
        With .Font
            .Name = "Meiryo UI"
            .Size = fontSize
            .Bold = boldFlg
        End With
    End With

    Set AddLabel = ctl
End Function

Private Function AddTextBox(ByVal parent As Object, ByVal ctlName As String, _
                            ByVal leftX As Double, ByVal topY As Double, _
                            ByVal ctlW As Double, ByVal ctlH As Double) As Object
    Dim ctl As Object
    Set ctl = parent.Controls.Add("Forms.TextBox.1", ctlName, True)

    With ctl
        .Left = S(leftX)
        .Top = S(topY)
        .Width = S(ctlW)
        .Height = S(ctlH)
        .BorderStyle = fmBorderStyleSingle
        .SpecialEffect = fmSpecialEffectFlat
        .BackColor = CLR_FIELD_BG
        .ForeColor = CLR_TEXT
        With .Font
            .Name = "Meiryo UI"
            .Size = 10
        End With
    End With

    Set AddTextBox = ctl
End Function

Private Function AddCheckBox(ByVal parent As Object, ByVal ctlName As String, ByVal captionText As String, _
                             ByVal leftX As Double, ByVal topY As Double, _
                             ByVal ctlW As Double, ByVal ctlH As Double) As Object
    Dim ctl As Object
    Set ctl = parent.Controls.Add("Forms.CheckBox.1", ctlName, True)

    With ctl
        .Caption = captionText
        .Left = S(leftX)
        .Top = S(topY)
        .Width = S(ctlW)
        .Height = S(ctlH)
        .BackStyle = 0
        .ForeColor = CLR_TEXT
        With .Font
            .Name = "Meiryo UI"
            .Size = 10
        End With
    End With

    Set AddCheckBox = ctl
End Function

Private Function AddComboBox(ByVal parent As Object, ByVal ctlName As String, _
                             ByVal leftX As Double, ByVal topY As Double, _
                             ByVal ctlW As Double, ByVal ctlH As Double) As Object
    Dim ctl As Object
    Set ctl = parent.Controls.Add("Forms.ComboBox.1", ctlName, True)

    With ctl
        .Left = S(leftX)
        .Top = S(topY)
        .Width = S(ctlW)
        .Height = S(ctlH)
        .Style = fmStyleDropDownList
        .SpecialEffect = fmSpecialEffectFlat
        .BackColor = CLR_FIELD_BG
        .ForeColor = CLR_TEXT
        With .Font
            .Name = "Meiryo UI"
            .Size = 10
        End With
    End With

    Set AddComboBox = ctl
End Function

Private Function AddCommandButton(ByVal parent As Object, ByVal ctlName As String, ByVal captionText As String, _
                                  ByVal leftX As Double, ByVal topY As Double, _
                                  ByVal ctlW As Double, ByVal ctlH As Double, _
                                  Optional ByVal primaryButton As Boolean = False) As Object
    Dim ctl As Object
    Set ctl = parent.Controls.Add("Forms.CommandButton.1", ctlName, True)

    With ctl
        .Caption = captionText
        .Left = S(leftX)
        .Top = S(topY)
        .Width = S(ctlW)
        .Height = S(ctlH)
        If primaryButton Then
            .BackColor = CLR_BUTTON
            .ForeColor = CLR_BUTTON_TEXT
        End If
        With .Font
            .Name = "Meiryo UI"
            .Size = 11
            .Bold = True
        End With
    End With

    Set AddCommandButton = ctl
End Function

Private Sub AddCodeStub(ByVal vbComp As Object)
    Dim cm As Object
    Dim src As String

    Set cm = vbComp.CodeModule

    src = _
        "Option Explicit" & vbCrLf & vbCrLf & _
        "Private Sub UserForm_Initialize()" & vbCrLf & _
        "    Dim i As Long" & vbCrLf & _
        "    For i = 1 To 3" & vbCrLf & _
        "        With Me.Controls(""cmbType"" & i)" & vbCrLf & _
        "            .Clear" & vbCrLf & _
        "            .AddItem """"" & vbCrLf & _
        "            .AddItem ""通常""" & vbCrLf & _
        "            .AddItem ""確認""" & vbCrLf & _
        "            .AddItem ""代理対応""" & vbCrLf & _
        "            .AddItem ""その他""" & vbCrLf & _
        "        End With" & vbCrLf & _
        "    Next i" & vbCrLf & _
        "End Sub" & vbCrLf & vbCrLf & _
        "Private Sub cmdSend_Click()" & vbCrLf & _
        "    MsgBox ""送信処理は未実装です。ここに処理を書いてください。"", vbInformation" & vbCrLf & _
        "End Sub" & vbCrLf

    cm.AddFromString src
End Sub

Public Sub SetRequestHeader(ByVal frm As Object, _
                            ByVal requestBase As String, _
                            ByVal xuNumber As String, _
                            ByVal requester As String)

    frm.Controls("lblhdrBaseValue").Caption = " " & requestBase
    frm.Controls("lblhdrXUValue").Caption = " " & xuNumber
    frm.Controls("lblhdrRequesterValue").Caption = " " & requester
End Sub

Private Function S(ByVal v As Double) As Double
    S = v * UI_SCALE
End Function