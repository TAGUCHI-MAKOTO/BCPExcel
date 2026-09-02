Attribute VB_Name = "modMentionRequestForm24V2"
Option Explicit

'============================================================
' 依頼入力ユーザーフォーム 自動生成（24インチモニター版）
' 変更点：
' ・代理CA名3欄の Top を固定値に統一
' ・各依頼セクションで同一座標を使用
' ・通常の TextBox 枠線に戻して余計な補助枠は使わない
' 24インチ FHD（1920×1080 / Windows表示倍率100%前後）を想定
' 通常版のレイアウト・配色はそのままに、全体を82%へ縮小
' ・「処理日時」の次に手入力用の「期日」を追加
' ・「オプション」「タイプ」は期日の右側へ移動
' ・上部の「XU番号」を削除
' ・「拠点」はプルダウン選択
' ・「依頼者」はWindowsアカウント表示名を自動取得
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

Public Sub BuildMentionRequestForm24V2()

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
        .Item("Width") = S(1610)
        .Item("Height") = S(820)
        .Item("StartUpPosition") = 2
        .Item("BackColor") = CLR_FORM_BG
        .Item("ScrollBars") = fmScrollBarsVertical
        .Item("KeepScrollBarsVisible") = 1
    End With

    Set designer = vbComp.Designer

    BuildFormLayout designer
    AddCodeStub vbComp

    MsgBox "24インチ版フォーム（拠点選択・依頼者自動取得）を作成しました。", vbInformation, "作成完了"
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
    contentW = 1550

    y = 22
    secGap = 24

    BuildHeaderArea designer, marginX, y, contentW, 52
    y = y + 78

    BuildRequestSection designer, 1, marginX, y, contentW, 184, CLR_SECTION1
    y = y + 184 + secGap

    BuildRequestSection designer, 2, marginX, y, contentW, 184, CLR_SECTION2
    y = y + 184 + secGap

    BuildRequestSection designer, 3, marginX, y, contentW, 184, CLR_SECTION3
    y = y + 184 + 28

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

    gapX = 36
    cardW = (areaW - 54 - gapX) / 2

    AddLabel fra, "lblBaseTitle", "拠点", _
             18, 7, cardW, 13, 8.5, True, False, CLR_MUTED

    AddComboBox fra, "cmbRequestBase", _
                18, 22, cardW, 22

    BuildHeaderCard fra, "hdrRequester", "依頼者", _
                    18 + cardW + gapX, 7, cardW, 34

End Sub

Private Sub BuildHeaderCard(ByVal parent As Object, _
                            ByVal key As String, ByVal titleText As String, _
                            ByVal leftX As Double, ByVal topY As Double, _
                            ByVal cardW As Double, ByVal cardH As Double)

    Dim box As Object

    AddLabel parent, "lbl" & key & "Title", titleText, _
             leftX, topY, cardW, 13, 8.5, True, False, CLR_MUTED

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
    Dim rightLeft As Double, rightW As Double
    Dim proxyX As Double
    Dim mailMemoX As Double
    Dim processedX As Double
    Dim dueDateX As Double

    Const LABEL_TOP As Double = 38
    Const ROW1_TOP As Double = 58
    Const FIELD_H As Double = 28
    Const ROW_GAP As Double = 12

    Const ROW2_TOP As Double = ROW1_TOP + FIELD_H + ROW_GAP
    Const ROW3_TOP As Double = ROW2_TOP + FIELD_H + ROW_GAP

    Set fra = AddFrame(designer, "fraSection" & idx, "", _
                       leftX, topY, areaW, areaH, backColor)

    AddColorBar fra, "barSection" & idx, 0, 0, 9, areaH, CLR_ACCENT
    AddSectionTitle fra, "lblSectionTitle" & idx, "依頼 " & idx, _
                    18, 10, 74, 20

    innerLeft = 18
    gapX = 10

    w1 = 165
    w2 = 165
    w3 = 165
    w4 = 165
    w5 = 165
    w6 = 165
    w7 = 135

    proxyX = innerLeft + (w1 + gapX) + (w2 + gapX) + (w3 + gapX)
    mailMemoX = proxyX + w4 + gapX
    processedX = mailMemoX + w5 + gapX
    dueDateX = processedX + w6 + gapX

    rightLeft = dueDateX + w7 + gapX
    rightW = areaW - rightLeft - 20

    AddLabel fra, "lblOrg" & idx, "組織", innerLeft, LABEL_TOP, w1, 16, 9, True, False, CLR_MUTED
    AddLabel fra, "lblCA" & idx, "CA名", innerLeft + w1 + gapX, LABEL_TOP, w2, 16, 9, True, False, CLR_MUTED
    AddLabel fra, "lblProxyOrg" & idx, "代理CA組織", innerLeft + (w1 + gapX) + (w2 + gapX), LABEL_TOP, w3, 16, 9, True, False, CLR_MUTED
    AddLabel fra, "lblProxyCA" & idx, "代理CA名", proxyX, LABEL_TOP, w4, 16, 9, True, False, CLR_MUTED
    AddLabel fra, "lblMailMemo" & idx, "メールメモ", mailMemoX, LABEL_TOP, w5, 16, 9, True, False, CLR_MUTED
    AddLabel fra, "lblProcessedAt" & idx, "処理日時", processedX, LABEL_TOP, w6, 16, 9, True, False, CLR_MUTED
    AddLabel fra, "lblDueDate" & idx, "期日", dueDateX, LABEL_TOP, w7, 16, 9, True, False, CLR_MUTED

    AddLabel fra, "lblOption" & idx, "オプション", _
             rightLeft, LABEL_TOP, rightW * 0.42, 16, 9, True, False, CLR_MUTED
    AddLabel fra, "lblType" & idx, "タイプ", _
             rightLeft + rightW * 0.46, LABEL_TOP, rightW * 0.5, 16, 9, True, False, CLR_MUTED

    AddTextBox fra, "txtOrg" & idx, innerLeft, ROW1_TOP, w1, FIELD_H
    AddTextBox fra, "txtCA" & idx, innerLeft + w1 + gapX, ROW1_TOP, w2, FIELD_H
    AddTextBox fra, "txtProxyOrg" & idx, innerLeft + (w1 + gapX) + (w2 + gapX), ROW1_TOP, w3, FIELD_H
    AddTextBox fra, "txtProxyCA" & idx & "_1", proxyX, ROW1_TOP, w4, FIELD_H
    AddTextBox fra, "txtMailMemo" & idx, mailMemoX, ROW1_TOP, w5, FIELD_H
    AddTextBox fra, "txtProcessedAt" & idx, processedX, ROW1_TOP, w6, FIELD_H
    AddTextBox fra, "txtDueDate" & idx, dueDateX, ROW1_TOP, w7, FIELD_H

    AddTextBox fra, "txtProxyCA" & idx & "_2", proxyX, ROW2_TOP, w4, FIELD_H
    AddTextBox fra, "txtProxyCA" & idx & "_3", proxyX, ROW3_TOP, w4, FIELD_H

    AddOptionPanel fra, idx, rightLeft, ROW1_TOP, rightW * 0.42, 58
    AddComboBox fra, "cmbType" & idx, _
                rightLeft + rightW * 0.46, ROW1_TOP, rightW * 0.5, FIELD_H

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
        "    Dim i As Long" & vbCrLf & vbCrLf & _
        "    With Me.Controls(""cmbRequestBase"")" & vbCrLf & _
        "        .Clear" & vbCrLf & _
        "        .AddItem ""呉服""" & vbCrLf & _
        "        .AddItem ""札幌""" & vbCrLf & _
        "        .AddItem ""新潟""" & vbCrLf & _
        "        .ListIndex = -1" & vbCrLf & _
        "    End With" & vbCrLf & vbCrLf & _
        "    Me.Controls(""lblhdrRequesterValue"").Caption = "" "" & GetWindowsDisplayName()" & vbCrLf & vbCrLf & _
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
        "Private Function GetWindowsDisplayName() As String" & vbCrLf & _
        "    Dim userName As String" & vbCrLf & _
        "    Dim domainName As String" & vbCrLf & _
        "    Dim objUser As Object" & vbCrLf & vbCrLf & _
        "    On Error GoTo Fallback" & vbCrLf & vbCrLf & _
        "    userName = Environ$(""USERNAME"")" & vbCrLf & _
        "    domainName = Environ$(""USERDOMAIN"")" & vbCrLf & vbCrLf & _
        "    If Len(domainName) > 0 And Len(userName) > 0 Then" & vbCrLf & _
        "        Set objUser = GetObject(""WinNT://"" & domainName & ""/"" & userName & "",user"")" & vbCrLf & _
        "        If Len(Trim$(CStr(objUser.FullName))) > 0 Then" & vbCrLf & _
        "            GetWindowsDisplayName = CStr(objUser.FullName)" & vbCrLf & _
        "            Exit Function" & vbCrLf & _
        "        End If" & vbCrLf & _
        "    End If" & vbCrLf & vbCrLf & _
        "Fallback:" & vbCrLf & _
        "    If Len(userName) = 0 Then userName = Environ$(""USERNAME"")" & vbCrLf & _
        "    GetWindowsDisplayName = userName" & vbCrLf & _
        "End Function" & vbCrLf & vbCrLf & _
        "Private Sub cmdSend_Click()" & vbCrLf & _
        "    If Me.Controls(""cmbRequestBase"").ListIndex < 0 Then" & vbCrLf & _
        "        MsgBox ""拠点を選択してください。"", vbExclamation" & vbCrLf & _
        "        Me.Controls(""cmbRequestBase"").SetFocus" & vbCrLf & _
        "        Exit Sub" & vbCrLf & _
        "    End If" & vbCrLf & vbCrLf & _
        "    MsgBox ""送信処理は未実装です。ここに処理を書いてください。"", vbInformation" & vbCrLf & _
        "End Sub" & vbCrLf

    cm.AddFromString src

End Sub

Public Sub SetRequestHeader(ByVal frm As Object, _
                            ByVal requestBase As String)

    On Error Resume Next
    frm.Controls("cmbRequestBase").Value = requestBase
    On Error GoTo 0

End Sub

Private Function S(ByVal v As Double) As Double
    S = v * UI_SCALE
End Function
