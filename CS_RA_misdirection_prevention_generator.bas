Attribute VB_Name = "modCSRAMisdirectionDeck"
Option Explicit

' ============================================================
' CSにRA宛内容誤送信防止 研修資料 自動生成マクロ
' Excel VBA の「標準モジュール」に貼り付けて実行してください。
' PowerPointへの参照設定は不要（Late Binding）。
' ============================================================

' ----- Office / PowerPoint constants (Late Binding用) -----
Private Const msoFalse As Long = 0
Private Const msoTrue As Long = -1
Private Const msoTextOrientationHorizontal As Long = 1
Private Const msoShapeRectangle As Long = 1
Private Const msoShapeOval As Long = 9
Private Const msoAnchorMiddle As Long = 3
Private Const ppLayoutBlank As Long = 12
Private Const ppAlignLeft As Long = 1
Private Const ppAlignCenter As Long = 2
Private Const ppAlignRight As Long = 3

Private Const FONT_NAME As String = "メイリオ"

' ============================================================
' 実行入口
' ============================================================
Public Sub CS_RA誤送信防止資料を作成()
    BuildDeck "", ""
End Sub

' 2ページ目のスクリーンショットも同時に貼りたい場合はこちら。
Public Sub CS_RA誤送信防止資料を作成_スクショ指定()
    Dim imgPath As String
    imgPath = InputBox( _
        "2ページ目に貼り付ける検索画面スクリーンショットのフルパスを入力してください。" & vbCrLf & _
        "空欄ならプレースホルダーのまま作成します。", _
        "スクリーンショット指定")

    BuildDeck imgPath, ""
End Sub

' 保存先も指定したい場合のサンプル。
Public Sub CS_RA誤送信防止資料を作成して保存()
    Dim savePath As String
    savePath = ThisWorkbook.Path & "\CS_RA誤送信防止_自動生成.pptx"
    BuildDeck "", savePath
End Sub

' ============================================================
' メイン
' ============================================================
Private Sub BuildDeck(ByVal screenshotPath As String, ByVal savePath As String)
    Dim ppApp As Object
    Dim pres As Object
    Dim sl As Object

    On Error Resume Next
    Set ppApp = GetObject(, "PowerPoint.Application")
    On Error GoTo 0

    If ppApp Is Nothing Then
        Set ppApp = CreateObject("PowerPoint.Application")
    End If

    ppApp.Visible = True
    Set pres = ppApp.Presentations.Add

    ' 16:9 / ワイド
    pres.PageSetup.SlideWidth = Pt(13.333333)
    pres.PageSetup.SlideHeight = Pt(7.5)

    ' --- Slide 1 ---
    Set sl = pres.Slides.Add(1, ppLayoutBlank)
    BuildSlide1 sl

    ' --- Slide 2 ---
    Set sl = pres.Slides.Add(2, ppLayoutBlank)
    BuildSlide2 sl, screenshotPath

    If Len(savePath) > 0 Then
        pres.SaveAs savePath
        MsgBox "PowerPointを作成しました。" & vbCrLf & savePath, vbInformation
    Else
        MsgBox "PowerPointを作成しました。" & vbCrLf & _
               "内容を確認して、必要に応じて保存してください。", vbInformation
    End If
End Sub

' ============================================================
' Slide 1 : 原因①
' ============================================================
Private Sub BuildSlide1(ByVal sl As Object)
    AddCommonHeader sl, _
        "原因①｜RAのメールを選び直していない", _
        "RA宛ではないメールが管理画面に残ることがあるため、メール作成前に送りたい相手のメールを選び直す"

    ' ヘッダー副題の先頭を赤強調
    EmphasizeText sl.Shapes("subtitle"), "RA宛ではないメールが管理画面に残る", C_Red(), True

    AddSectionTitle sl, 0.562, 1.667, 6.25, "■どこで工程が抜けたのか", "cause-title"
    AddOutlinePanel sl, 0.562, 2.104, 6.25, 2.167, C_GrayLine(), 1.5, C_White(), "cause-panel"

    AddText sl, 0.792, 2.229, 3.646, 0.271, _
            "(例)｜CSの完了報告を、RAへ送信する", 12, C_Dark(), True, ppAlignLeft, "cause-panel-label"
    AddText sl, 4.75, 2.229, 1.812, 0.271, _
            "RAメールの選択が必要", 12, C_Red(), True, ppAlignRight, "mixed-label"
    AddText sl, 2.558, 2.62, 2.26, 0.312, _
            "CSからのメール内容を確認", 12, C_GrayText(), False, ppAlignLeft, "search-strip-text"

    AddChoiceRow sl, 0.792, 3.01, "正", "RA宛メールを選択　→　作成", True
    AddChoiceRow sl, 0.792, 3.521, "誤", "表示されているCSのメールにそのまま返信", False

    AddText sl, 0.792, 3.998, 5.792, 0.25, _
            "CSからの完了報告メールにそのまま返信すると、【既読】→【対応漏れ】のリスク！！", _
            12, C_Red(), True, ppAlignCenter, "cause-result"

    ' 誤った流れ
    AddOutlinePanel sl, 0.562, 4.458, 6.25, 1.589, C_DarkRed(), 1.5, C_LightRed(), "filter-panel"
    AddText sl, 0.792, 4.552, 3.958, 0.312, "誤った流れ", 15.75, C_DarkRed(), True, ppAlignLeft, "filter-title"
    AddText sl, 0.792, 4.865, 5.208, 0.26, _
            "CSのメールを返信元にしたまま、RA宛として作成", 12, C_GrayText(), False, ppAlignLeft, "filter-note"

    AddText sl, 0.792, 5.167, 0.938, 0.229, "内容確認", 12, C_GrayText(), True, ppAlignLeft, "sender-label"
    AddOutlinePanel sl, 0.792, 5.427, 1.979, 0.469, C_DarkRed(), 1.1, C_White(), "sender-input"
    AddText sl, 0.958, 5.474, 1.646, 0.448, "CSからの内容", 17, C_DarkRed(), True, ppAlignLeft, "sender-domain"

    AddText sl, 2.979, 5.167, 0.938, 0.229, "返信元", 12, C_GrayText(), True, ppAlignLeft, "recipient-label"
    AddOutlinePanel sl, 2.979, 5.427, 1.979, 0.469, C_DarkRed(), 1.1, C_White(), "recipient-input"
    AddText sl, 3.13, 5.458, 1.74, 0.448, "CSからのメール", 15, C_DarkRed(), True, ppAlignLeft, "recipient-domain"

    AddText sl, 5.024, 5.455, 0.469, 0.5, "→", 21, C_DarkRed(), True, ppAlignCenter, "filter-arrow"
    AddText sl, 5.396, 5.442, 1.384, 0.492, "RA宛で作成", 15, C_DarkRed(), True, ppAlignCenter, "filter-result"

    ' 右側：本来のフロー
    AddSectionTitle sl, 7.292, 1.667, 5.479, "■本来行うべきフロー", "flow-title"
    AddStepFlow sl, _
        "RAのメールを選択", "作成前に、CSのメッセージ検索画面でRA宛メールを選び直す", _
        "メッセージ管理画面でRA宛か確認", "選択したRA宛メールが表示されているか確認", _
        "返信／再送でメールを作成", "RA宛であることを確認してからボタンを押す", _
        "送信前チェック後に送付", "宛先・本文・添付を確認する"

    AddBottomBanner sl, "必須動作", "メール作成前に、必ずRA宛メールを選び直す"
End Sub

' ============================================================
' Slide 2 : 原因②
' ============================================================
Private Sub BuildSlide2(ByVal sl As Object, ByVal screenshotPath As String)
    AddCommonHeader sl, _
        "原因②｜RA宛メールを選び間違えている", _
        "RA／CSのメールが混在する検索結果から、CSのメールを誤って選択し、そのまま作成している"

    EmphasizeText sl.Shapes("subtitle"), "CSのメールを誤って選択", C_Red(), True

    AddSectionTitle sl, 0.562, 1.667, 6.25, "■選択ミスが起きる場面", "cause-title"
    AddOutlinePanel sl, 0.562, 2.104, 6.25, 2.167, C_GrayLine(), 1.5, C_White(), "cause-panel"

    AddText sl, 0.792, 2.229, 3.646, 0.271, _
            "(例)｜CSの完了報告を、RAへ送信する", 12, C_Dark(), True, ppAlignLeft, "cause-panel-label"
    AddText sl, 3.962, 2.229, 2.671, 0.271, _
            "RA／CSメールが混在している場合", 11, C_Red(), True, ppAlignRight, "mixed-label"
    AddText sl, 2.551, 2.419, 2.274, 0.312, _
            "CSからのメール内容を確認", 12, C_GrayText(), False, ppAlignLeft, "search-strip-text"

    ' 「誤→正→誤」の並び
    AddChoiceRow sl, 0.792, 2.741, "誤", "CSのメール", False
    AddChoiceRow sl, 0.792, 3.249, "正", "RAのメール", True
    AddChoiceRow sl, 0.792, 3.759, "誤", "CSのメール", False

    ' 左下：ドメイン絞り込み対策
    AddOutlinePanel sl, 0.562, 4.448, 6.25, 1.812, C_Blue(), 1.5, C_VeryLightBlue(), "filter-panel"
    AddText sl, 0.792, 4.552, 5.792, 0.312, _
            "絶対に選択ミスをしないための対策｜RA宛メールだけに絞り込む", _
            15.75, C_Blue(), True, ppAlignLeft, "filter-title"
    AddText sl, 0.792, 4.865, 5.208, 0.26, _
            "送信者・受信者の両方にRAのドメインを入力", _
            12, C_GrayText(), False, ppAlignLeft, "filter-note"

    AddScreenshotArea sl, 0.792, 5.208, 5.792, 0.854, screenshotPath

    ' 右側：検索・選択のポイント
    AddSectionTitle sl, 7.292, 1.667, 5.479, "■対策｜検索・選択のポイント", "flow-title"
    AddStepFlow sl, _
        "送信者に「@～」を入力", "RAが使用する実ドメインを入力", _
        "受信者にも「@～」を入力", "送信者と同じRAのドメインを入力", _
        "検索結果を確認", "RA宛またはRAからのメールだけ表示", _
        "RA宛メールを選択", "管理画面でもRA宛メールか確認"

    AddBottomBanner sl, "追加対策", "送信者・受信者の両方に、同じRAドメインを入力"
End Sub

' ============================================================
' 共通ヘッダー
' ============================================================
Private Sub AddCommonHeader(ByVal sl As Object, ByVal titleText As String, ByVal subText As String)
    AddFilledRect sl, 0.562, 0.354, 0.073, 0.792, C_Red(), "header-accent"

    AddText sl, 0.792, 0.312, 5, 0.271, _
            "CSにRA宛内容誤送信防止", 12, C_Red(), True, ppAlignLeft, "eyebrow"

    AddText sl, 0.792, 0.573, 11.667, 0.604, _
            titleText, 30, C_Dark(), True, ppAlignLeft, "title"

    AddText sl, 0.792, 1.167, 11.667, 0.312, _
            subText, 12.75, C_GrayText(), False, ppAlignLeft, "subtitle"

    AddFilledRect sl, 0.562, 1.542, 12.208, 0.016, C_GrayLine(), "header-rule"
End Sub

Private Sub AddSectionTitle(ByVal sl As Object, ByVal x As Double, ByVal y As Double, _
                            ByVal w As Double, ByVal txt As String, ByVal shpName As String)
    Dim shp As Object
    Set shp = AddText(sl, x, y, w, 0.375, txt, 18, C_Dark(), True, ppAlignLeft, shpName)
    EmphasizeText shp, "■", C_Red(), True
End Sub

' ============================================================
' 正／誤 行
' ============================================================
Private Sub AddChoiceRow(ByVal sl As Object, ByVal x As Double, ByVal y As Double, _
                         ByVal chipText As String, ByVal bodyText As String, ByVal isCorrect As Boolean)
    Dim borderColor As Long
    Dim fillColor As Long
    Dim textColor As Long
    Dim prefix As String

    If isCorrect Then
        borderColor = C_Blue()
        fillColor = C_White()
        textColor = C_Dark()
        prefix = "ok"
    Else
        borderColor = C_DarkRed()
        fillColor = C_LightRed()
        textColor = C_DarkRed()
        prefix = "ng"
    End If

    AddOutlinePanel sl, x, y, 5.792, 0.438, borderColor, 1.2, fillColor, prefix & "-row-" & Format(y, "0.000")
    AddCircle sl, x + 0.125, y + 0.055, 0.292, borderColor, borderColor, 0, prefix & "-chip-" & Format(y, "0.000")
    AddText sl, x + 0.125, y + 0.073, 0.292, 0.292, chipText, 12, C_White(), True, ppAlignCenter, prefix & "-chiptext-" & Format(y, "0.000")
    AddText sl, x + 0.562, y + 0.021, 4.85, 0.396, bodyText, 12.75, textColor, True, ppAlignLeft, prefix & "-body-" & Format(y, "0.000")
End Sub

' ============================================================
' 右側4ステップ
' ============================================================
Private Sub AddStepFlow(ByVal sl As Object, _
                        ByVal t1 As String, ByVal b1 As String, _
                        ByVal t2 As String, ByVal b2 As String, _
                        ByVal t3 As String, ByVal b3 As String, _
                        ByVal t4 As String, ByVal b4 As String)
    Dim ln As Object

    Set ln = sl.Shapes.AddLine(Pt(7.542), Pt(2.333), Pt(7.542), Pt(5.75))
    With ln.Line
        .ForeColor.RGB = C_GrayLine()
        .Weight = 1.5
    End With
    ln.Name = "step-connector"

    ' step 1 highlight
    AddFilledRect sl, 7.208, 2.125, 5.562, 0.896, C_LightBlue(), "step-one-highlight"

    AddOneStep sl, 1, 2.208, t1, b1, True
    AddOneStep sl, 2, 3.188, t2, b2, False
    AddOneStep sl, 3, 4.167, t3, b3, False
    AddOneStep sl, 4, 5.146, t4, b4, False
End Sub

Private Sub AddOneStep(ByVal sl As Object, ByVal stepNo As Long, ByVal y As Double, _
                       ByVal titleText As String, ByVal bodyText As String, ByVal highlighted As Boolean)
    Dim fillColor As Long
    Dim titleColor As Long

    If highlighted Then
        fillColor = C_Blue()
        titleColor = C_Blue()
    Else
        fillColor = C_White()
        titleColor = C_Dark()
    End If

    AddCircle sl, 7.312, y, 0.458, fillColor, C_Blue(), 1.5, "step-" & stepNo & "-circle"
    AddText sl, 7.312, y, 0.458, 0.458, CStr(stepNo), 15, IIf(highlighted, C_White(), C_Blue()), True, ppAlignCenter, "step-" & stepNo & "-number"
    AddText sl, 7.979, y - 0.021, 4.562, 0.312, titleText, 15, titleColor, True, ppAlignLeft, "step-" & stepNo & "-title"
    AddText sl, 7.979, y + 0.313, 4.562, 0.354, bodyText, 12, C_GrayText(), False, ppAlignLeft, "step-" & stepNo & "-body"
End Sub

' ============================================================
' 下部注意帯
' ============================================================
Private Sub AddBottomBanner(ByVal sl As Object, ByVal keyText As String, ByVal ruleText As String)
    AddFilledRect sl, 0.562, 6.458, 12.208, 0.75, C_LightBlue(), "bottom-banner"
    AddText sl, 0.854, 6.601, 3.438, 0.51, keyText, 19.5, C_Red(), True, ppAlignLeft, "banner-key"
    AddFilledRect sl, 4.479, 6.625, 0.021, 0.417, C_Red(), "banner-divider"
    AddText sl, 4.729, 6.625, 7.604, 0.479, ruleText, 14.25, C_Red(), True, ppAlignLeft, "banner-rule"
End Sub

' ============================================================
' スクリーンショット枠
' ============================================================
Private Sub AddScreenshotArea(ByVal sl As Object, ByVal x As Double, ByVal y As Double, _
                              ByVal w As Double, ByVal h As Double, ByVal screenshotPath As String)
    Dim frame As Object
    Dim pic As Object

    Set frame = AddOutlinePanel(sl, x, y, w, h, C_Blue(), 1.1, C_White(), "screenshot-frame")

    If Len(screenshotPath) > 0 And Len(Dir(screenshotPath)) > 0 Then
        Set pic = sl.Shapes.AddPicture(screenshotPath, msoFalse, msoTrue, Pt(x + 0.04), Pt(y + 0.04), Pt(w - 0.08), Pt(h - 0.08))
        pic.Name = "search-screenshot"
    Else
        AddText sl, x + 0.146, y + 0.188, w - 0.292, 0.479, _
                "ここに検索欄のスクリーンショットを貼付", _
                13.5, C_GrayText(), True, ppAlignCenter, "screenshot-placeholder"
    End If
End Sub

' ============================================================
' 基本図形ヘルパー
' ============================================================
Private Function AddText(ByVal sl As Object, ByVal x As Double, ByVal y As Double, _
                         ByVal w As Double, ByVal h As Double, ByVal txt As String, _
                         ByVal fontSize As Double, ByVal fontColor As Long, _
                         ByVal isBold As Boolean, ByVal align As Long, _
                         ByVal shpName As String) As Object
    Dim shp As Object

    Set shp = sl.Shapes.AddTextbox(msoTextOrientationHorizontal, Pt(x), Pt(y), Pt(w), Pt(h))
    shp.Name = shpName

    With shp.TextFrame
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .WordWrap = msoTrue
        .VerticalAnchor = msoAnchorMiddle

        With .TextRange
            .Text = txt
            .ParagraphFormat.Alignment = align
            With .Font
                .Name = FONT_NAME
                .Size = fontSize
                .Bold = IIf(isBold, msoTrue, msoFalse)
                .Color.RGB = fontColor
            End With
        End With
    End With

    shp.Line.Visible = msoFalse
    shp.Fill.Visible = msoFalse

    Set AddText = shp
End Function

Private Function AddOutlinePanel(ByVal sl As Object, ByVal x As Double, ByVal y As Double, _
                                 ByVal w As Double, ByVal h As Double, ByVal lineColor As Long, _
                                 ByVal lineWeight As Double, ByVal fillColor As Long, _
                                 ByVal shpName As String) As Object
    Dim shp As Object

    Set shp = sl.Shapes.AddShape(msoShapeRectangle, Pt(x), Pt(y), Pt(w), Pt(h))
    shp.Name = shpName

    With shp.Fill
        .Visible = msoTrue
        .Solid
        .ForeColor.RGB = fillColor
    End With

    With shp.Line
        .Visible = msoTrue
        .ForeColor.RGB = lineColor
        .Weight = lineWeight
    End With

    Set AddOutlinePanel = shp
End Function

Private Function AddFilledRect(ByVal sl As Object, ByVal x As Double, ByVal y As Double, _
                               ByVal w As Double, ByVal h As Double, ByVal fillColor As Long, _
                               ByVal shpName As String) As Object
    Dim shp As Object

    Set shp = sl.Shapes.AddShape(msoShapeRectangle, Pt(x), Pt(y), Pt(w), Pt(h))
    shp.Name = shpName

    With shp.Fill
        .Visible = msoTrue
        .Solid
        .ForeColor.RGB = fillColor
    End With

    shp.Line.Visible = msoFalse
    Set AddFilledRect = shp
End Function

Private Function AddCircle(ByVal sl As Object, ByVal x As Double, ByVal y As Double, _
                           ByVal size As Double, ByVal fillColor As Long, _
                           ByVal lineColor As Long, ByVal lineWeight As Double, _
                           ByVal shpName As String) As Object
    Dim shp As Object

    Set shp = sl.Shapes.AddShape(msoShapeOval, Pt(x), Pt(y), Pt(size), Pt(size))
    shp.Name = shpName

    With shp.Fill
        .Visible = msoTrue
        .Solid
        .ForeColor.RGB = fillColor
    End With

    If lineWeight > 0 Then
        With shp.Line
            .Visible = msoTrue
            .ForeColor.RGB = lineColor
            .Weight = lineWeight
        End With
    Else
        shp.Line.Visible = msoFalse
    End If

    Set AddCircle = shp
End Function

Private Sub EmphasizeText(ByVal shp As Object, ByVal targetText As String, _
                          ByVal fontColor As Long, ByVal makeBold As Boolean)
    Dim p As Long
    Dim tr As Object

    p = InStr(1, shp.TextFrame.TextRange.Text, targetText, vbTextCompare)
    If p = 0 Then Exit Sub

    Set tr = shp.TextFrame.TextRange.Characters(p, Len(targetText))
    tr.Font.Color.RGB = fontColor
    tr.Font.Bold = IIf(makeBold, msoTrue, msoFalse)
End Sub

Private Function Pt(ByVal inches As Double) As Single
    Pt = CSng(inches * 72#)
End Function

' ============================================================
' 色
' ============================================================
Private Function C_Dark() As Long
    C_Dark = RGB(17, 24, 39)          ' #111827
End Function

Private Function C_GrayText() As Long
    C_GrayText = RGB(95, 107, 122)    ' #5F6B7A
End Function

Private Function C_GrayLine() As Long
    C_GrayLine = RGB(201, 208, 216)   ' #C9D0D8
End Function

Private Function C_Red() As Long
    C_Red = RGB(255, 0, 0)            ' #FF0000
End Function

Private Function C_DarkRed() As Long
    C_DarkRed = RGB(217, 45, 32)      ' #D92D20
End Function

Private Function C_Blue() As Long
    C_Blue = RGB(22, 103, 217)        ' #1667D9
End Function

Private Function C_LightRed() As Long
    C_LightRed = RGB(255, 240, 238)   ' #FFF0EE
End Function

Private Function C_LightBlue() As Long
    C_LightBlue = RGB(234, 243, 255)  ' #EAF3FF
End Function

Private Function C_VeryLightBlue() As Long
    C_VeryLightBlue = RGB(239, 246, 255) ' #EFF6FF
End Function

Private Function C_White() As Long
    C_White = RGB(255, 255, 255)
End Function
