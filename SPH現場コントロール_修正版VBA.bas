Attribute VB_Name = "SPH_Control_Create_v3fix"
Option Explicit

' =========================================================
' SPHと現場コントロール_修正版3 をVBAで新規作成（色・配置修正版）
' 実行場所：Excel / PowerPoint どちらでも可
' 実行マクロ：CreateSPHControlDeck
'
' ・PowerPoint参照設定不要（遅延バインディング）
' ・16:9 / 5スライド
' ・図形、配色、レイアウトをVBAで生成
' ・外部画像/外部ファイル不要
' =========================================================

Private Const FONT_JP As String = "メイリオ"

' PowerPoint定数（参照設定なしで動かすため数値で定義）
Private Const ppLayoutBlank As Long = 12

Private Const ppAlignLeft As Long = 1
Private Const ppAlignCenter As Long = 2
Private Const ppAlignRight As Long = 3

Private Const ppAutoSizeNone As Long = 0

' =========================================================
' メイン処理
' =========================================================
Public Sub CreateSPHControlDeck()

    Dim ppApp As Object
    Dim pres As Object
    Dim sld As Object

    On Error Resume Next
    Set ppApp = GetObject(, "PowerPoint.Application")
    On Error GoTo 0

    If ppApp Is Nothing Then
        Set ppApp = CreateObject("PowerPoint.Application")
    End If

    ppApp.Visible = True

    Set pres = ppApp.Presentations.Add

    With pres.PageSetup
        .SlideWidth = PtIn(13.333333)
        .SlideHeight = PtIn(7.5)
    End With

    Do While pres.Slides.Count > 0
        pres.Slides(1).Delete
    Loop

    Set sld = pres.Slides.Add(1, ppLayoutBlank)
    BuildSlide1 sld

    Set sld = pres.Slides.Add(2, ppLayoutBlank)
    BuildSlide2 sld

    Set sld = pres.Slides.Add(3, ppLayoutBlank)
    BuildSlide3 sld

    Set sld = pres.Slides.Add(4, ppLayoutBlank)
    BuildSlide4 sld

    Set sld = pres.Slides.Add(5, ppLayoutBlank)
    BuildSlide5 sld

    pres.Slides(1).Select

    MsgBox "SPHと現場コントロール資料を作成しました。" & vbCrLf & _
           "内容を確認して、任意の場所へ保存してください。", vbInformation

    Set sld = Nothing
    Set pres = Nothing
    Set ppApp = Nothing

End Sub

Public Sub TestPowerPointConnection()

    Dim ppApp As Object

    On Error Resume Next
    Set ppApp = GetObject(, "PowerPoint.Application")
    On Error GoTo 0

    If ppApp Is Nothing Then
        On Error Resume Next
        Set ppApp = CreateObject("PowerPoint.Application")
        On Error GoTo 0
    End If

    If ppApp Is Nothing Then
        MsgBox "PowerPointを起動できませんでした。" & vbCrLf & _
               "Microsoft PowerPointがインストールされているか確認してください。", vbExclamation
        Exit Sub
    End If

    ppApp.Visible = True
    MsgBox "PowerPointへの接続はOKです。", vbInformation

    Set ppApp = Nothing

End Sub

' =========================================================
' Slide 1
' =========================================================
Private Sub BuildSlide1(ByVal sld As Object)

    ApplyBase sld, 1

    AddText sld, 0.55, 0.3, 12, 0.52, _
            "SPHと現場コントロール", 22, True, ColWhite, ppAlignLeft

    AddText sld, 0.58, 0.88, 11.83, 0.31, _
            "17時コミットと生産性目標を達成するために、残件数と配置(らぼろぐ)を適切に管理する", _
            13.5, False, ColSub, ppAlignLeft

    AddRoundBox sld, 0.75, 1.57, 5.69, 1.31, ColPanel, ColCyan, 1.5
    Dim sh As Object
    Set sh = AddText(sld, 0.89, 1.73, 5.42, 0.99, _
                     "SPH ＝ 対応件数 ÷ Aチーム稼働時間（h）" & vbCrLf & _
                     "*らぼろぐ【本体】Aチームを選択していた時間", _
                     15, True, ColWhite, ppAlignCenter)
    SetParagraphFont sh, 2, 10, True, ColWhite, ppAlignCenter

    AddRoundBox sld, 6.81, 1.57, 5.69, 1.31, ColPanel, ColOrange, 1.5
    AddText sld, 6.96, 1.73, 5.4, 0.99, _
            "SVが管理すること" & vbCrLf & _
            "サポートの投入・解除、らぼろぐの切替", _
            15, True, ColWhite, ppAlignCenter

    AddRoundBox sld, 0.75, 3.28, 2.75, 1.08, ColBlue, ColBlue, 0
    AddText sld, 0.88, 3.36, 2.5, 0.9, _
            "CL目標" & vbCrLf & "SPH 20.4", _
            16.5, True, ColWhite, ppAlignCenter

    AddRoundBox sld, 3.75, 3.28, 2.75, 1.08, ColGrayBox, ColGrayBox, 0
    AddText sld, 3.88, 3.36, 2.5, 0.9, _
            "現状のボリュームゾーン" & vbCrLf & "SPH 13〜16", _
            15.5, True, ColWhite, ppAlignCenter

    AddRoundBox sld, 6.75, 3.28, 2.75, 1.08, ColBrown, ColBrown, 0
    AddText sld, 6.88, 3.36, 2.5, 0.9, _
            "中間目標" & vbCrLf & "SPH 18", _
            16.5, True, ColWhite, ppAlignCenter

    AddRoundBox sld, 9.75, 3.28, 2.75, 1.08, ColGreen, ColGreen, 0
    AddText sld, 9.88, 3.36, 2.5, 0.9, _
            "改善の重点" & vbCrLf & _
            "個人の対応力" & vbCrLf & _
            "→SVの現場コントロール", _
            12.7, True, ColWhite, ppAlignCenter

    AddText sld, 0.75, 4.88, 11.75, 0.35, _
            "チームSPHの算出方法", 16.5, True, ColCyan, ppAlignLeft

    AddText sld, 0.75, 5.34, 11.75, 0.77, _
            "チームSPH ＝ 全員の対応件数合計 ÷ Aチーム時間合計", _
            16.5, False, ColBody, ppAlignLeft

End Sub

' =========================================================
' Slide 2
' =========================================================
Private Sub BuildSlide2(ByVal sld As Object)

    ApplyBase sld, 2

    AddText sld, 0.55, 0.3, 12, 0.52, _
            "SVの現場コントロール", 22, True, ColWhite, ppAlignLeft

    AddText sld, 0.58, 0.88, 11.83, 0.31, _
            "見わたす・予測する・判断する・締めるを、共通の手順にする", _
            13.5, False, ColSub, ppAlignLeft

    AddRoundBox sld, 0.75, 1.51, 2.27, 1.46, ColStepBlue, ColCyan, 1.5
    AddText sld, 0.83, 1.64, 2.1, 0.44, "見わたす", 20, True, ColWhite, ppAlignCenter
    AddText sld, 0.8, 2.17, 2.17, 0.59, _
            "・残件数・受信状況" & vbCrLf & "・TSRのコンディション", _
            12.2, False, ColWhite, ppAlignCenter

    AddArrowLine sld, 3.1, 2.24, 3.63, 2.24

    AddRoundBox sld, 3.75, 1.51, 2.27, 1.46, ColStepBlue, ColCyan, 1.5
    AddText sld, 3.83, 1.64, 2.1, 0.44, "予測する", 20, True, ColWhite, ppAlignCenter
    AddText sld, 3.8, 2.17, 2.17, 0.59, _
            "残り時間と対応力", 12.2, False, ColWhite, ppAlignCenter

    AddArrowLine sld, 6.1, 2.24, 6.63, 2.24

    AddRoundBox sld, 6.75, 1.51, 2.27, 1.46, ColBrown, ColOrange, 1.5
    AddText sld, 6.83, 1.64, 2.1, 0.44, "判断する", 20, True, ColWhite, ppAlignCenter
    AddText sld, 6.8, 2.17, 2.17, 0.59, _
            "投入先・人数・時間" & vbCrLf & "優先順を決定", _
            12.2, False, ColWhite, ppAlignCenter

    AddArrowLine sld, 9.1, 2.24, 9.63, 2.24

    AddRoundBox sld, 9.75, 1.51, 2.27, 1.46, ColGreen, ColGreenBright, 1.5
    AddText sld, 9.83, 1.64, 2.1, 0.44, "締める", 20, True, ColWhite, ppAlignCenter
    AddText sld, 9.8, 2.17, 2.17, 0.59, _
            "・サポート解除" & vbCrLf & "・「らぼろぐ」を余資へ", _
            12.2, False, ColWhite, ppAlignCenter

    AddRoundBox sld, 0.75, 3.38, 5.69, 2.12, ColPanel, ColRed, 1.5
    AddRoundBox sld, 0.94, 3.55, 3.45, 0.4, ColRed, ColRed, 0
    AddText sld, 1.09, 3.6, 3.22, 0.3, _
            "サポートの効果と負荷", 13.5, True, ColWhite, ppAlignLeft
    AddText sld, 0.96, 4.05, 5.29, 1.25, _
            "・自身の担当CAと同様のSPHは出にくい" & vbCrLf & _
            "└案件の選定に、時間がかかる。" & vbCrLf & _
            "・対象案件を指定することで、案件選択の時間を減らす", _
            13.2, False, ColBody, ppAlignLeft

    AddRoundBox sld, 6.81, 3.38, 5.69, 2.12, ColPanel, ColCyan, 1.5
    AddRoundBox sld, 7.02, 3.55, 3.79, 0.4, ColStepBlue, ColCyan, 1.5
    AddText sld, 7.19, 3.6, 3.46, 0.3, _
            "投入から解除までの判断", 13.5, True, ColWhite, ppAlignLeft
    AddText sld, 7.04, 4.05, 5.26, 1.25, _
            "・サポートが必要な件数(人数)を把握する" & vbCrLf & _
            "・サポートの条件を明確に伝える　*～●時までの未読など" & vbCrLf & _
            "・手隙が生まれたら「らぼろぐ」を「【その他】余資」へ", _
            12.5, False, ColBody, ppAlignLeft

    AddRoundBox sld, 0.75, 5.98, 11.75, 0.85, ColGreen, ColGreenBright, 1.5
    AddText sld, 0.96, 6.07, 11.33, 0.67, _
            "自走コミット困難が予測できた場合は、早めにサポートを投入する。" & vbCrLf & _
            "時間を空けて再判定し、サポートの継続 or 解除を判断する。", _
            15.3, True, ColWhite, ppAlignCenter

End Sub

' =========================================================
' Slide 3
' =========================================================
Private Sub BuildSlide3(ByVal sld As Object)

    ApplyBase sld, 3

    AddText sld, 0.55, 0.3, 12, 0.52, _
            "同じ144件でも、Aチーム稼働した時間でSPHは変わる", _
            22, True, ColWhite, ppAlignLeft

    AddText sld, 0.58, 0.88, 11.83, 0.31, _
            "例：9〜18時勤務、昼休憩1時間。17時に対象処理・サポートを完了した場合", _
            12.75, False, ColSub, ppAlignLeft

    AddRoundBox sld, 0.75, 1.34, 11.75, 0.49, ColPanel, ColCyan, 1.2
    AddText sld, 0.91, 1.42, 11.44, 0.32, _
            "SPH ＝ 処理件数 ÷ Aチーム稼働時間（h）　※図は昼休憩を除く8時間", _
            15.3, True, ColWhite, ppAlignCenter

    AddText sld, 0.75, 2.17, 7.19, 0.32, _
            "ケースA：17時に対象処理完了→らぼろぐ切替なし", _
            15, True, ColWhite, ppAlignLeft

    AddText sld, 0.95, 2.64, 0.52, 0.26, "0h", 11.25, False, ColSub, ppAlignLeft
    AddText sld, 6.54, 2.64, 0.48, 0.26, "7h", 11.25, False, ColSub, ppAlignCenter
    AddText sld, 7.17, 2.64, 0.46, 0.26, "8h", 11.25, False, ColSub, ppAlignRight

    AddRect sld, 0.95, 2.97, 6.67, 0.5, ColStepBlue, ColStepBlue, 0
    AddText sld, 1.04, 3.03, 6.46, 0.38, _
            "Aチーム 8時間", 15.75, True, ColWhite, ppAlignCenter

    AddRoundBox sld, 8.22, 2.91, 4.28, 0.72, ColGrayBox, ColBorder, 1
    AddText sld, 8.34, 3.03, 4.03, 0.46, _
            "144件 ÷ 8h ＝ SPH 18.0", 17.25, True, ColWhite, ppAlignCenter

    AddText sld, 0.75, 4.18, 7.19, 0.32, _
            "ケースB： 17時に対象処理完了→らぼろぐ切替あり", _
            15, True, ColWhite, ppAlignLeft

    AddText sld, 0.95, 4.66, 0.52, 0.26, "0h", 11.25, False, ColSub, ppAlignLeft
    AddText sld, 6.54, 4.66, 0.48, 0.26, "7h", 11.25, False, ColSub, ppAlignCenter
    AddText sld, 7.17, 4.66, 0.46, 0.26, "8h", 11.25, False, ColSub, ppAlignRight

    AddRect sld, 0.95, 4.99, 5.83, 0.5, ColStepBlue, ColStepBlue, 0
    AddRect sld, 6.78, 4.99, 0.83, 0.5, ColGreen, ColGreen, 0
    AddText sld, 1.04, 5.05, 5.65, 0.38, _
            "Aチーム 7時間", 15.75, True, ColWhite, ppAlignCenter
    AddText sld, 6.84, 5.05, 0.71, 0.38, _
            "余資", 14.25, True, ColWhite, ppAlignCenter

    AddRoundBox sld, 8.22, 4.93, 4.28, 0.72, ColGreen, ColGreenBright, 1.2
    AddText sld, 8.34, 5.05, 4.03, 0.46, _
            "144件 ÷ 7h ≒ SPH 20.6", 17.25, True, ColWhite, ppAlignCenter

End Sub

' =========================================================
' Slide 4
' =========================================================
Private Sub BuildSlide4(ByVal sld As Object)

    ApplyBase sld, 4

    AddText sld, 0.55, 0.3, 12, 0.52, _
            "残件数管理の目安：既存案件を1時間に16件対応すると仮定する", _
            22, True, ColWhite, ppAlignLeft

    AddText sld, 0.58, 0.88, 11.83, 0.31, _
            "昼休憩を除く1時間の試算例。", _
            13.5, False, ColSub, ppAlignLeft

    AddRoundBox sld, 0.75, 1.34, 3.38, 0.44, ColGrayBox, ColBorder, 1
    AddText sld, 0.86, 1.4, 3.15, 0.32, _
            "平均処理時間：約2.5分", 13.5, True, ColWhite, ppAlignCenter

    AddText sld, 0.75, 1.98, 5.62, 0.32, _
            "1時間＝60分の配分例", 15, True, ColWhite, ppAlignLeft

    AddRect sld, 0.83, 2.56, 7.5, 0.92, ColStepBlue, ColStepBlue, 0
    AddRect sld, 8.33, 2.56, 1.88, 0.92, ColGrayBox, ColGrayBox, 0
    AddRect sld, 10.21, 2.56, 1.88, 0.92, ColBrown, ColBrown, 0

    AddText sld, 1.02, 2.7, 7.12, 0.64, _
            "既存案件の処理" & vbCrLf & "2.5分 × 16件 ＝ 40分", _
            17.25, True, ColWhite, ppAlignCenter

    AddText sld, 8.39, 2.71, 1.77, 0.62, _
            "・10分休憩/離席" & vbCrLf & "・エスカレ等 10分", _
            12.38, True, ColWhite, ppAlignCenter

    AddText sld, 10.26, 2.71, 1.77, 0.62, _
            "・仕分け" & vbCrLf & "・至急対応 10分", _
            12.75, True, ColWhite, ppAlignCenter

    AddText sld, 4.19, 3.57, 0.79, 0.3, "40分", 12.75, False, ColSub, ppAlignCenter
    AddText sld, 8.88, 3.57, 0.79, 0.3, "10分", 12.75, False, ColSub, ppAlignCenter
    AddText sld, 10.75, 3.57, 0.79, 0.3, "10分", 12.75, False, ColSub, ppAlignCenter

    AddPlainLine sld, 0.83, 4.06, 12.08, 4.06, ColBorder, 1

    AddRoundBox sld, 4.79, 4.85, 3.54, 1.15, ColGreen, ColGreenBright, 1.2
    AddText sld, 4.92, 4.98, 3.29, 0.9, _
            "既存案件の対応目安" & vbCrLf & "16件/h" & vbCrLf & "*SPH18達成に必要な数", _
            17.25, True, ColWhite, ppAlignCenter

End Sub

' =========================================================
' Slide 5
' =========================================================
Private Sub BuildSlide5(ByVal sld As Object)

    ApplyBase sld, 5

    AddText sld, 0.55, 0.3, 12, 0.52, _
            "12時時点の残件数と、残り時間による判定例", _
            22, True, ColWhite, ppAlignLeft

    AddText sld, 0.58, 0.88, 11.83, 0.31, _
            "18時完了・休憩未取得1時間・面談等なしの例（担当者1人あたり）", _
            13.5, False, ColSub, ppAlignLeft

    AddRoundBox sld, 0.83, 1.38, 4.07, 0.92, ColGrayBox, ColGrayBox, 0
    AddText sld, 0.98, 1.48, 3.78, 0.69, _
            "12〜18時の6時間 − 休憩1時間" & vbCrLf & _
            "残り処理時間 5時間", _
            13.5, True, ColWhite, ppAlignCenter

    AddRoundBox sld, 5.31, 1.38, 3.06, 0.92, ColStepBlue, ColStepBlue, 0
    AddText sld, 5.42, 1.48, 2.85, 0.69, _
            "既存案件の処理可能目安" & vbCrLf & _
            "16件/h × 5h ＝ 80件", _
            13.5, True, ColWhite, ppAlignCenter

    AddRoundBox sld, 8.79, 1.38, 3.71, 0.92, ColAmberDark, ColAmberDark, 0
    AddText sld, 8.92, 1.48, 3.46, 0.69, _
            "5時間で80件が対応可能目安", _
            14.25, True, ColWhite, ppAlignCenter

    AddText sld, 0.83, 2.45, 11.62, 0.28, _
            "※残件数は、その時点の当日対応対象（「□自動」を除く未読・フラグを含む）", _
            11.63, False, ColSub, ppAlignLeft

    AddRoundBox sld, 0.83, 2.94, 1.19, 0.68, ColSky, ColSky, 0
    AddText sld, 0.89, 3.01, 1.08, 0.52, "80件", 17.25, True, ColWhite, ppAlignCenter

    AddRoundBox sld, 2.31, 2.94, 1.93, 0.68, ColSky, ColSky, 0
    AddText sld, 2.4, 2.99, 1.76, 0.57, _
            "16件/時間" & vbCrLf & "経過観察", _
            13.5, True, ColWhite, ppAlignCenter

    AddText sld, 4.57, 2.98, 7.93, 0.61, _
            "自走可能ライン！" & vbCrLf & _
            "昼休憩以降に新規受信が増加した場合、早めにサポート投入を検討。", _
            13.5, False, ColBody, ppAlignLeft

    AddRoundBox sld, 0.83, 3.94, 1.19, 0.68, ColYellow, ColYellow, 0
    AddText sld, 0.89, 4.01, 1.08, 0.52, "100件", 17.25, True, ColWhite, ppAlignCenter

    AddRoundBox sld, 2.31, 3.94, 1.93, 0.68, ColYellow, ColYellow, 0
    AddText sld, 2.4, 3.99, 1.76, 0.57, _
            "20件/h" & vbCrLf & "サポート検討", _
            13.5, True, ColWhite, ppAlignCenter

    AddText sld, 4.57, 3.98, 7.93, 0.61, _
            "SPH20未満のTSRは逼迫、自走が難しくなるライン。サポートを投入。" & vbCrLf & _
            "ハイパー層(SPH20以上)は個別にサポート投入要否を判断。", _
            13.5, False, ColBody, ppAlignLeft

    AddRoundBox sld, 0.83, 4.94, 1.19, 0.68, ColRed, ColRed, 0
    AddText sld, 0.89, 5.01, 1.08, 0.52, "120件", 17.25, True, ColWhite, ppAlignCenter

    AddRoundBox sld, 2.31, 4.94, 1.93, 0.68, ColRed, ColRed, 0
    AddText sld, 2.4, 4.99, 1.76, 0.57, _
            "24件/h" & vbCrLf & "原則サポート", _
            13.5, True, ColWhite, ppAlignCenter

    AddText sld, 4.57, 4.98, 7.93, 0.61, _
            "「□自動」を除く「純粋な未読」が50件以上ある場合、対応遅延のリスクあり！" & vbCrLf & _
            "スキルに関わらず、優先的にサポートを投入！", _
            13.5, False, ColBody, ppAlignLeft

    AddRoundBox sld, 0.83, 6.05, 11.67, 0.78, ColPanel, ColRed, 1.2
    AddText sld, 1, 6.14, 11.33, 0.61, _
            "全体的に逼迫しており、サポート投入が難しい場合は、統括・LSVへアラート出し！", _
            15, True, ColWhite, ppAlignCenter

End Sub

' =========================================================
' 共通部品
' =========================================================
Private Sub ApplyBase(ByVal sld As Object, ByVal slideNo As Long)

    sld.FollowMasterBackground = False
    With sld.Background.Fill
        .Visible = True
        .Solid
        .ForeColor.RGB = ColBg
    End With

    AddRect sld, 0, 0, 13.333333, 0.05, ColCyan, ColCyan, 0
    AddRect sld, 0, 7.43, 13.333333, 0.07, ColOrange, ColOrange, 0

    AddText sld, 12.38, 7.05, 0.31, 0.24, _
            CStr(slideNo), 8.25, False, ColSub, ppAlignRight

End Sub

Private Function AddText(ByVal sld As Object, _
                         ByVal x As Double, ByVal y As Double, _
                         ByVal w As Double, ByVal h As Double, _
                         ByVal txt As String, ByVal fontSize As Double, _
                         ByVal isBold As Boolean, ByVal fontColor As Long, _
                         ByVal align As Long) As Object

    Dim sh As Object
    Set sh = sld.Shapes.AddTextbox(1, PtIn(x), PtIn(y), PtIn(w), PtIn(h))

    With sh
        .Fill.Visible = False
        .Line.Visible = False

        With .TextFrame
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = 0
            .WordWrap = True
            .AutoSize = ppAutoSizeNone
            .VerticalAnchor = 3

            .TextRange.Text = txt
            .TextRange.ParagraphFormat.Alignment = align

            With .TextRange.Font
                .Name = FONT_JP
                On Error Resume Next
                .NameFarEast = FONT_JP
                On Error GoTo 0
                .Size = fontSize
                .Bold = isBold
                .Color.RGB = fontColor
            End With
        End With
    End With

    Set AddText = sh

End Function

Private Sub SetParagraphFont(ByVal sh As Object, _
                             ByVal paragraphIndex As Long, _
                             ByVal fontSize As Double, _
                             ByVal isBold As Boolean, _
                             ByVal fontColor As Long, _
                             ByVal align As Long)

    On Error Resume Next
    With sh.TextFrame.TextRange.Paragraphs(paragraphIndex)
        .ParagraphFormat.Alignment = align
        With .Font
            .Name = FONT_JP
            .NameFarEast = FONT_JP
            .Size = fontSize
            .Bold = isBold
            .Color.RGB = fontColor
        End With
    End With
    On Error GoTo 0

End Sub

Private Function AddRect(ByVal sld As Object, _
                         ByVal x As Double, ByVal y As Double, _
                         ByVal w As Double, ByVal h As Double, _
                         ByVal fillColor As Long, ByVal lineColor As Long, _
                         ByVal lineWeight As Double) As Object

    Dim sh As Object
    Set sh = sld.Shapes.AddShape(1, PtIn(x), PtIn(y), PtIn(w), PtIn(h))

    With sh
        .Fill.Solid
        .Fill.ForeColor.RGB = fillColor

        If lineWeight <= 0 Then
            .Line.Visible = False
        Else
            .Line.Visible = True
            .Line.ForeColor.RGB = lineColor
            .Line.Weight = lineWeight
        End If
    End With

    Set AddRect = sh

End Function

Private Function AddRoundBox(ByVal sld As Object, _
                             ByVal x As Double, ByVal y As Double, _
                             ByVal w As Double, ByVal h As Double, _
                             ByVal fillColor As Long, ByVal lineColor As Long, _
                             ByVal lineWeight As Double) As Object

    Dim sh As Object
    Set sh = sld.Shapes.AddShape(5, PtIn(x), PtIn(y), PtIn(w), PtIn(h))

    With sh
        .Fill.Solid
        .Fill.ForeColor.RGB = fillColor

        If lineWeight <= 0 Then
            .Line.Visible = False
        Else
            .Line.Visible = True
            .Line.ForeColor.RGB = lineColor
            .Line.Weight = lineWeight
        End If
    End With

    Set AddRoundBox = sh

End Function

Private Sub AddArrowLine(ByVal sld As Object, _
                         ByVal x1 As Double, ByVal y1 As Double, _
                         ByVal x2 As Double, ByVal y2 As Double)

    Dim sh As Object
    Set sh = sld.Shapes.AddConnector(1, PtIn(x1), PtIn(y1), PtIn(x2), PtIn(y2))

    With sh.Line
        .Visible = True
        .ForeColor.RGB = ColArrow
        .Weight = 1.7
        .BeginArrowheadStyle = 1
        .EndArrowheadStyle = 3
    End With

End Sub

Private Sub AddPlainLine(ByVal sld As Object, _
                         ByVal x1 As Double, ByVal y1 As Double, _
                         ByVal x2 As Double, ByVal y2 As Double, _
                         ByVal color As Long, ByVal weight As Double)

    Dim sh As Object
    Set sh = sld.Shapes.AddLine(PtIn(x1), PtIn(y1), PtIn(x2), PtIn(y2))

    With sh.Line
        .Visible = True
        .ForeColor.RGB = color
        .Weight = weight
    End With

End Sub

Private Function PtIn(ByVal inches As Double) As Single
    PtIn = CSng(inches * 72)
End Function

' =========================================================
' 色
' =========================================================
Private Function ColBg() As Long
    ColBg = RGB(13, 23, 40)
End Function

Private Function ColPanel() As Long
    ColPanel = RGB(17, 29, 49)
End Function

Private Function ColCyan() As Long
    ColCyan = RGB(51, 181, 229)
End Function

Private Function ColOrange() As Long
    ColOrange = RGB(243, 155, 23)
End Function

Private Function ColWhite() As Long
    ColWhite = RGB(255, 255, 255)
End Function

Private Function ColSub() As Long
    ColSub = RGB(182, 196, 216)
End Function

Private Function ColBody() As Long
    ColBody = RGB(220, 229, 243)
End Function

Private Function ColBlue() As Long
    ColBlue = RGB(27, 118, 160)
End Function

Private Function ColStepBlue() As Long
    ColStepBlue = RGB(14, 99, 139)
End Function

Private Function ColGrayBox() As Long
    ColGrayBox = RGB(42, 50, 67)
End Function

Private Function ColBrown() As Long
    ColBrown = RGB(163, 71, 11)
End Function

Private Function ColGreen() As Long
    ColGreen = RGB(18, 107, 53)
End Function

Private Function ColGreenBright() As Long
    ColGreenBright = RGB(34, 197, 94)
End Function

Private Function ColRed() As Long
    ColRed = RGB(179, 40, 40)
End Function

Private Function ColBorder() As Long
    ColBorder = RGB(51, 65, 88)
End Function

Private Function ColArrow() As Long
    ColArrow = RGB(207, 214, 230)
End Function

Private Function ColAmber() As Long
    ColAmber = RGB(255, 179, 59)
End Function

Private Function ColAmberDark() As Long
    ColAmberDark = RGB(138, 74, 16)
End Function

Private Function ColSky() As Long
    ColSky = RGB(0, 176, 240)
End Function

Private Function ColYellow() As Long
    ColYellow = RGB(255, 192, 0)
End Function

Private Function ColOrange2() As Long
    ColOrange2 = RGB(212, 122, 29)
End Function
