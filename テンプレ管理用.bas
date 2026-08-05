Attribute VB_Name = "modTemplateManagerInstaller"
Option Explicit

Private Const FORM_NAME As String = "frmTemplateManager"
Private Const STANDARD_SCALE As Double = 1#
Private Const LARGE_MONITOR_SCALE As Double = 1.25
Private mFormScale As Double

'============================================================
' テンプレート管理フォームを自動作成
'============================================================
Public Sub テンプレート管理フォームを自動作成()

    CreateTemplateManagerForm STANDARD_SCALE

End Sub


'============================================================
' 24インチモニター向け（125%）の管理フォームを自動作成
'============================================================
Public Sub テンプレート管理フォームを24インチ用に自動作成()

    CreateTemplateManagerForm LARGE_MONITOR_SCALE

End Sub


'============================================================
' 指定倍率でテンプレート管理フォームを自動作成
'============================================================
Private Sub CreateTemplateManagerForm(ByVal formScale As Double)

    Const vbext_ct_MSForm As Long = 3

    Dim vbProject As Object
    Dim vbComponent As Object
    Dim designer As Object
    Dim oldComponent As Object
    Dim answer As VbMsgBoxResult
    Dim currentStep As String
    Dim controlIndex As Long

    On Error GoTo TrustError

    mFormScale = formScale

    currentStep = "VBAプロジェクトの取得"
    Set vbProject = ThisWorkbook.VBProject

    On Error Resume Next
    Set oldComponent = vbProject.VBComponents(FORM_NAME)
    On Error GoTo TrustError

    If Not oldComponent Is Nothing Then
        answer = MsgBox( _
            "既にテンプレート管理フォームが存在します。" & vbCrLf & _
            "既存フォームの中身を作り直しますか？", _
            vbQuestion + vbYesNo + vbDefaultButton2, _
            "フォーム再作成")

        If answer <> vbYes Then Exit Sub

        currentStep = "既存フォームの取得"
        Set vbComponent = oldComponent
        Set designer = vbComponent.Designer

        currentStep = "既存フォームの部品を初期化"
        For controlIndex = designer.Controls.Count - 1 To 0 Step -1
            designer.Controls.Remove designer.Controls.Item(controlIndex).Name
        Next controlIndex

        currentStep = "既存フォームのコードを初期化"
        With vbComponent.CodeModule
            If .CountOfLines > 0 Then
                .DeleteLines 1, .CountOfLines
            End If
        End With
    Else
        currentStep = "ユーザーフォームの追加"
        Set vbComponent = vbProject.VBComponents.Add(vbext_ct_MSForm)
        vbComponent.Name = FORM_NAME

        currentStep = "フォーム編集画面の取得"
        Set designer = vbComponent.Designer
    End If

    currentStep = "フォーム本体の設定"
    With vbComponent
        .Properties("Caption").Value = "テンプレート管理"
        .Properties("Width").Value = 800 * mFormScale
        .Properties("Height").Value = 480 * mFormScale
        .Properties("BackColor").Value = RGB(245, 247, 250)
        .Properties("StartUpPosition").Value = 1
    End With

    currentStep = "件名欄の配置"
    AddLabel designer, "lblColumnB", "件名", 18, 18, 300, 18
    AddTextBox designer, "txtColumnB", 18, 37, 465, 24, False

    currentStep = "本文欄の配置"
    AddLabel designer, "lblColumnC", "本文", 18, 75, 300, 18
    AddTextBox designer, "txtColumnC", 18, 94, 465, 238, True

    currentStep = "タグ欄の配置"
    AddLabel designer, "lblColumnD", "タグ", 18, 347, 300, 18
    AddTextBox designer, "txtColumnD", 18, 366, 465, 24, False

    currentStep = "検索欄の配置"
    AddLabel designer, "lblSearch", "キーワード検索", 510, 18, 270, 18
    AddTextBox designer, "txtSearch", 510, 37, 270, 24, False

    currentStep = "検索結果リストの配置"
    AddLabel designer, "lblResults", "検索結果", 510, 75, 270, 18
    AddListBox designer, "lstResults", 510, 94, 270, 304

    currentStep = "ボタンの配置"
    AddButton designer, "btnRegister", "登録", 18, 410, 85, 32, RGB(37, 99, 235)
    AddButton designer, "btnDelete", "削除", 113, 410, 85, 32, RGB(220, 38, 38)
    AddButton designer, "btnClear", "クリア", 208, 410, 85, 32, RGB(100, 116, 139)
    AddButton designer, "btnExport", "エクスポート", 303, 410, 85, 32, RGB(5, 150, 105)
    AddButton designer, "btnCopyOriginal", "原本を反映", 398, 410, 85, 32, RGB(124, 58, 237)

    currentStep = "フォーム処理コードの登録"
    vbComponent.CodeModule.AddFromString GetFormCode()

    currentStep = "ブックの保存"
    ThisWorkbook.Save

    MsgBox _
        IIf(mFormScale > STANDARD_SCALE, _
            "24インチモニター用テンプレート管理フォームを作成しました。", _
            "テンプレート管理フォームを作成しました。") & vbCrLf & _
        "続けてフォームを開きます。", _
        vbInformation

    テンプレート管理フォームを開く
    Exit Sub

TrustError:
    MsgBox _
        "フォームを自動作成できませんでした。" & vbCrLf & vbCrLf & _
        "Excelの次の設定を確認してください。" & vbCrLf & _
        "ファイル → オプション → トラストセンター" & vbCrLf & _
        "→ トラストセンターの設定 → マクロの設定" & vbCrLf & _
        "→「VBAプロジェクト オブジェクトモデルへのアクセスを信頼する」" & vbCrLf & vbCrLf & _
        "処理箇所：" & currentStep & vbCrLf & _
        "エラー番号：" & Err.Number & vbCrLf & _
        "エラー内容：" & Err.Description, _
        vbExclamation, _
        "設定確認"

End Sub


'============================================================
' 作成済みフォームを開く
'============================================================
Public Sub テンプレート管理フォームを開く()

    Dim frm As Object

    On Error GoTo FormNotFound

    Set frm = VBA.UserForms.Add(FORM_NAME)
    frm.Show
    Exit Sub

FormNotFound:
    MsgBox _
        "テンプレート管理フォームがまだ作成されていません。" & vbCrLf & _
        "先に「テンプレート管理フォームを自動作成」を実行してください。", _
        vbExclamation

End Sub


'============================================================
' フォームのエクスポートボタンから呼び出すJSON出力処理
' 原本シート：A=ID / B=件名 / C=本文 / D=タグ
'============================================================
Public Sub ExportTemplatesJsonFromForm()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim outPath As String
    Dim jsonText As String
    Dim rowNumber As Long
    Dim isFirst As Boolean
    Dim subjectText As String
    Dim bodyText As String
    Dim tagsText As String

    On Error GoTo ExportError

    Set ws = ThisWorkbook.Worksheets("原本")

    '本文が入るC列を基準に最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row

    If lastRow < 2 Then
        MsgBox _
            "データがありません。" & vbCrLf & _
            "原本シートの2行目以降に本文を入力してください。", _
            vbExclamation
        Exit Sub
    End If

    outPath = _
        "C:\Users\tagut\OneDrive\デスクトップ\【拡張機能】\" & _
        "テンプレート.json"

    jsonText = "[" & vbCrLf
    isFirst = True

    For rowNumber = 2 To lastRow

        subjectText = CStr(ws.Cells(rowNumber, "B").Value)
        bodyText = CStr(ws.Cells(rowNumber, "C").Value)
        tagsText = CStr(ws.Cells(rowNumber, "D").Value)

        '本文が空の行は出力しない
        If Len(Trim$(bodyText)) > 0 Then

            If Not isFirst Then jsonText = jsonText & "," & vbCrLf
            isFirst = False

            jsonText = jsonText & "  {" & vbCrLf _
                & "    ""id"": " _
                & TM_JsonString("row-" & Format$(rowNumber - 1, "0000")) _
                & "," & vbCrLf _
                & "    ""subject"": " & TM_JsonString(subjectText) _
                & "," & vbCrLf _
                & "    ""body"": " & TM_JsonString(bodyText) _
                & "," & vbCrLf _
                & "    ""tags"": " & TM_TagsToJsonArray(tagsText) _
                & vbCrLf _
                & "  }"

        End If

    Next rowNumber

    jsonText = jsonText & vbCrLf & "]" & vbCrLf

    TM_SaveUtf8NoBom outPath, jsonText

    MsgBox _
        "出力しました：" & vbCrLf & outPath, _
        vbInformation
    Exit Sub

ExportError:
    MsgBox _
        "エクスポートできませんでした。" & vbCrLf & vbCrLf & _
        "エラー内容：" & Err.Description, _
        vbExclamation

End Sub


'============================================================
' フォームの「原本を反映」ボタンから呼び出す転記処理
' このブックの原本シートを、選択したブックの原本シートへ上書きする
' 両方のシートがVeryHiddenでも表示状態を変更せず処理できる
'============================================================
Public Sub CopyOriginalSheetToSelectedWorkbook()

    Const msoFileDialogFilePicker As Long = 3
    Const msoAutomationSecurityForceDisable As Long = 3

    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim targetWorkbook As Workbook
    Dim openedWorkbook As Workbook
    Dim fileDialog As Object
    Dim selectedPath As String
    Dim answer As VbMsgBoxResult
    Dim previousScreenUpdating As Boolean
    Dim previousEnableEvents As Boolean
    Dim previousDisplayAlerts As Boolean
    Dim previousAutomationSecurity As Long
    Dim applicationStateSaved As Boolean
    Dim automationSecurityChanged As Boolean
    Dim targetOpenedByProcedure As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo CopyError

    Set sourceSheet = ThisWorkbook.Worksheets("原本")
    Set fileDialog = Application.FileDialog(msoFileDialogFilePicker)

    With fileDialog
        .Title = "原本を反映するExcelファイルを選択してください"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add _
            "Excelファイル", _
            "*.xlsx;*.xlsm;*.xlsb;*.xls"

        If Len(ThisWorkbook.Path) > 0 Then
            .InitialFileName = _
                ThisWorkbook.Path & Application.PathSeparator
        End If

        If .Show <> -1 Then Exit Sub
        selectedPath = CStr(.SelectedItems(1))
    End With

    If StrComp( _
        selectedPath, _
        ThisWorkbook.FullName, _
        vbTextCompare) = 0 Then

        MsgBox _
            "現在開いているファイル自身には反映できません。" & vbCrLf & _
            "別のExcelファイルを選択してください。", _
            vbExclamation, _
            "ファイル選択"
        Exit Sub
    End If

    Set openedWorkbook = TM_FindOpenWorkbook(selectedPath)

    If Not openedWorkbook Is Nothing Then
        MsgBox _
            "選択したファイルは既に開かれています。" & vbCrLf & _
            "ファイルを閉じてから、もう一度実行してください。", _
            vbExclamation, _
            "ファイル確認"
        Exit Sub
    End If

    answer = MsgBox( _
        "次のファイルの「原本」シートを全消去し、" & vbCrLf & _
        "このファイルの「原本」に置き換えます。" & vbCrLf & vbCrLf & _
        selectedPath & vbCrLf & vbCrLf & _
        "処理を続けますか？", _
        vbQuestion + vbYesNo + vbDefaultButton2, _
        "原本の反映確認")

    If answer <> vbYes Then Exit Sub

    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents
    previousDisplayAlerts = Application.DisplayAlerts
    previousAutomationSecurity = Application.AutomationSecurity
    applicationStateSaved = True

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.AutomationSecurity = _
        msoAutomationSecurityForceDisable
    automationSecurityChanged = True

    Set targetWorkbook = Workbooks.Open( _
        Filename:=selectedPath, _
        UpdateLinks:=0, _
        ReadOnly:=False, _
        IgnoreReadOnlyRecommended:=False)
    targetOpenedByProcedure = True

    Application.AutomationSecurity = previousAutomationSecurity
    automationSecurityChanged = False

    If targetWorkbook.ReadOnly Then
        Err.Raise _
            vbObjectError + 2101, _
            "CopyOriginalSheetToSelectedWorkbook", _
            "選択したファイルが読み取り専用で開かれました。" & _
            vbCrLf & _
            "編集用パスワードを入力して開ける状態で、" & _
            "もう一度実行してください。"
    End If

    On Error Resume Next
    Set targetSheet = targetWorkbook.Worksheets("原本")
    On Error GoTo CopyError

    If targetSheet Is Nothing Then
        Err.Raise _
            vbObjectError + 2102, _
            "CopyOriginalSheetToSelectedWorkbook", _
            "選択したファイルに「原本」シートがありません。"
    End If

    TM_ReplaceWorksheetContents sourceSheet, targetSheet

    targetWorkbook.Save
    targetWorkbook.Close SaveChanges:=False
    targetOpenedByProcedure = False

    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEnableEvents
    Application.DisplayAlerts = previousDisplayAlerts

    MsgBox _
        "選択したファイルへ原本を反映しました。" & vbCrLf & _
        selectedPath, _
        vbInformation, _
        "反映完了"
    Exit Sub

CopyError:
    errorNumber = Err.Number
    errorDescription = Err.Description

    On Error Resume Next

    If automationSecurityChanged Then
        Application.AutomationSecurity = previousAutomationSecurity
    End If

    Application.CutCopyMode = False

    If targetOpenedByProcedure Then
        Application.DisplayAlerts = False
        targetWorkbook.Close SaveChanges:=False
    End If

    If applicationStateSaved Then
        Application.ScreenUpdating = previousScreenUpdating
        Application.EnableEvents = previousEnableEvents
        Application.DisplayAlerts = previousDisplayAlerts
    End If

    On Error GoTo 0

    MsgBox _
        "原本を反映できませんでした。" & vbCrLf & vbCrLf & _
        "エラー番号：" & CStr(errorNumber) & vbCrLf & _
        "エラー内容：" & errorDescription, _
        vbExclamation, _
        "反映エラー"

End Sub


'============================================================
' 指定したファイルが既に開かれているか確認
'============================================================
Private Function TM_FindOpenWorkbook( _
    ByVal targetPath As String) As Workbook

    Dim workbookItem As Workbook

    For Each workbookItem In Application.Workbooks
        If StrComp( _
            workbookItem.FullName, _
            targetPath, _
            vbTextCompare) = 0 Then

            Set TM_FindOpenWorkbook = workbookItem
            Exit Function
        End If
    Next workbookItem

End Function


'============================================================
' 転記先を全消去し、転記元の使用範囲をそのまま複製
' セル内容・数式・書式・列幅・行高・非表示状態を引き継ぐ
'============================================================
Private Sub TM_ReplaceWorksheetContents( _
    ByVal sourceSheet As Worksheet, _
    ByVal targetSheet As Worksheet)

    Dim sourceRange As Range
    Dim targetRange As Range
    Dim rowNumber As Long
    Dim columnNumber As Long
    Dim lastRow As Long
    Dim lastColumn As Long

    Set sourceRange = sourceSheet.UsedRange
    Set targetRange = targetSheet.Range(sourceRange.Address)

    targetSheet.Cells.Clear

    sourceRange.Copy
    targetRange.PasteSpecial Paste:=xlPasteAll
    targetRange.PasteSpecial Paste:=xlPasteColumnWidths

    lastRow = sourceRange.Row + sourceRange.Rows.Count - 1
    lastColumn = sourceRange.Column + sourceRange.Columns.Count - 1

    For rowNumber = sourceRange.Row To lastRow
        targetSheet.Rows(rowNumber).RowHeight = _
            sourceSheet.Rows(rowNumber).RowHeight
        targetSheet.Rows(rowNumber).Hidden = _
            sourceSheet.Rows(rowNumber).Hidden
    Next rowNumber

    For columnNumber = sourceRange.Column To lastColumn
        targetSheet.Columns(columnNumber).Hidden = _
            sourceSheet.Columns(columnNumber).Hidden
    Next columnNumber

    Application.CutCopyMode = False

End Sub


'============================================================
' タグ文字列をJSON配列へ変換
'============================================================
Private Function TM_TagsToJsonArray(ByVal sourceText As String) As String

    Dim tagArray() As String
    Dim tagDictionary As Object
    Dim index As Long
    Dim tagText As String
    Dim key As Variant
    Dim jsonText As String
    Dim isFirst As Boolean

    sourceText = Trim$(sourceText)

    If Len(sourceText) = 0 Then
        TM_TagsToJsonArray = "[]"
        Exit Function
    End If

    sourceText = Replace(sourceText, ChrW(&H3000), " ")
    sourceText = WorksheetFunction.Trim(sourceText)

    tagArray = Split(sourceText, " ")

    Set tagDictionary = CreateObject("Scripting.Dictionary")
    tagDictionary.CompareMode = 1

    For index = LBound(tagArray) To UBound(tagArray)

        tagText = Trim$(tagArray(index))

        If Len(tagText) > 0 Then
            If Left$(tagText, 1) = "#" Then
                tagText = Mid$(tagText, 2)
            End If

            tagText = Trim$(tagText)

            If Len(tagText) > 0 Then
                If Not tagDictionary.Exists(tagText) Then
                    tagDictionary.Add tagText, True
                End If
            End If
        End If

    Next index

    jsonText = "["
    isFirst = True

    For Each key In tagDictionary.Keys
        If Not isFirst Then jsonText = jsonText & ","
        isFirst = False
        jsonText = jsonText & TM_JsonString(CStr(key))
    Next key

    jsonText = jsonText & "]"
    TM_TagsToJsonArray = jsonText

End Function


'============================================================
' JSON文字列用エスケープ
'============================================================
Private Function TM_JsonString(ByVal sourceText As String) As String

    Dim index As Long
    Dim characterText As String
    Dim characterCode As Long
    Dim escapedText As String

    ' UserFormの複数行TextBoxは改行をCRLFで保持することがある。
    ' 1文字ずつJSONエスケープする前にLFへ統一し、
    ' Enter 1回が「\\n\\n」と二重出力されるのを防ぐ。
    sourceText = TM_NormalizeLineBreaks(sourceText)

    escapedText = """"

    For index = 1 To Len(sourceText)

        characterText = Mid$(sourceText, index, 1)
        characterCode = AscW(characterText)

        Select Case characterText
            Case """"
                escapedText = escapedText & "\"""
            Case "\"
                escapedText = escapedText & "\\"
            Case vbTab
                escapedText = escapedText & "\t"
            Case vbCr
                escapedText = escapedText & "\n"
            Case vbLf
                escapedText = escapedText & "\n"
            Case Else
                If characterCode < 32 Then
                    escapedText = escapedText & "\u" _
                        & Right$("000" & Hex$(characterCode), 4)
                Else
                    escapedText = escapedText & characterText
                End If
        End Select

    Next index

    escapedText = escapedText & """"
    TM_JsonString = escapedText

End Function


'============================================================
' 改行コードをLFへ統一
' CRLFを先に置換することで、意図した改行数を維持する
'============================================================
Private Function TM_NormalizeLineBreaks(ByVal sourceText As String) As String

    sourceText = Replace(sourceText, vbCrLf, vbLf)
    sourceText = Replace(sourceText, vbCr, vbLf)
    TM_NormalizeLineBreaks = sourceText

End Function


'============================================================
' UTF-8（BOMなし）で保存
'============================================================
Private Sub TM_SaveUtf8NoBom( _
    ByVal filePath As String, _
    ByVal outputText As String)

    Dim textStream As Object
    Dim binaryStream As Object
    Dim byteData() As Byte
    Dim noBomData() As Byte
    Dim index As Long

    Set textStream = CreateObject("ADODB.Stream")

    With textStream
        .Type = 2
        .Charset = "utf-8"
        .Open
        .WriteText outputText
        .Position = 0
        .Type = 1
        byteData = .Read
        .Close
    End With

    If UBound(byteData) >= 2 Then
        If byteData(0) = &HEF _
           And byteData(1) = &HBB _
           And byteData(2) = &HBF Then

            ReDim noBomData(0 To UBound(byteData) - 3)

            For index = 3 To UBound(byteData)
                noBomData(index - 3) = byteData(index)
            Next index

            byteData = noBomData
        End If
    End If

    Set binaryStream = CreateObject("ADODB.Stream")

    With binaryStream
        .Type = 1
        .Open
        .Write byteData
        .SaveToFile filePath, 2
        .Close
    End With

End Sub


'============================================================
' ラベル追加
'============================================================
Private Sub AddLabel( _
    ByVal designer As Object, _
    ByVal controlName As String, _
    ByVal captionText As String, _
    ByVal leftPos As Single, _
    ByVal topPos As Single, _
    ByVal controlWidth As Single, _
    ByVal controlHeight As Single)

    Dim ctl As Object

    Set ctl = designer.Controls.Add("Forms.Label.1", controlName, True)

    With ctl
        .Caption = captionText
        .Left = leftPos * mFormScale
        .Top = topPos * mFormScale
        .Width = controlWidth * mFormScale
        .Height = controlHeight * mFormScale
        .BackStyle = 0
        .Font.Name = "Meiryo UI"
        .Font.Size = 9 * mFormScale
        .ForeColor = RGB(51, 65, 85)
    End With

End Sub


'============================================================
' テキストボックス追加
'============================================================
Private Sub AddTextBox( _
    ByVal designer As Object, _
    ByVal controlName As String, _
    ByVal leftPos As Single, _
    ByVal topPos As Single, _
    ByVal controlWidth As Single, _
    ByVal controlHeight As Single, _
    ByVal isMultiLine As Boolean)

    Dim ctl As Object

    Set ctl = designer.Controls.Add("Forms.TextBox.1", controlName, True)

    With ctl
        .Left = leftPos * mFormScale
        .Top = topPos * mFormScale
        .Width = controlWidth * mFormScale
        .Height = controlHeight * mFormScale
        .Font.Name = "Meiryo UI"
        .Font.Size = 10 * mFormScale
        .BackColor = RGB(255, 255, 255)
        .BorderStyle = 1
        .SpecialEffect = 0
        .MultiLine = isMultiLine
        .WordWrap = isMultiLine

        If isMultiLine Then
            .EnterKeyBehavior = True
            .ScrollBars = 2
        End If
    End With

End Sub


'============================================================
' リストボックス追加
'============================================================
Private Sub AddListBox( _
    ByVal designer As Object, _
    ByVal controlName As String, _
    ByVal leftPos As Single, _
    ByVal topPos As Single, _
    ByVal controlWidth As Single, _
    ByVal controlHeight As Single)

    Dim ctl As Object

    Set ctl = designer.Controls.Add("Forms.ListBox.1", controlName, True)

    With ctl
        .Left = leftPos * mFormScale
        .Top = topPos * mFormScale
        .Width = controlWidth * mFormScale
        .Height = controlHeight * mFormScale
        .Font.Name = "Meiryo UI"
        .Font.Size = 9 * mFormScale
        .BackColor = RGB(255, 255, 255)
        .BorderStyle = 1
        .SpecialEffect = 0
        .ColumnCount = 2
        .ColumnWidths = "0 pt;" & CStr(250 * mFormScale) & " pt"
        .IntegralHeight = False
    End With

End Sub


'============================================================
' ボタン追加
'============================================================
Private Sub AddButton( _
    ByVal designer As Object, _
    ByVal controlName As String, _
    ByVal captionText As String, _
    ByVal leftPos As Single, _
    ByVal topPos As Single, _
    ByVal controlWidth As Single, _
    ByVal controlHeight As Single, _
    ByVal buttonColor As Long)

    Dim ctl As Object

    Set ctl = designer.Controls.Add("Forms.CommandButton.1", controlName, True)

    With ctl
        .Caption = captionText
        .Left = leftPos * mFormScale
        .Top = topPos * mFormScale
        .Width = controlWidth * mFormScale
        .Height = controlHeight * mFormScale
        .Font.Name = "Meiryo UI"
        '小数のフォントサイズは表示倍率によって位置がずれることがあるため、
        '24インチ版は整数の12ptに固定する
        If mFormScale > STANDARD_SCALE Then
            .Font.Size = 12
        Else
            .Font.Size = 10
        End If
        .Font.Bold = True
        .ForeColor = RGB(255, 255, 255)
        .BackColor = buttonColor
        .TakeFocusOnClick = False
    End With

End Sub


'============================================================
' 自動生成するユーザーフォーム内のコード
'============================================================
Private Function GetFormCode() As String

    Dim code As String

    AddCodeLine code, "Option Explicit"
    AddCodeLine code, ""
    AddCodeLine code, "Private Const TARGET_SHEET As String = ""原本"""
    AddCodeLine code, "Private Const FIRST_DATA_ROW As Long = 2"
    AddCodeLine code, "Private Const SEARCH_PLACEHOLDER As String = ""キーワードを入力してください"""
    AddCodeLine code, ""
    AddCodeLine code, "Private SelectedID As Long"
    AddCodeLine code, "Private IsPlaceholder As Boolean"
    AddCodeLine code, ""
    AddCodeLine code, "Private Sub UserForm_Initialize()"
    AddCodeLine code, "    SelectedID = 0"
    AddCodeLine code, "    With Me.lstResults"
    AddCodeLine code, "        .Clear"
    AddCodeLine code, "        .ColumnCount = 2"
    AddCodeLine code, "        .ColumnWidths = ""0 pt;"" & CStr(.Width - 20) & "" pt"""
    AddCodeLine code, "    End With"
    AddCodeLine code, "    SetSearchPlaceholder"
    AddCodeLine code, "    ClearInputFields"
    AddCodeLine code, "    SetDefaultButtonCaptions"
    AddCodeLine code, "    LoadList """""
    AddCodeLine code, "End Sub"
    AddCodeLine code, ""
    AddCodeLine code, "Private Sub txtSearch_Enter()"
    AddCodeLine code, "    If IsPlaceholder Then"
    AddCodeLine code, "        Me.txtSearch.Value = """""
    AddCodeLine code, "        Me.txtSearch.ForeColor = RGB(0, 0, 0)"
    AddCodeLine code, "        IsPlaceholder = False"
    AddCodeLine code, "    End If"
    AddCodeLine code, "End Sub"
    AddCodeLine code, ""
    AddCodeLine code, "Private Sub txtSearch_Exit(ByVal Cancel As MSForms.ReturnBoolean)"
    AddCodeLine code, "    If Trim$(Me.txtSearch.Value) = """" Then SetSearchPlaceholder"
    AddCodeLine code, "End Sub"
    AddCodeLine code, ""
    AddCodeLine code, "Private Sub txtSearch_Change()"
    AddCodeLine code, "    If IsPlaceholder Then Exit Sub"
    AddCodeLine code, "    SelectedID = 0"
    AddCodeLine code, "    ClearInputFields"
    AddCodeLine code, "    SetDefaultButtonCaptions"
    AddCodeLine code, "    LoadList Trim$(Me.txtSearch.Value)"
    AddCodeLine code, "End Sub"
    AddCodeLine code, ""
    AddCodeLine code, "Private Sub lstResults_Click()"
    AddCodeLine code, "    Dim ws As Worksheet"
    AddCodeLine code, "    Dim targetRow As Long"
    AddCodeLine code, "    If Me.lstResults.ListIndex < 0 Then Exit Sub"
    AddCodeLine code, "    SelectedID = CLng(Me.lstResults.List(Me.lstResults.ListIndex, 0))"
    AddCodeLine code, "    Set ws = ThisWorkbook.Worksheets(TARGET_SHEET)"
    AddCodeLine code, "    targetRow = FindRowByID(ws, SelectedID)"
    AddCodeLine code, "    If targetRow = 0 Then"
    AddCodeLine code, "        MsgBox ""選択したデータが見つかりませんでした。"", vbExclamation"
    AddCodeLine code, "        ResetForm"
    AddCodeLine code, "        Exit Sub"
    AddCodeLine code, "    End If"
    AddCodeLine code, "    Me.txtColumnB.Value = ws.Cells(targetRow, ""B"").Value"
    AddCodeLine code, "    Me.txtColumnC.Value = ws.Cells(targetRow, ""C"").Value"
    AddCodeLine code, "    Me.txtColumnD.Value = ws.Cells(targetRow, ""D"").Value"
    AddCodeLine code, "    Me.btnRegister.Caption = ""上書き"""
    AddCodeLine code, "End Sub"
    AddCodeLine code, ""
    AddCodeLine code, "Private Sub btnRegister_Click()"
    AddCodeLine code, "    Dim ws As Worksheet"
    AddCodeLine code, "    Dim targetRow As Long"
    AddCodeLine code, "    Dim lastRow As Long"
    AddCodeLine code, "    Dim nextID As Long"
    AddCodeLine code, "    Dim result As VbMsgBoxResult"
    AddCodeLine code, "    Set ws = ThisWorkbook.Worksheets(TARGET_SHEET)"
    AddCodeLine code, "    If Trim$(Me.txtColumnC.Value) = """" Then"
    AddCodeLine code, "        MsgBox ""本文を入力してください。"", vbExclamation"
    AddCodeLine code, "        Me.txtColumnC.SetFocus"
    AddCodeLine code, "        Exit Sub"
    AddCodeLine code, "    End If"
    AddCodeLine code, "    If SelectedID = 0 Then"
    AddCodeLine code, "        lastRow = GetLastRow(ws)"
    AddCodeLine code, "        targetRow = lastRow + 1"
    AddCodeLine code, "        If lastRow < FIRST_DATA_ROW Then"
    AddCodeLine code, "            nextID = 1"
    AddCodeLine code, "        ElseIf IsNumeric(ws.Cells(lastRow, ""A"").Value) Then"
    AddCodeLine code, "            nextID = CLng(ws.Cells(lastRow, ""A"").Value) + 1"
    AddCodeLine code, "        Else"
    AddCodeLine code, "            nextID = lastRow - FIRST_DATA_ROW + 2"
    AddCodeLine code, "        End If"
    AddCodeLine code, "        ws.Cells(targetRow, ""A"").Value = nextID"
    AddCodeLine code, "        ws.Cells(targetRow, ""B"").Value = Me.txtColumnB.Value"
    AddCodeLine code, "        ws.Cells(targetRow, ""C"").Value = Me.txtColumnC.Value"
    AddCodeLine code, "        ws.Cells(targetRow, ""D"").Value = Me.txtColumnD.Value"
    AddCodeLine code, "        MsgBox ""登録しました。"", vbInformation"
    AddCodeLine code, "    Else"
    AddCodeLine code, "        targetRow = FindRowByID(ws, SelectedID)"
    AddCodeLine code, "        If targetRow = 0 Then"
    AddCodeLine code, "            MsgBox ""上書き対象のデータが見つかりませんでした。"", vbExclamation"
    AddCodeLine code, "            ResetForm"
    AddCodeLine code, "            Exit Sub"
    AddCodeLine code, "        End If"
    AddCodeLine code, "        result = MsgBox(""選択中のデータを上書きしますか？"", vbQuestion + vbYesNo, ""上書き確認"")"
    AddCodeLine code, "        If result <> vbYes Then Exit Sub"
    AddCodeLine code, "        ws.Cells(targetRow, ""B"").Value = Me.txtColumnB.Value"
    AddCodeLine code, "        ws.Cells(targetRow, ""C"").Value = Me.txtColumnC.Value"
    AddCodeLine code, "        ws.Cells(targetRow, ""D"").Value = Me.txtColumnD.Value"
    AddCodeLine code, "        MsgBox ""上書きしました。"", vbInformation"
    AddCodeLine code, "    End If"
    AddCodeLine code, "    ResetForm"
    AddCodeLine code, "End Sub"
    AddCodeLine code, ""
    AddCodeLine code, "Private Sub btnDelete_Click()"
    AddCodeLine code, "    Dim ws As Worksheet"
    AddCodeLine code, "    Dim targetRow As Long"
    AddCodeLine code, "    Dim result As VbMsgBoxResult"
    AddCodeLine code, "    If SelectedID = 0 Then"
    AddCodeLine code, "        MsgBox ""削除するデータを検索結果から選択してください。"", vbExclamation"
    AddCodeLine code, "        Exit Sub"
    AddCodeLine code, "    End If"
    AddCodeLine code, "    result = MsgBox(""選択中のデータを削除しますか？"" & vbCrLf & ""削除したデータは元に戻せません。"", vbExclamation + vbYesNo + vbDefaultButton2, ""削除確認"")"
    AddCodeLine code, "    If result <> vbYes Then Exit Sub"
    AddCodeLine code, "    Set ws = ThisWorkbook.Worksheets(TARGET_SHEET)"
    AddCodeLine code, "    targetRow = FindRowByID(ws, SelectedID)"
    AddCodeLine code, "    If targetRow = 0 Then"
    AddCodeLine code, "        MsgBox ""削除対象のデータが見つかりませんでした。"", vbExclamation"
    AddCodeLine code, "        ResetForm"
    AddCodeLine code, "        Exit Sub"
    AddCodeLine code, "    End If"
    AddCodeLine code, "    ws.Rows(targetRow).Delete Shift:=xlUp"
    AddCodeLine code, "    RenumberIDs ws"
    AddCodeLine code, "    MsgBox ""削除しました。"", vbInformation"
    AddCodeLine code, "    ResetForm"
    AddCodeLine code, "End Sub"
    AddCodeLine code, ""
    AddCodeLine code, "Private Sub btnClear_Click()"
    AddCodeLine code, "    ResetForm"
    AddCodeLine code, "End Sub"
    AddCodeLine code, ""
    AddCodeLine code, "Private Sub btnExport_Click()"
    AddCodeLine code, "    Application.Run ""'"" & ThisWorkbook.Name & ""'!ExportTemplatesJsonFromForm"""
    AddCodeLine code, "End Sub"
    AddCodeLine code, ""
    AddCodeLine code, "Private Sub btnCopyOriginal_Click()"
    AddCodeLine code, "    Application.Run ""'"" & ThisWorkbook.Name & ""'!CopyOriginalSheetToSelectedWorkbook"""
    AddCodeLine code, "End Sub"
    AddCodeLine code, ""
    AddCodeLine code, "Private Sub LoadList(ByVal keyword As String)"
    AddCodeLine code, "    Dim ws As Worksheet"
    AddCodeLine code, "    Dim lastRow As Long"
    AddCodeLine code, "    Dim rowNumber As Long"
    AddCodeLine code, "    Dim searchText As String"
    AddCodeLine code, "    Dim displayText As String"
    AddCodeLine code, "    Dim bodyPreview As String"
    AddCodeLine code, "    Set ws = ThisWorkbook.Worksheets(TARGET_SHEET)"
    AddCodeLine code, "    Me.lstResults.Clear"
    AddCodeLine code, "    lastRow = GetLastRow(ws)"
    AddCodeLine code, "    If lastRow < FIRST_DATA_ROW Then Exit Sub"
    AddCodeLine code, "    For rowNumber = FIRST_DATA_ROW To lastRow"
    AddCodeLine code, "        searchText = CStr(ws.Cells(rowNumber, ""C"").Value)"
    AddCodeLine code, "        If keyword = """" Or InStr(1, searchText, keyword, vbTextCompare) > 0 Then"
    AddCodeLine code, "            displayText = Trim$(CStr(ws.Cells(rowNumber, ""B"").Value))"
    AddCodeLine code, "            If Len(displayText) = 0 Then"
    AddCodeLine code, "                bodyPreview = CStr(ws.Cells(rowNumber, ""C"").Value)"
    AddCodeLine code, "                bodyPreview = Replace(bodyPreview, vbCrLf, "" "")"
    AddCodeLine code, "                bodyPreview = Replace(bodyPreview, vbCr, "" "")"
    AddCodeLine code, "                bodyPreview = Replace(bodyPreview, vbLf, "" "")"
    AddCodeLine code, "                bodyPreview = Replace(bodyPreview, vbTab, "" "")"
    AddCodeLine code, "                bodyPreview = Trim$(bodyPreview)"
    AddCodeLine code, "                If Len(bodyPreview) > 30 Then bodyPreview = Left$(bodyPreview, 30) & ""…"""
    AddCodeLine code, "                displayText = ""【件名なし｜ID:"" & CStr(ws.Cells(rowNumber, ""A"").Value) & ""】 "" & bodyPreview"
    AddCodeLine code, "            End If"
    AddCodeLine code, "            Me.lstResults.AddItem CStr(ws.Cells(rowNumber, ""A"").Value)"
    AddCodeLine code, "            Me.lstResults.List(Me.lstResults.ListCount - 1, 1) = displayText"
    AddCodeLine code, "        End If"
    AddCodeLine code, "    Next rowNumber"
    AddCodeLine code, "End Sub"
    AddCodeLine code, ""
    AddCodeLine code, "Private Function FindRowByID(ByVal ws As Worksheet, ByVal searchID As Long) As Long"
    AddCodeLine code, "    Dim foundCell As Range"
    AddCodeLine code, "    Set foundCell = ws.Columns(""A"").Find(What:=searchID, After:=ws.Cells(1, ""A""), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)"
    AddCodeLine code, "    If foundCell Is Nothing Then"
    AddCodeLine code, "        FindRowByID = 0"
    AddCodeLine code, "    ElseIf foundCell.Row < FIRST_DATA_ROW Then"
    AddCodeLine code, "        FindRowByID = 0"
    AddCodeLine code, "    Else"
    AddCodeLine code, "        FindRowByID = foundCell.Row"
    AddCodeLine code, "    End If"
    AddCodeLine code, "End Function"
    AddCodeLine code, ""
    AddCodeLine code, "Private Function GetLastRow(ByVal ws As Worksheet) As Long"
    AddCodeLine code, "    Dim lastA As Long"
    AddCodeLine code, "    Dim lastB As Long"
    AddCodeLine code, "    Dim lastC As Long"
    AddCodeLine code, "    Dim lastD As Long"
    AddCodeLine code, "    lastA = ws.Cells(ws.Rows.Count, ""A"").End(xlUp).Row"
    AddCodeLine code, "    lastB = ws.Cells(ws.Rows.Count, ""B"").End(xlUp).Row"
    AddCodeLine code, "    lastC = ws.Cells(ws.Rows.Count, ""C"").End(xlUp).Row"
    AddCodeLine code, "    lastD = ws.Cells(ws.Rows.Count, ""D"").End(xlUp).Row"
    AddCodeLine code, "    GetLastRow = Application.Max(lastA, lastB, lastC, lastD)"
    AddCodeLine code, "End Function"
    AddCodeLine code, ""
    AddCodeLine code, "Private Sub RenumberIDs(ByVal ws As Worksheet)"
    AddCodeLine code, "    Dim lastRow As Long"
    AddCodeLine code, "    Dim rowNumber As Long"
    AddCodeLine code, "    lastRow = GetLastRow(ws)"
    AddCodeLine code, "    If lastRow < FIRST_DATA_ROW Then Exit Sub"
    AddCodeLine code, "    For rowNumber = FIRST_DATA_ROW To lastRow"
    AddCodeLine code, "        ws.Cells(rowNumber, ""A"").Value = rowNumber - 1"
    AddCodeLine code, "    Next rowNumber"
    AddCodeLine code, "End Sub"
    AddCodeLine code, ""
    AddCodeLine code, "Private Sub ResetForm()"
    AddCodeLine code, "    SelectedID = 0"
    AddCodeLine code, "    ClearInputFields"
    AddCodeLine code, "    SetSearchPlaceholder"
    AddCodeLine code, "    SetDefaultButtonCaptions"
    AddCodeLine code, "    Me.lstResults.ListIndex = -1"
    AddCodeLine code, "    LoadList """""
    AddCodeLine code, "End Sub"
    AddCodeLine code, ""
    AddCodeLine code, "Private Sub ClearInputFields()"
    AddCodeLine code, "    Me.txtColumnB.Value = """""
    AddCodeLine code, "    Me.txtColumnC.Value = """""
    AddCodeLine code, "    Me.txtColumnD.Value = """""
    AddCodeLine code, "End Sub"
    AddCodeLine code, ""
    AddCodeLine code, "Private Sub SetDefaultButtonCaptions()"
    AddCodeLine code, "    Me.btnRegister.Caption = ""登録"""
    AddCodeLine code, "End Sub"
    AddCodeLine code, ""
    AddCodeLine code, "Private Sub SetSearchPlaceholder()"
    AddCodeLine code, "    IsPlaceholder = True"
    AddCodeLine code, "    Me.txtSearch.Value = SEARCH_PLACEHOLDER"
    AddCodeLine code, "    Me.txtSearch.ForeColor = RGB(128, 128, 128)"
    AddCodeLine code, "End Sub"

    GetFormCode = code

End Function


'============================================================
' フォーム用コードへ1行追加
'============================================================
Private Sub AddCodeLine(ByRef code As String, ByVal lineText As String)

    code = code & lineText & vbCrLf

End Sub
