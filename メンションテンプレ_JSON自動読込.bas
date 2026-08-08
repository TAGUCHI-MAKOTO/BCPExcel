Attribute VB_Name = "modTemplateFormBuilderJson"
Option Explicit

' ============================================================
' テンプレート選択フォーム 自動作成マクロ（JSON自動読込版）
'
' 【前提ファイル・シート】
'   テンプレート.json  JSON_FOLDER_PATHで指定したフォルダーを優先（空欄時は従来どおり自動検索）
'   休日マスタ         A列:休日の日付（見出しの有無は問いません）
'   リスト             A列:Windowsユーザー名 / B列:表示名
'
' 【動作】
'   フォームを開くたびにテンプレート.jsonを自動で読み込みます。
'   読み込めない場合は「JSONテンプレート読込状況を確認」を実行してください。
'   原本シートは使用しません。
'   JSONの読み込みに失敗した場合、前回正常に読み込んだデータを維持します。
'
' 【使い方】
'   1. このbasを標準モジュールとしてインポート
'   2. Excelの「VBAプロジェクト オブジェクト モデルへのアクセスを信頼する」をON
'   3. メンションテンプレ_JSON版を作成 を実行（初回のみ）
'   4. メンションテンプレ_JSON版を開く を実行してフォームを表示
' ============================================================

Private Const FORM_NAME As String = "frmTemplateSelectorJson"
Private Const USER_LIST_SHEET As String = "リスト"
Private Const USER_PLACEHOLDER As String = "$name$"
Private Const JSON_FILE_NAME As String = "テンプレート.json"
' JSON格納フォルダーを指定します。
' 例: "C:\共有\メンションテンプレ"
' %USERPROFILE% / %OneDrive% などの環境変数も使用できます。
' 空欄の場合は従来どおりブックの場所から自動検索します。
Private Const JSON_FOLDER_PATH As String = ""

Private mCurrentDisplayName As String
Private mUserNameInitialized As Boolean
Private mUserNameWarningShown As Boolean

Private mJsonTemplateIDs() As String
Private mJsonTemplateSubjects() As String
Private mJsonTemplateBodies() As String
Private mJsonTemplateTags() As String
Private mJsonTemplateCount As Long
Private mJsonTemplateLoaded As Boolean
Private mJsonTemplateLoadedPath As String
Private mJsonTemplateLastModified As Date
Private mJsonTemplateLastSize As Long
Private mJsonTemplateLastError As String
Private mJsonTemplateResolvedLocalBookFolder As String

Public Sub メンションテンプレ_JSON版を作成()
    Dim targetBook As Workbook
    Dim vbProj As Object, vbComp As Object, frm As Object
    Dim stage As String, errNo As Long, errText As String

    On Error GoTo CreateError

    stage = "作成先ブックの確認"
    Set targetBook = ThisWorkbook
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
        .Properties("BackColor").Value = RGB(246, 237, 241)
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

    AddButton frm, "btnClear", "クリア", 944, 545, 130, 36, RGB(219, 188, 202)

    stage = "フォーム処理コードの登録"
    vbComp.CodeModule.AddFromString BuildFormCode()

    MsgBox "テンプレート選択フォーム（JSON版）を作成しました。" & vbCrLf & _
           "次回から「メンションテンプレ_JSON版を開く」を実行してください。", vbInformation
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

Public Sub JsonTemplateInitializeCurrentUserName()
    Dim ws As Worksheet
    Dim loginUserName As String
    Dim foundCell As Range

    mUserNameInitialized = True
    mCurrentDisplayName = vbNullString

    loginUserName = Trim$(Environ$("USERNAME"))
    If Len(loginUserName) = 0 Then
        JsonTemplateShowUserNameWarning "Windowsのユーザー名を取得できませんでした。"
        Exit Sub
    End If

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(USER_LIST_SHEET)
    Err.Clear
    On Error GoTo UserNameError

    If ws Is Nothing Then
        JsonTemplateShowUserNameWarning "シート「リスト」が見つかりません。"
        Exit Sub
    End If

    Set foundCell = ws.Columns("A").Find( _
        What:=loginUserName, _
        After:=ws.Cells(ws.Rows.Count, "A"), _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False)

    If foundCell Is Nothing Then
        JsonTemplateShowUserNameWarning "リストシートのA列にWindowsユーザー名が登録されていません。" & _
                                        vbCrLf & "ユーザー名: " & loginUserName
        Exit Sub
    End If

    mCurrentDisplayName = Trim$(CStr(ws.Cells(foundCell.Row, "B").Value))

    If Len(mCurrentDisplayName) = 0 Then
        JsonTemplateShowUserNameWarning "リストシートのB列に表示名が登録されていません。" & _
                                        vbCrLf & "対象行: " & CStr(foundCell.Row)
    End If
    Exit Sub

UserNameError:
    mCurrentDisplayName = vbNullString
    JsonTemplateShowUserNameWarning "利用者名の取得中にエラーが発生しました。" & _
                                    vbCrLf & Err.Description
End Sub

Public Function JsonTemplateReplaceUserName(ByVal sourceText As String) As String
    If Not mUserNameInitialized Then JsonTemplateInitializeCurrentUserName

    If Len(mCurrentDisplayName) = 0 Then
        JsonTemplateReplaceUserName = sourceText
    Else
        JsonTemplateReplaceUserName = Replace(sourceText, USER_PLACEHOLDER, _
                                              mCurrentDisplayName, 1, -1, vbTextCompare)
    End If
End Function

Private Sub JsonTemplateShowUserNameWarning(ByVal messageText As String)
    If mUserNameWarningShown Then Exit Sub

    mUserNameWarningShown = True
    MsgBox messageText & vbCrLf & vbCrLf & _
           "$name$は置換せず、そのまま表示します。", vbExclamation
End Sub

Public Sub メンションテンプレ_JSON版を開く()
    On Error GoTo NotCreated
    VBA.UserForms.Add(FORM_NAME).Show
    Exit Sub
NotCreated:
    MsgBox "JSON版フォームがまだ作成されていません。" & vbCrLf & _
           "先に「メンションテンプレ_JSON版を作成」を実行してください。", vbExclamation
End Sub

' ============================================================
' JSONファイルの自動読み込み
' ============================================================
Public Function JsonTemplateLoadLatest( _
    Optional ByVal showError As Boolean = True, _
    Optional ByVal forceReload As Boolean = False) As Boolean

    Dim filePath As String
    Dim modifiedAt As Date
    Dim fileSize As Long
    Dim jsonText As String
    Dim ids As Collection
    Dim subjects As Collection
    Dim bodies As Collection
    Dim tags As Collection

    On Error GoTo LoadError

    filePath = JsonTemplateFilePath()

    If Len(Dir$(filePath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) = 0 Then
        Err.Raise vbObjectError + 2100, , _
                  "テンプレート.jsonが見つかりません。" & vbCrLf & filePath
    End If

    modifiedAt = FileDateTime(filePath)
    fileSize = FileLen(filePath)

    If mJsonTemplateLoaded And Not forceReload Then
        If StrComp(mJsonTemplateLoadedPath, filePath, vbTextCompare) = 0 _
           And mJsonTemplateLastModified = modifiedAt _
           And mJsonTemplateLastSize = fileSize Then
            JsonTemplateLoadLatest = True
            Exit Function
        End If
    End If

    jsonText = JsonTemplateReadUtf8Text(filePath)

    Set ids = New Collection
    Set subjects = New Collection
    Set bodies = New Collection
    Set tags = New Collection

    JsonTemplateParseRoot jsonText, ids, subjects, bodies, tags
    JsonTemplateCommit ids, subjects, bodies, tags

    mJsonTemplateLoaded = True
    mJsonTemplateLoadedPath = filePath
    mJsonTemplateLastModified = modifiedAt
    mJsonTemplateLastSize = fileSize
    mJsonTemplateLastError = vbNullString

    JsonTemplateLoadLatest = True
    Exit Function

LoadError:
    mJsonTemplateLastError = Err.Description

    If showError Then
        MsgBox "テンプレート.jsonを読み込めませんでした。" & vbCrLf & vbCrLf & _
               "参照先: " & IIf(Len(filePath) > 0, filePath, "取得できませんでした") & vbCrLf & _
               "エラー内容: " & mJsonTemplateLastError & vbCrLf & vbCrLf & _
               IIf(mJsonTemplateLoaded, _
                   "前回正常に読み込んだテンプレートを使用します。", _
                   "テンプレート一覧は空の状態で開きます。"), _
               vbExclamation
    End If

    JsonTemplateLoadLatest = mJsonTemplateLoaded
End Function

Public Function JsonTemplateFilePath() As String
    Dim thisBookPath As String
    Dim activeBookPath As String
    Dim candidatePath As String
    Dim mappedPath As String
    Dim desktopPath As String
    Dim localBookFolder As String
    Dim configuredFolder As String

    '指定フォルダーが設定されている場合は、そこを最優先で使用します。
    configuredFolder = JsonTemplateExpandEnvironmentVariables(Trim$(JSON_FOLDER_PATH))

    If Len(configuredFolder) > 0 Then
        If Not JsonTemplateFolderExists(configuredFolder) Then
            Err.Raise vbObjectError + 2103, , _
                      "JSON格納フォルダーが見つかりません。" & vbCrLf & configuredFolder
        End If

        JsonTemplateFilePath = JsonTemplateJoinPath(configuredFolder, JSON_FILE_NAME)
        Exit Function
    End If

    '一度正常に読み込めた参照先が残っていれば最優先で使用します。
    If mJsonTemplateLoaded Then
        If JsonTemplateFileExists(mJsonTemplateLoadedPath) Then
            JsonTemplateFilePath = mJsonTemplateLoadedPath
            Exit Function
        End If
    End If

    thisBookPath = Trim$(ThisWorkbook.Path)

    '通常のローカル／共有フォルダー
    If Len(thisBookPath) > 0 And Not JsonTemplateIsWebPath(thisBookPath) Then
        candidatePath = JsonTemplateJoinPath(thisBookPath, JSON_FILE_NAME)
        JsonTemplateFilePath = candidatePath
        Exit Function
    End If

    On Error Resume Next
    If Not Application.ActiveWorkbook Is Nothing Then
        activeBookPath = Trim$(Application.ActiveWorkbook.Path)
    End If
    On Error GoTo 0

    If Len(activeBookPath) > 0 And Not JsonTemplateIsWebPath(activeBookPath) Then
        candidatePath = JsonTemplateJoinPath(activeBookPath, JSON_FILE_NAME)
        JsonTemplateFilePath = candidatePath
        Exit Function
    End If

    'OneDrive上のファイルは、デスクトップから開いても
    'Excel側ではhttps://～のURLとして開かれる場合があります。
    '最初にURLをOneDriveのローカル同期先へ読み替えます。
    If JsonTemplateIsWebPath(thisBookPath) Then
        mappedPath = JsonTemplateMappedJsonPath(thisBookPath)
        If Len(mappedPath) > 0 Then
            JsonTemplateFilePath = mappedPath
            Exit Function
        End If
    End If

    If JsonTemplateIsWebPath(activeBookPath) Then
        mappedPath = JsonTemplateMappedJsonPath(activeBookPath)
        If Len(mappedPath) > 0 Then
            JsonTemplateFilePath = mappedPath
            Exit Function
        End If
    End If

    'URLから正確なローカル階層へ変換できない場合は、
    'デスクトップ／OneDrive内で、このExcelブック自体を探します。
    'ブックが見つかったフォルダーを「同じ階層」としてJSONを参照します。
    localBookFolder = JsonTemplateFindLocalWorkbookFolder()
    If Len(localBookFolder) > 0 Then
        mJsonTemplateResolvedLocalBookFolder = localBookFolder
        JsonTemplateFilePath = _
            JsonTemplateJoinPath(localBookFolder, JSON_FILE_NAME)
        Exit Function
    End If

    '最後に、Windowsが認識しているデスクトップ直下を確認します。
    desktopPath = JsonTemplateDesktopFolder()
    If Len(desktopPath) > 0 Then
        JsonTemplateFilePath = _
            JsonTemplateJoinPath(desktopPath, JSON_FILE_NAME)
        Exit Function
    End If

    Err.Raise vbObjectError + 2102, , _
              "ExcelがOneDrive／SharePointのURL形式で開かれています。" & vbCrLf & _
              "URLからローカルの保存フォルダーを特定できませんでした。" & vbCrLf & _
              "このExcelとテンプレート.jsonを同じフォルダーに置いた状態で、" & _
              "OneDriveの同期完了後に再度開いてください。"
End Function

Private Function JsonTemplateMappedJsonPath( _
    ByVal webFolderPath As String) As String

    Dim decodedPath As String
    Dim relativePath As String
    Dim markerPosition As Long
    Dim rootValues As Variant
    Dim rootValue As Variant
    Dim candidateFolder As String
    Dim candidatePath As String

    decodedPath = JsonTemplateUrlDecode(webFolderPath)

    markerPosition = InStr(1, decodedPath, "/Documents/", vbTextCompare)
    If markerPosition > 0 Then
        relativePath = Mid$(decodedPath, markerPosition + Len("/Documents/"))
    Else
        markerPosition = InStr(1, decodedPath, "/Shared Documents/", vbTextCompare)
        If markerPosition > 0 Then
            relativePath = Mid$(decodedPath, markerPosition + Len("/Shared Documents/"))
        Else
            Exit Function
        End If
    End If

    relativePath = Replace(relativePath, "/", "\")
    Do While Left$(relativePath, 1) = "\"
        relativePath = Mid$(relativePath, 2)
    Loop

    rootValues = Array( _
        Trim$(Environ$("OneDriveCommercial")), _
        Trim$(Environ$("OneDrive")), _
        Trim$(Environ$("OneDriveConsumer")))

    For Each rootValue In rootValues
        If Len(CStr(rootValue)) > 0 Then
            candidatePath = JsonTemplateMappedCandidate( _
                CStr(rootValue), relativePath)

            If Len(candidatePath) > 0 Then
                JsonTemplateMappedJsonPath = candidatePath
                Exit Function
            End If
        End If
    Next rootValue
End Function

Private Function JsonTemplateDesktopFolder() As String
    Dim shellObject As Object

    On Error Resume Next
    Set shellObject = CreateObject("WScript.Shell")
    If Not shellObject Is Nothing Then
        JsonTemplateDesktopFolder = _
            Trim$(CStr(shellObject.SpecialFolders("Desktop")))
    End If
    On Error GoTo 0
End Function

Private Function JsonTemplateFindJsonInKnownDesktopFolders() As String
    Dim rootValues As Variant
    Dim rootValue As Variant
    Dim candidateFolder As Variant
    Dim candidatePath As String
    Dim desktopPath As String

    desktopPath = JsonTemplateDesktopFolder()

    rootValues = Array( _
        desktopPath, _
        JsonTemplateJoinPath(Trim$(Environ$("USERPROFILE")), "Desktop"), _
        JsonTemplateJoinPath(Trim$(Environ$("OneDriveCommercial")), "Desktop"), _
        JsonTemplateJoinPath(Trim$(Environ$("OneDriveCommercial")), "デスクトップ"), _
        JsonTemplateJoinPath(Trim$(Environ$("OneDrive")), "Desktop"), _
        JsonTemplateJoinPath(Trim$(Environ$("OneDrive")), "デスクトップ"), _
        JsonTemplateJoinPath(Trim$(Environ$("OneDriveConsumer")), "Desktop"), _
        JsonTemplateJoinPath(Trim$(Environ$("OneDriveConsumer")), "デスクトップ"))

    For Each rootValue In rootValues
        candidateFolder = Trim$(CStr(rootValue))

        If Len(CStr(candidateFolder)) > 0 Then
            candidatePath = JsonTemplateJoinPath( _
                CStr(candidateFolder), JSON_FILE_NAME)

            If JsonTemplateFileExists(candidatePath) Then
                JsonTemplateFindJsonInKnownDesktopFolders = candidatePath
                Exit Function
            End If
        End If
    Next rootValue
End Function


Private Function JsonTemplateMappedCandidate( _
    ByVal oneDriveRoot As String, _
    ByVal relativePath As String) As String

    Dim pathVariants As Variant
    Dim pathVariant As Variant
    Dim candidateFolder As String
    Dim candidatePath As String

    pathVariants = Array( _
        relativePath, _
        JsonTemplateReplaceFirstFolder(relativePath, "Desktop", "デスクトップ"), _
        JsonTemplateReplaceFirstFolder(relativePath, "デスクトップ", "Desktop"))

    For Each pathVariant In pathVariants
        If Len(Trim$(CStr(pathVariant))) > 0 Then
            candidateFolder = JsonTemplateJoinPath( _
                oneDriveRoot, CStr(pathVariant))
            candidatePath = JsonTemplateJoinPath( _
                candidateFolder, JSON_FILE_NAME)

            If JsonTemplateFileExists(candidatePath) Then
                JsonTemplateMappedCandidate = candidatePath
                Exit Function
            End If
        End If
    Next pathVariant
End Function

Private Function JsonTemplateReplaceFirstFolder( _
    ByVal relativePath As String, _
    ByVal sourceFolder As String, _
    ByVal targetFolder As String) As String

    Dim separatorPosition As Long
    Dim firstFolder As String
    Dim remainingPath As String

    separatorPosition = InStr(1, relativePath, "\", vbBinaryCompare)

    If separatorPosition > 0 Then
        firstFolder = Left$(relativePath, separatorPosition - 1)
        remainingPath = Mid$(relativePath, separatorPosition)
    Else
        firstFolder = relativePath
        remainingPath = vbNullString
    End If

    If StrComp(firstFolder, sourceFolder, vbTextCompare) = 0 Then
        JsonTemplateReplaceFirstFolder = targetFolder & remainingPath
    Else
        JsonTemplateReplaceFirstFolder = relativePath
    End If
End Function

Private Function JsonTemplateFindLocalWorkbookFolder() As String
    Dim desktopRoots As Variant
    Dim oneDriveRoots As Variant
    Dim workbookNames As Variant
    Dim rootValue As Variant
    Dim workbookName As Variant
    Dim foundFolder As String
    Dim desktopPath As String
    Dim activeBookName As String

    If Len(mJsonTemplateResolvedLocalBookFolder) > 0 Then
        If JsonTemplateFolderExists(mJsonTemplateResolvedLocalBookFolder) Then
            JsonTemplateFindLocalWorkbookFolder = _
                mJsonTemplateResolvedLocalBookFolder
            Exit Function
        End If
    End If

    On Error Resume Next
    If Not Application.ActiveWorkbook Is Nothing Then
        activeBookName = Application.ActiveWorkbook.Name
    End If
    On Error GoTo 0

    workbookNames = Array(ThisWorkbook.Name, activeBookName)
    desktopPath = JsonTemplateDesktopFolder()

    desktopRoots = Array( _
        desktopPath, _
        JsonTemplateJoinPath(Trim$(Environ$("USERPROFILE")), "Desktop"), _
        JsonTemplateJoinPath(Trim$(Environ$("USERPROFILE")), "デスクトップ"), _
        JsonTemplateJoinPath(Trim$(Environ$("OneDriveCommercial")), "Desktop"), _
        JsonTemplateJoinPath(Trim$(Environ$("OneDriveCommercial")), "デスクトップ"), _
        JsonTemplateJoinPath(Trim$(Environ$("OneDrive")), "Desktop"), _
        JsonTemplateJoinPath(Trim$(Environ$("OneDrive")), "デスクトップ"), _
        JsonTemplateJoinPath(Trim$(Environ$("OneDriveConsumer")), "Desktop"), _
        JsonTemplateJoinPath(Trim$(Environ$("OneDriveConsumer")), "デスクトップ"))

    'まずデスクトップ配下を深めに検索します。
    For Each rootValue In desktopRoots
        If JsonTemplateFolderExists(CStr(rootValue)) Then
            For Each workbookName In workbookNames
                If Len(Trim$(CStr(workbookName))) > 0 Then
                    foundFolder = JsonTemplateSearchFolderForFile( _
                        CStr(rootValue), CStr(workbookName), 0, 10)

                    If Len(foundFolder) > 0 Then
                        JsonTemplateFindLocalWorkbookFolder = foundFolder
                        Exit Function
                    End If
                End If
            Next workbookName
        End If
    Next rootValue

    oneDriveRoots = Array( _
        Trim$(Environ$("OneDriveCommercial")), _
        Trim$(Environ$("OneDrive")), _
        Trim$(Environ$("OneDriveConsumer")))

    'デスクトップで見つからない場合のみOneDrive配下を検索します。
    '検索が重くなりすぎないよう、階層は6段までに制限します。
    For Each rootValue In oneDriveRoots
        If JsonTemplateFolderExists(CStr(rootValue)) Then
            For Each workbookName In workbookNames
                If Len(Trim$(CStr(workbookName))) > 0 Then
                    foundFolder = JsonTemplateSearchFolderForFile( _
                        CStr(rootValue), CStr(workbookName), 0, 6)

                    If Len(foundFolder) > 0 Then
                        JsonTemplateFindLocalWorkbookFolder = foundFolder
                        Exit Function
                    End If
                End If
            Next workbookName
        End If
    Next rootValue
End Function

Private Function JsonTemplateSearchFolderForFile( _
    ByVal folderPath As String, _
    ByVal targetFileName As String, _
    ByVal currentDepth As Long, _
    ByVal maxDepth As Long) As String

    Dim fso As Object
    Dim folderObject As Object
    Dim subFolder As Object
    Dim candidateFile As String
    Dim foundFolder As String

    On Error GoTo SearchEnd

    If currentDepth > maxDepth Then Exit Function
    If Not JsonTemplateFolderExists(folderPath) Then Exit Function

    candidateFile = JsonTemplateJoinPath(folderPath, targetFileName)
    If JsonTemplateFileExists(candidateFile) Then
        JsonTemplateSearchFolderForFile = folderPath
        Exit Function
    End If

    If currentDepth = maxDepth Then Exit Function

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folderObject = fso.GetFolder(folderPath)

    For Each subFolder In folderObject.SubFolders
        foundFolder = JsonTemplateSearchFolderForFile( _
            CStr(subFolder.Path), targetFileName, _
            currentDepth + 1, maxDepth)

        If Len(foundFolder) > 0 Then
            JsonTemplateSearchFolderForFile = foundFolder
            Exit Function
        End If
    Next subFolder

SearchEnd:
End Function

Private Function JsonTemplateFolderExists( _
    ByVal folderPath As String) As Boolean

    Dim fso As Object

    If Len(Trim$(folderPath)) = 0 Then Exit Function

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then
        JsonTemplateFolderExists = fso.FolderExists(folderPath)
    End If
    On Error GoTo 0
End Function

Private Function JsonTemplateUrlDecode(ByVal sourceText As String) As String
    Dim htmlDocument As Object
    Dim decodedText As String

    decodedText = sourceText

    On Error Resume Next
    Set htmlDocument = CreateObject("htmlfile")
    If Not htmlDocument Is Nothing Then
        decodedText = _
            CStr(htmlDocument.parentWindow.decodeURIComponent(sourceText))
    End If
    On Error GoTo 0

    'decodeURIComponentが利用できない環境向けの最低限の補完
    decodedText = Replace(decodedText, "%20", " ", 1, -1, vbTextCompare)
    decodedText = Replace(decodedText, "%23", "#", 1, -1, vbTextCompare)
    decodedText = Replace(decodedText, "%25", "%", 1, -1, vbTextCompare)

    JsonTemplateUrlDecode = decodedText
End Function

Private Function JsonTemplateExpandEnvironmentVariables( _
    ByVal sourcePath As String) As String

    Dim shellObject As Object

    If Len(Trim$(sourcePath)) = 0 Then Exit Function

    On Error GoTo ExpandError
    Set shellObject = CreateObject("WScript.Shell")
    JsonTemplateExpandEnvironmentVariables = _
        Trim$(CStr(shellObject.ExpandEnvironmentStrings(sourcePath)))
    Exit Function

ExpandError:
    JsonTemplateExpandEnvironmentVariables = sourcePath
End Function
Private Function JsonTemplateJoinPath( _
    ByVal folderPath As String, _
    ByVal fileName As String) As String

    Dim lastCharacter As String

    If Len(folderPath) = 0 Then
        JsonTemplateJoinPath = fileName
        Exit Function
    End If

    lastCharacter = Right$(folderPath, 1)

    If lastCharacter = "\" Or lastCharacter = "/" Then
        JsonTemplateJoinPath = folderPath & fileName
    Else
        JsonTemplateJoinPath = folderPath & Application.PathSeparator & fileName
    End If
End Function

Private Function JsonTemplateFileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    JsonTemplateFileExists = _
        (Len(Dir$(filePath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) > 0)
    On Error GoTo 0
End Function

Private Function JsonTemplateIsWebPath(ByVal folderPath As String) As Boolean
    Dim lowerPath As String

    lowerPath = LCase$(Trim$(folderPath))
    JsonTemplateIsWebPath = _
        (Left$(lowerPath, 7) = "http://" Or Left$(lowerPath, 8) = "https://")
End Function

Public Sub JSONテンプレート読込状況を確認()
    Dim resolvedPath As String
    Dim existsText As String
    Dim loadText As String
    Dim details As String
    Dim activeBookName As String
    Dim activeBookPath As String

    On Error Resume Next
    If Not Application.ActiveWorkbook Is Nothing Then
        activeBookName = Application.ActiveWorkbook.Name
        activeBookPath = Application.ActiveWorkbook.Path
    End If
    On Error GoTo DiagnosticError

    resolvedPath = JsonTemplateFilePath()

    If JsonTemplateFileExists(resolvedPath) Then
        existsText = "見つかりました"
    Else
        existsText = "見つかりません"
    End If

    If JsonTemplateLoadLatest(False, True) Then
        loadText = "読込成功（" & CStr(JsonTemplateCount()) & "件）"
    Else
        loadText = "読込失敗"
        If Len(mJsonTemplateLastError) > 0 Then
            loadText = loadText & vbCrLf & mJsonTemplateLastError
        End If
    End If

    details = _
        "【JSON読込状況】" & vbCrLf & vbCrLf & _
        "マクロ格納ブック: " & ThisWorkbook.Name & vbCrLf & _
        "マクロ格納先: " & ThisWorkbook.Path & vbCrLf & vbCrLf & _
        "現在のブック: " & activeBookName & vbCrLf & _
        "現在の保存先: " & activeBookPath & vbCrLf & vbCrLf & _
        "指定JSONフォルダー: " & IIf(Len(Trim$(JSON_FOLDER_PATH)) > 0, JSON_FOLDER_PATH, "未指定（自動検索）") & vbCrLf & _
        "JSON参照先: " & resolvedPath & vbCrLf & _
        "ファイル確認: " & existsText & vbCrLf & _
        "読込結果: " & loadText

    MsgBox details, vbInformation, "JSON読込状況"
    Exit Sub

DiagnosticError:
    MsgBox _
        "JSONの参照先を確認できませんでした。" & vbCrLf & vbCrLf & _
        "マクロ格納ブック: " & ThisWorkbook.Name & vbCrLf & _
        "マクロ格納先: " & ThisWorkbook.Path & vbCrLf & _
        "現在のブック: " & activeBookName & vbCrLf & _
        "現在の保存先: " & activeBookPath & vbCrLf & vbCrLf & _
        "エラー内容: " & Err.Description, _
        vbExclamation, _
        "JSON読込状況"
End Sub

Public Function JsonTemplateCount() As Long
    JsonTemplateCount = mJsonTemplateCount
End Function

Public Function JsonTemplateDisplayID(ByVal templateIndex As Long) As Variant
    Dim idText As String
    Dim suffixText As String

    JsonTemplateValidateIndex templateIndex
    idText = mJsonTemplateIDs(templateIndex)

    If Len(idText) > 4 And LCase$(Left$(idText, 4)) = "row-" Then
        suffixText = Mid$(idText, 5)
        If IsNumeric(suffixText) Then
            JsonTemplateDisplayID = CLng(suffixText)
            Exit Function
        End If
    End If

    JsonTemplateDisplayID = idText
End Function

Public Function JsonTemplateSubject(ByVal templateIndex As Long) As String
    JsonTemplateValidateIndex templateIndex
    JsonTemplateSubject = mJsonTemplateSubjects(templateIndex)
End Function

Public Function JsonTemplateBody(ByVal templateIndex As Long) As String
    JsonTemplateValidateIndex templateIndex
    JsonTemplateBody = mJsonTemplateBodies(templateIndex)
End Function

Public Function JsonTemplateTagText(ByVal templateIndex As Long) As String
    JsonTemplateValidateIndex templateIndex
    JsonTemplateTagText = mJsonTemplateTags(templateIndex)
End Function

Private Sub JsonTemplateValidateIndex(ByVal templateIndex As Long)
    If templateIndex < 1 Or templateIndex > mJsonTemplateCount Then
        Err.Raise vbObjectError + 2102, , _
                  "テンプレートの参照位置が範囲外です。"
    End If
End Sub

Private Sub JsonTemplateCommit( _
    ByVal ids As Collection, _
    ByVal subjects As Collection, _
    ByVal bodies As Collection, _
    ByVal tags As Collection)

    Dim itemIndex As Long
    Dim itemCount As Long

    itemCount = ids.Count
    mJsonTemplateCount = itemCount

    Erase mJsonTemplateIDs
    Erase mJsonTemplateSubjects
    Erase mJsonTemplateBodies
    Erase mJsonTemplateTags

    If itemCount = 0 Then Exit Sub

    ReDim mJsonTemplateIDs(1 To itemCount)
    ReDim mJsonTemplateSubjects(1 To itemCount)
    ReDim mJsonTemplateBodies(1 To itemCount)
    ReDim mJsonTemplateTags(1 To itemCount)

    For itemIndex = 1 To itemCount
        mJsonTemplateIDs(itemIndex) = CStr(ids(itemIndex))
        mJsonTemplateSubjects(itemIndex) = CStr(subjects(itemIndex))
        mJsonTemplateBodies(itemIndex) = CStr(bodies(itemIndex))
        mJsonTemplateTags(itemIndex) = CStr(tags(itemIndex))
    Next itemIndex
End Sub

Private Function JsonTemplateReadUtf8Text(ByVal filePath As String) As String
    Dim streamObject As Object
    Dim loadedText As String

    Set streamObject = CreateObject("ADODB.Stream")

    With streamObject
        .Type = 2
        .Charset = "utf-8"
        .Open
        .LoadFromFile filePath
        .Position = 0
        loadedText = .ReadText(-1)
        .Close
    End With

    'BOM付きUTF-8にも対応
    If Len(loadedText) > 0 Then
        If AscW(Left$(loadedText, 1)) = &HFEFF Then
            loadedText = Mid$(loadedText, 2)
        End If
    End If

    JsonTemplateReadUtf8Text = loadedText
End Function

Private Sub JsonTemplateParseRoot( _
    ByVal sourceText As String, _
    ByVal ids As Collection, _
    ByVal subjects As Collection, _
    ByVal bodies As Collection, _
    ByVal tags As Collection)

    Dim currentPosition As Long

    currentPosition = 1
    JsonTemplateSkipWhitespace sourceText, currentPosition
    JsonTemplateExpectCharacter sourceText, currentPosition, "["
    JsonTemplateSkipWhitespace sourceText, currentPosition

    If JsonTemplatePeekCharacter(sourceText, currentPosition) = "]" Then
        currentPosition = currentPosition + 1
        JsonTemplateEnsureEnd sourceText, currentPosition
        Exit Sub
    End If

    Do
        JsonTemplateParseObject sourceText, currentPosition, ids, subjects, bodies, tags
        JsonTemplateSkipWhitespace sourceText, currentPosition

        Select Case JsonTemplatePeekCharacter(sourceText, currentPosition)
            Case ","
                currentPosition = currentPosition + 1
            Case "]"
                currentPosition = currentPosition + 1
                Exit Do
            Case Else
                JsonTemplateRaiseParseError currentPosition, _
                    "配列の区切り「,」または終了「]」が必要です。"
        End Select
    Loop

    JsonTemplateEnsureEnd sourceText, currentPosition
End Sub

Private Sub JsonTemplateParseObject( _
    ByVal sourceText As String, _
    ByRef currentPosition As Long, _
    ByVal ids As Collection, _
    ByVal subjects As Collection, _
    ByVal bodies As Collection, _
    ByVal tags As Collection)

    Dim keyText As String
    Dim idText As String
    Dim subjectText As String
    Dim bodyText As String
    Dim tagText As String

    JsonTemplateSkipWhitespace sourceText, currentPosition
    JsonTemplateExpectCharacter sourceText, currentPosition, "{"
    JsonTemplateSkipWhitespace sourceText, currentPosition

    If JsonTemplatePeekCharacter(sourceText, currentPosition) <> "}" Then
        Do
            keyText = JsonTemplateParseString(sourceText, currentPosition)
            JsonTemplateSkipWhitespace sourceText, currentPosition
            JsonTemplateExpectCharacter sourceText, currentPosition, ":"
            JsonTemplateSkipWhitespace sourceText, currentPosition

            Select Case LCase$(keyText)
                Case "id"
                    idText = JsonTemplateParseString(sourceText, currentPosition)
                Case "subject"
                    subjectText = JsonTemplateParseString(sourceText, currentPosition)
                Case "body"
                    bodyText = JsonTemplateParseString(sourceText, currentPosition)
                Case "tags"
                    tagText = JsonTemplateParseStringArray(sourceText, currentPosition)
                Case Else
                    JsonTemplateSkipValue sourceText, currentPosition
            End Select

            JsonTemplateSkipWhitespace sourceText, currentPosition

            Select Case JsonTemplatePeekCharacter(sourceText, currentPosition)
                Case ","
                    currentPosition = currentPosition + 1
                    JsonTemplateSkipWhitespace sourceText, currentPosition
                Case "}"
                    Exit Do
                Case Else
                    JsonTemplateRaiseParseError currentPosition, _
                        "項目の区切り「,」または終了「}」が必要です。"
            End Select
        Loop
    End If

    JsonTemplateExpectCharacter sourceText, currentPosition, "}"

    If Len(idText) = 0 Then
        idText = "row-" & Format$(ids.Count + 1, "0000")
    End If

    ids.Add idText
    subjects.Add subjectText
    bodies.Add bodyText
    tags.Add tagText
End Sub

Private Function JsonTemplateParseStringArray( _
    ByVal sourceText As String, _
    ByRef currentPosition As Long) As String

    Dim resultText As String
    Dim itemText As String

    JsonTemplateExpectCharacter sourceText, currentPosition, "["
    JsonTemplateSkipWhitespace sourceText, currentPosition

    If JsonTemplatePeekCharacter(sourceText, currentPosition) = "]" Then
        currentPosition = currentPosition + 1
        JsonTemplateParseStringArray = vbNullString
        Exit Function
    End If

    Do
        itemText = JsonTemplateParseString(sourceText, currentPosition)

        If Len(itemText) > 0 Then
            If Len(resultText) > 0 Then resultText = resultText & " "
            resultText = resultText & itemText
        End If

        JsonTemplateSkipWhitespace sourceText, currentPosition

        Select Case JsonTemplatePeekCharacter(sourceText, currentPosition)
            Case ","
                currentPosition = currentPosition + 1
                JsonTemplateSkipWhitespace sourceText, currentPosition
            Case "]"
                currentPosition = currentPosition + 1
                Exit Do
            Case Else
                JsonTemplateRaiseParseError currentPosition, _
                    "タグ配列の区切り「,」または終了「]」が必要です。"
        End Select
    Loop

    JsonTemplateParseStringArray = resultText
End Function

Private Function JsonTemplateParseString( _
    ByVal sourceText As String, _
    ByRef currentPosition As Long) As String

    Dim resultText As String
    Dim characterText As String
    Dim escapeText As String
    Dim unicodeText As String
    Dim unicodeValue As Long

    JsonTemplateSkipWhitespace sourceText, currentPosition
    JsonTemplateExpectCharacter sourceText, currentPosition, """"

    Do While currentPosition <= Len(sourceText)
        characterText = Mid$(sourceText, currentPosition, 1)
        currentPosition = currentPosition + 1

        If characterText = """" Then
            JsonTemplateParseString = resultText
            Exit Function
        End If

        If characterText = "\" Then
            If currentPosition > Len(sourceText) Then
                JsonTemplateRaiseParseError currentPosition, _
                    "エスケープ文字の途中でJSONが終了しました。"
            End If

            escapeText = Mid$(sourceText, currentPosition, 1)
            currentPosition = currentPosition + 1

            Select Case escapeText
                Case """", "\", "/"
                    resultText = resultText & escapeText
                Case "b"
                    resultText = resultText & Chr$(8)
                Case "f"
                    resultText = resultText & Chr$(12)
                Case "n"
                    resultText = resultText & vbLf
                Case "r"
                    resultText = resultText & vbCr
                Case "t"
                    resultText = resultText & vbTab
                Case "u"
                    If currentPosition + 3 > Len(sourceText) Then
                        JsonTemplateRaiseParseError currentPosition, _
                            "Unicodeエスケープが途中で終了しました。"
                    End If

                    unicodeText = Mid$(sourceText, currentPosition, 4)
                    If Not JsonTemplateIsHex4(unicodeText) Then
                        JsonTemplateRaiseParseError currentPosition, _
                            "Unicodeエスケープの形式が正しくありません。"
                    End If

                    unicodeValue = CLng("&H" & unicodeText)
                    resultText = resultText & ChrW$(unicodeValue)
                    currentPosition = currentPosition + 4
                Case Else
                    JsonTemplateRaiseParseError currentPosition - 1, _
                        "未対応のエスケープ文字です: \" & escapeText
            End Select
        Else
            resultText = resultText & characterText
        End If
    Loop

    JsonTemplateRaiseParseError currentPosition, _
        "文字列の終了記号が見つかりません。"
End Function

Private Function JsonTemplateIsHex4(ByVal sourceText As String) As Boolean
    Dim characterIndex As Long
    Dim characterText As String

    If Len(sourceText) <> 4 Then Exit Function

    For characterIndex = 1 To 4
        characterText = Mid$(sourceText, characterIndex, 1)
        If InStr(1, "0123456789ABCDEFabcdef", characterText, vbBinaryCompare) = 0 Then
            Exit Function
        End If
    Next characterIndex

    JsonTemplateIsHex4 = True
End Function

Private Sub JsonTemplateSkipValue( _
    ByVal sourceText As String, _
    ByRef currentPosition As Long)

    Dim dummyText As String
    Dim characterText As String

    JsonTemplateSkipWhitespace sourceText, currentPosition
    characterText = JsonTemplatePeekCharacter(sourceText, currentPosition)

    Select Case characterText
        Case """"
            dummyText = JsonTemplateParseString(sourceText, currentPosition)

        Case "{"
            currentPosition = currentPosition + 1
            JsonTemplateSkipWhitespace sourceText, currentPosition

            If JsonTemplatePeekCharacter(sourceText, currentPosition) <> "}" Then
                Do
                    dummyText = JsonTemplateParseString(sourceText, currentPosition)
                    JsonTemplateSkipWhitespace sourceText, currentPosition
                    JsonTemplateExpectCharacter sourceText, currentPosition, ":"
                    JsonTemplateSkipValue sourceText, currentPosition
                    JsonTemplateSkipWhitespace sourceText, currentPosition

                    If JsonTemplatePeekCharacter(sourceText, currentPosition) = "," Then
                        currentPosition = currentPosition + 1
                        JsonTemplateSkipWhitespace sourceText, currentPosition
                    Else
                        Exit Do
                    End If
                Loop
            End If

            JsonTemplateExpectCharacter sourceText, currentPosition, "}"

        Case "["
            currentPosition = currentPosition + 1
            JsonTemplateSkipWhitespace sourceText, currentPosition

            If JsonTemplatePeekCharacter(sourceText, currentPosition) <> "]" Then
                Do
                    JsonTemplateSkipValue sourceText, currentPosition
                    JsonTemplateSkipWhitespace sourceText, currentPosition

                    If JsonTemplatePeekCharacter(sourceText, currentPosition) = "," Then
                        currentPosition = currentPosition + 1
                        JsonTemplateSkipWhitespace sourceText, currentPosition
                    Else
                        Exit Do
                    End If
                Loop
            End If

            JsonTemplateExpectCharacter sourceText, currentPosition, "]"

        Case Else
            Do While currentPosition <= Len(sourceText)
                characterText = Mid$(sourceText, currentPosition, 1)
                If characterText = "," Or characterText = "]" Or characterText = "}" _
                   Or characterText = " " Or characterText = vbTab _
                   Or characterText = vbCr Or characterText = vbLf Then
                    Exit Do
                End If
                currentPosition = currentPosition + 1
            Loop
    End Select
End Sub

Private Sub JsonTemplateSkipWhitespace( _
    ByVal sourceText As String, _
    ByRef currentPosition As Long)

    Dim characterText As String

    Do While currentPosition <= Len(sourceText)
        characterText = Mid$(sourceText, currentPosition, 1)

        If characterText <> " " And characterText <> vbTab _
           And characterText <> vbCr And characterText <> vbLf Then
            Exit Do
        End If

        currentPosition = currentPosition + 1
    Loop
End Sub

Private Function JsonTemplatePeekCharacter( _
    ByVal sourceText As String, _
    ByVal currentPosition As Long) As String

    If currentPosition >= 1 And currentPosition <= Len(sourceText) Then
        JsonTemplatePeekCharacter = Mid$(sourceText, currentPosition, 1)
    Else
        JsonTemplatePeekCharacter = vbNullString
    End If
End Function

Private Sub JsonTemplateExpectCharacter( _
    ByVal sourceText As String, _
    ByRef currentPosition As Long, _
    ByVal expectedCharacter As String)

    JsonTemplateSkipWhitespace sourceText, currentPosition

    If JsonTemplatePeekCharacter(sourceText, currentPosition) <> expectedCharacter Then
        JsonTemplateRaiseParseError currentPosition, _
            "「" & expectedCharacter & "」が必要です。"
    End If

    currentPosition = currentPosition + 1
End Sub

Private Sub JsonTemplateEnsureEnd( _
    ByVal sourceText As String, _
    ByRef currentPosition As Long)

    JsonTemplateSkipWhitespace sourceText, currentPosition

    If currentPosition <= Len(sourceText) Then
        JsonTemplateRaiseParseError currentPosition, _
            "JSON末尾に不要な文字があります。"
    End If
End Sub

Private Sub JsonTemplateRaiseParseError( _
    ByVal currentPosition As Long, _
    ByVal messageText As String)

    Err.Raise vbObjectError + 2199, , _
              messageText & vbCrLf & "文字位置: " & CStr(currentPosition)
End Sub

Private Sub AddLabel(ByVal frm As Object, ByVal nm As String, ByVal cap As String, _
                     ByVal x As Single, ByVal y As Single, ByVal w As Single, _
                     ByVal h As Single, ByVal isBold As Boolean)
    Dim c As Object
    Set c = frm.Controls.Add("Forms.Label.1", nm, True)
    With c
        .Caption = cap: .Left = x: .Top = y: .Width = w: .Height = h
        .Font.Name = "Yu Gothic UI": .Font.Size = 10: .Font.Bold = isBold
        .BackStyle = 0: .ForeColor = RGB(52, 37, 44)
    End With
End Sub

Private Sub AddHeaderBar(ByVal frm As Object, ByVal x As Single, ByVal y As Single, _
                         ByVal w As Single, ByVal h As Single)
    Dim c As Object
    Set c = frm.Controls.Add("Forms.Label.1", "lblHeaderBar", True)
    With c
        .Caption = "": .Left = x: .Top = y: .Width = w: .Height = h
        .BackStyle = 1: .BackColor = RGB(229, 198, 210): .BorderStyle = 0
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
        .BackStyle = 0: .ForeColor = RGB(59, 31, 43): .BorderStyle = 0
    End With
End Sub

Private Sub AddHeaderDivider(ByVal frm As Object, ByVal nm As String, _
                             ByVal x As Single, ByVal y As Single, ByVal h As Single)
    Dim c As Object
    Set c = frm.Controls.Add("Forms.Label.1", nm, True)
    With c
        .Caption = "": .Left = x: .Top = y + 3: .Width = 1: .Height = h - 6
        .BackStyle = 1: .BackColor = RGB(177, 135, 153): .BorderStyle = 0
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
        .BackColor = RGB(255, 251, 252): .ForeColor = RGB(34, 30, 32)
        .BorderStyle = 1: .BorderColor = RGB(190, 158, 172)
        .MultiLine = multi
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
        .BackStyle = 1: .BackColor = RGB(255, 251, 252)
        .ForeColor = RGB(121, 89, 102)
    End With
End Sub

Private Sub AddCheckBox(ByVal frm As Object, ByVal nm As String, ByVal cap As String, _
                        ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single)
    Dim c As Object
    Set c = frm.Controls.Add("Forms.CheckBox.1", nm, True)
    With c
        .Caption = cap: .Left = x: .Top = y: .Width = w: .Height = h
        .Font.Name = "Yu Gothic UI": .Font.Size = 12: .Font.Bold = True
        .BackColor = RGB(246, 237, 241): .ForeColor = RGB(52, 37, 44)
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
        .BackColor = bg: .ForeColor = RGB(59, 31, 43)
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
        .BackColor = RGB(255, 251, 252): .ForeColor = RGB(34, 30, 32)
        .BorderStyle = 1: .BorderColor = RGB(190, 158, 172)
        .ColumnCount = 4
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
    AddLine s, "  JsonTemplateInitializeCurrentUserName"
    AddLine s, "  JsonTemplateLoadLatest True, True"
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
    AddLine s, "  targetControl.BackColor = RGB(255, 251, 252)"
    AddLine s, "  On Error GoTo 0"
    AddLine s, "  MsgBox " & q & "コピーできませんでした。もう一度お試しください。" & q & ", vbExclamation"
    AddLine s, "End Sub"

    AddLine s, "Private Sub FlashCopyHighlight(ByVal targetControl As Object)"
    AddLine s, "  Dim started As Single, elapsed As Double"
    AddLine s, "  targetControl.BackColor = RGB(248, 225, 142)"
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
    AddLine s, "  targetControl.BackColor = RGB(255, 251, 252)"
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
    AddLine s, "  Dim templateIndex As Long"
    AddLine s, "  If lstTemplate.ListIndex < 0 Then Exit Sub"
    AddLine s, "  templateIndex = CLng(lstTemplate.List(lstTemplate.ListIndex, 3))"
    AddLine s, "  txtSubject.Value = JsonTemplateSubject(templateIndex)"
    AddLine s, "  txtBody.Value = JsonTemplateReplaceUserName(JsonTemplateBody(templateIndex))"
    AddLine s, "End Sub"

    AddLine s, "Private Sub RefreshList()"
    AddLine s, "  Dim templateCount As Long, templateIndex As Long, n As Long"
    AddLine s, "  Dim data() As Variant, targetDate As Date, holidayMode As Boolean"
    AddLine s, "  Dim idVal As Variant, subject As String, body As String, tag As String"
    AddLine s, "  Dim keyword As String, markFilter As Boolean, markOK As Boolean"
    AddLine s, "  Dim i As Long, j As Long, tmp As Variant"
    AddLine s, "  On Error GoTo TemplateError"
    AddLine s, "  lstTemplate.Clear: txtSubject.Value = " & q & q & ": txtBody.Value = " & q & q
    AddLine s, "  If Not IsDate(txtTargetDate.Value) Then"
    AddLine s, "    lblCount.Caption = " & q & "0件" & q
    AddLine s, "    Exit Sub"
    AddLine s, "  End If"
    AddLine s, "  targetDate = DateValue(CDate(txtTargetDate.Value))"
    AddLine s, "  holidayMode = IsHolidayDate(targetDate)"
    AddLine s, "  templateCount = JsonTemplateCount()"
    AddLine s, "  keyword = Trim$(CStr(txtSearch.Value))"
    AddLine s, "  markFilter = (chkMark1.Value Or chkMark2.Value Or chkMark4.Value)"
    AddLine s, "  ReDim data(1 To Application.Max(1, templateCount), 1 To 4)"
    AddLine s, "  For templateIndex = 1 To templateCount"
    AddLine s, "    idVal = JsonTemplateDisplayID(templateIndex)"
    AddLine s, "    subject = JsonTemplateSubject(templateIndex)"
    AddLine s, "    body = JsonTemplateBody(templateIndex)"
    AddLine s, "    tag = JsonTemplateTagText(templateIndex)"
    AddLine s, "    If Len(Trim$(CStr(idVal))) = 0 Then GoTo ContinueTemplate"
    AddLine s, "    If holidayMode Then"
    AddLine s, "      If InStr(1, tag, " & q & "平日のみ" & q & ", vbTextCompare) > 0 Then GoTo ContinueTemplate"
    AddLine s, "    Else"
    AddLine s, "      If InStr(1, tag, " & q & "土日のみ" & q & ", vbTextCompare) > 0 Then GoTo ContinueTemplate"
    AddLine s, "    End If"
    AddLine s, "    If chkSVOnly.Value Then"
    AddLine s, "      If InStr(1, tag, " & q & "SV限" & q & ", vbTextCompare) = 0 Then GoTo ContinueTemplate"
    AddLine s, "    Else"
    AddLine s, "      If InStr(1, tag, " & q & "SV限" & q & ", vbTextCompare) > 0 Then GoTo ContinueTemplate"
    AddLine s, "    End If"
    AddLine s, "    If keyword <> " & q & q & " Then"
    AddLine s, "      If InStr(1, body, keyword, vbTextCompare) = 0 And InStr(1, tag, keyword, vbTextCompare) = 0 Then GoTo ContinueTemplate"
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
    AddLine s, "      If Not markOK Then GoTo ContinueTemplate"
    AddLine s, "    End If"
    AddLine s, "    n = n + 1: data(n, 1) = idVal: data(n, 2) = subject: data(n, 3) = tag: data(n, 4) = templateIndex"
    AddLine s, "ContinueTemplate:"
    AddLine s, "  Next templateIndex"

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
    AddLine s, "TemplateError:"
    AddLine s, "  lblCount.Caption = " & q & "0件" & q
    AddLine s, "  MsgBox " & q & "テンプレート.jsonまたはシート「休日マスタ」を確認してください。" & q & " & vbCrLf & Err.Description, vbExclamation"
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