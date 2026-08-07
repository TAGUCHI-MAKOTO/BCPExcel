Attribute VB_Name = "modTemplateJsonTeamsUrl"
Option Explicit

' ============================================================
' メンションテンプレ JSON Teams / SharePoint URL連携
'
' 【目的】
'   Teamsチャネルへ配置したテンプレート.jsonをURLから取得し、
'   既存の「メンションテンプレ_JSON自動読込.bas」で利用します。
'
' 【初回】
'   1. 「メンションテンプレ_JSON自動読込.bas」と本basを同じブックへインポート
'   2. 「TeamsJSON_URLを設定」を実行
'   3. 「メンションテンプレ_TeamsJSON版を開く」を実行
'
' 【今回の修正】
'   Teams版を開いた時にJSON版UserFormが未作成なら、
'   「メンションテンプレ_JSON版を作成」を自動実行してから開きます。
'
' 【注意】
'   Teamsチャネルのファイル実体はSharePointです。
'   Microsoft 365の対話型サインインが必要なURLは、
'   VBAのHTTP通信だけでは取得できない場合があります。
'   取得失敗時は既存のローカルJSONを残して使用します。
' ============================================================

Private Const CONFIG_NAME As String = "MentionTemplateJsonSourceUrl"
Private Const FORM_NAME As String = "frmTemplateSelectorJson"
Private Const HTTP_TIMEOUT_MS As Long = 15000

Private mTeamsJsonLastError As String
Private mTeamsJsonLastStatus As Long
Private mTeamsJsonLastUrl As String
Private mTeamsJsonLastTargetPath As String

' ============================================================
' 公開マクロ
' ============================================================
Public Sub TeamsJSON_URLを設定()
    Dim currentUrl As String
    Dim inputUrl As String

    currentUrl = TeamsJsonSourceUrl()

    inputUrl = Trim$(InputBox( _
        "TeamsチャネルのJSONファイルを「SharePointで開く」から開き、" & vbCrLf & _
        "SharePoint側のファイルURLを貼り付けてください。" & vbCrLf & vbCrLf & _
        "※ teams.microsoft.com の画面URLではなく、" & vbCrLf & _
        "   tenant.sharepoint.com 形式のURLを推奨します。", _
        "Teams JSON URL設定", _
        currentUrl))

    If Len(inputUrl) = 0 Then Exit Sub

    If Len(inputUrl) >= 250 Then
        MsgBox _
            "URLが長いため、入力欄では途中で切れる可能性があります。" & vbCrLf & vbCrLf & _
            "URLをExcelのセルへ貼り付け、そのセルを選択してから" & vbCrLf & _
            "「TeamsJSON_URLを選択セルから設定」を実行してください。", _
            vbExclamation, _
            "Teams JSON URL設定"
        Exit Sub
    End If

    If Not TeamsJsonValidateSourceUrl(inputUrl) Then Exit Sub

    TeamsJsonSaveSourceUrl inputUrl

    MsgBox _
        "JSONの参照URLを保存しました。", _
        vbInformation, _
        "Teams JSON URL設定"
End Sub

Public Sub TeamsJSON_URLを選択セルから設定()
    Dim inputUrl As String

    On Error GoTo CellError

    If TypeName(Selection) <> "Range" Then
        MsgBox _
            "SharePointのURLを入力したセルを1つ選択してから実行してください。", _
            vbExclamation, _
            "Teams JSON URL設定"
        Exit Sub
    End If

    If Selection.Cells.CountLarge <> 1 Then
        MsgBox _
            "URLを入力したセルを1つだけ選択してください。", _
            vbExclamation, _
            "Teams JSON URL設定"
        Exit Sub
    End If

    inputUrl = Trim$(CStr(Selection.Value2))

    If Len(inputUrl) = 0 Then
        MsgBox _
            "選択セルにURLが入力されていません。", _
            vbExclamation, _
            "Teams JSON URL設定"
        Exit Sub
    End If

    If Not TeamsJsonValidateSourceUrl(inputUrl) Then Exit Sub

    TeamsJsonSaveSourceUrl inputUrl

    MsgBox _
        "選択セルからJSONの参照URLを保存しました。", _
        vbInformation, _
        "Teams JSON URL設定"
    Exit Sub

CellError:
    MsgBox _
        "選択セルからURLを設定できませんでした。" & vbCrLf & vbCrLf & _
        Err.Description, _
        vbExclamation, _
        "Teams JSON URL設定"
End Sub

Public Sub TeamsJSON_URL設定をクリア()
    On Error Resume Next
    ThisWorkbook.Names(CONFIG_NAME).Delete
    On Error GoTo 0

    mTeamsJsonLastError = vbNullString
    mTeamsJsonLastStatus = 0
    mTeamsJsonLastUrl = vbNullString

    MsgBox _
        "Teams JSON URL設定をクリアしました。" & vbCrLf & _
        "以後は既存のローカルJSON読込のみ使用します。", _
        vbInformation, _
        "Teams JSON URL設定"
End Sub

Public Sub メンションテンプレ_TeamsJSON版を初期設定()
    If Not TeamsJsonEnsureUserForm() Then Exit Sub
    TeamsJSON_URLを設定
End Sub

Public Sub メンションテンプレ_TeamsJSON版を開く()
    Dim sourceUrl As String

    ' JSON版UserFormがまだ存在しない場合は自動作成する。
    If Not TeamsJsonEnsureUserForm() Then Exit Sub

    sourceUrl = TeamsJsonSourceUrl()

    If Len(sourceUrl) = 0 Then
        MsgBox _
            "Teams JSON URLが未設定です。" & vbCrLf & vbCrLf & _
            "先に「TeamsJSON_URLを設定」を実行してください。" & vbCrLf & _
            "今回は従来のローカルJSONでフォームを開きます。", _
            vbInformation, _
            "Teams JSON"
    ElseIf Not TeamsJsonSync(False) Then
        MsgBox _
            "Teams / SharePointから最新JSONを取得できませんでした。" & vbCrLf & vbCrLf & _
            mTeamsJsonLastError & vbCrLf & vbCrLf & _
            "既存のローカルJSONがある場合は、その内容でフォームを開きます。", _
            vbExclamation, _
            "Teams JSON"
    End If

    メンションテンプレ_JSON版を開く
End Sub

Public Sub TeamsJSONを今すぐ同期()
    If TeamsJsonSync(True) Then
        JsonTemplateLoadLatest False, True
    End If
End Sub

Public Sub TeamsJSON_接続状況を確認()
    Dim sourceUrl As String
    Dim normalizedUrl As String
    Dim targetPath As String
    Dim localState As String
    Dim formState As String
    Dim resultText As String

    On Error GoTo DiagnosticError

    sourceUrl = TeamsJsonSourceUrl()
    If Len(sourceUrl) > 0 Then normalizedUrl = TeamsJsonNormalizeDownloadUrl(sourceUrl)

    targetPath = JsonTemplateFilePath()
    If TeamsJsonFileExists(targetPath) Then
        localState = "あり"
    Else
        localState = "なし"
    End If

    If TeamsJsonUserFormExists() Then
        formState = "作成済み"
    Else
        formState = "未作成"
    End If

    resultText = _
        "【Teams JSON接続状況】" & vbCrLf & vbCrLf & _
        "UserForm: " & formState & vbCrLf & _
        "設定URL: " & IIf(Len(sourceUrl) > 0, sourceUrl, "未設定") & vbCrLf & _
        "取得URL: " & IIf(Len(normalizedUrl) > 0, normalizedUrl, "未設定") & vbCrLf & vbCrLf & _
        "ローカル保存先: " & targetPath & vbCrLf & _
        "ローカルJSON: " & localState & vbCrLf & vbCrLf & _
        "直近HTTP状態: " & IIf(mTeamsJsonLastStatus > 0, CStr(mTeamsJsonLastStatus), "未実行") & vbCrLf & _
        "直近エラー: " & IIf(Len(mTeamsJsonLastError) > 0, mTeamsJsonLastError, "なし")

    MsgBox resultText, vbInformation, "Teams JSON接続状況"
    Exit Sub

DiagnosticError:
    MsgBox _
        "Teams JSONの接続状況を確認できませんでした。" & vbCrLf & vbCrLf & _
        Err.Description, _
        vbExclamation, _
        "Teams JSON接続状況"
End Sub

' ============================================================
' UserFormの自動作成
' ============================================================
Private Function TeamsJsonEnsureUserForm() As Boolean
    On Error GoTo EnsureError

    If TeamsJsonUserFormExists() Then
        TeamsJsonEnsureUserForm = True
        Exit Function
    End If

    ' 元モジュールの既存作成処理をそのまま利用する。
    メンションテンプレ_JSON版を作成

    If Not TeamsJsonUserFormExists() Then
        Err.Raise vbObjectError + 2450, , _
                  "UserFormを作成できませんでした。"
    End If

    TeamsJsonEnsureUserForm = True
    Exit Function

EnsureError:
    TeamsJsonEnsureUserForm = False
    MsgBox _
        "テンプレート選択フォームを作成できませんでした。" & vbCrLf & vbCrLf & _
        Err.Description & vbCrLf & vbCrLf & _
        "Excelの「VBAプロジェクト オブジェクト モデルへのアクセスを信頼する」が" & vbCrLf & _
        "ONになっているか確認してください。", _
        vbExclamation, _
        "UserForm作成"
End Function

Private Function TeamsJsonUserFormExists() As Boolean
    Dim vbComp As Object

    On Error GoTo CheckError

    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If StrComp(vbComp.Name, FORM_NAME, vbTextCompare) = 0 Then
            TeamsJsonUserFormExists = True
            Exit Function
        End If
    Next vbComp

    Exit Function

CheckError:
    TeamsJsonUserFormExists = False
End Function

' ============================================================
' URL設定
' ============================================================
Public Function TeamsJsonSourceUrl() As String
    Dim nameObject As Name
    Dim refersText As String

    On Error Resume Next
    Set nameObject = ThisWorkbook.Names(CONFIG_NAME)
    On Error GoTo 0

    If nameObject Is Nothing Then Exit Function

    refersText = CStr(nameObject.RefersTo)
    If Left$(refersText, 1) = "=" Then refersText = Mid$(refersText, 2)

    If Len(refersText) >= 2 Then
        If Left$(refersText, 1) = """" And Right$(refersText, 1) = """" Then
            refersText = Mid$(refersText, 2, Len(refersText) - 2)
            refersText = Replace(refersText, """""", """")
        End If
    End If

    TeamsJsonSourceUrl = Trim$(refersText)
End Function

Private Sub TeamsJsonSaveSourceUrl(ByVal sourceUrl As String)
    Dim refersText As String

    sourceUrl = Trim$(sourceUrl)
    If Len(sourceUrl) = 0 Then Exit Sub

    refersText = "=""" & Replace(sourceUrl, """", """""") & """"

    On Error Resume Next
    ThisWorkbook.Names(CONFIG_NAME).Delete
    On Error GoTo SaveError

    ThisWorkbook.Names.Add _
        Name:=CONFIG_NAME, _
        RefersTo:=refersText, _
        Visible:=False
    Exit Sub

SaveError:
    Err.Raise vbObjectError + 2400, , _
              "JSON URL設定をブックへ保存できませんでした。" & vbCrLf & _
              Err.Description
End Sub

Private Function TeamsJsonValidateSourceUrl(ByVal sourceUrl As String) As Boolean
    sourceUrl = Trim$(sourceUrl)

    If InStr(1, sourceUrl, "teams.microsoft.com", vbTextCompare) > 0 Then
        MsgBox _
            "Teamsアプリの画面URLが入力されています。" & vbCrLf & vbCrLf & _
            "対象JSONを「SharePointで開く」で表示し、" & vbCrLf & _
            "SharePoint側のファイルURLを設定してください。", _
            vbExclamation, _
            "Teams JSON URL設定"
        Exit Function
    End If

    If InStr(1, sourceUrl, "http://", vbTextCompare) <> 1 _
       And InStr(1, sourceUrl, "https://", vbTextCompare) <> 1 Then
        MsgBox _
            "URLの形式が正しくありません。", _
            vbExclamation, _
            "Teams JSON URL設定"
        Exit Function
    End If

    TeamsJsonValidateSourceUrl = True
End Function

' ============================================================
' 同期本体
' ============================================================
Public Function TeamsJsonSync( _
    Optional ByVal showResult As Boolean = False) As Boolean

    Dim sourceUrl As String
    Dim requestUrl As String
    Dim targetPath As String
    Dim tempPath As String
    Dim jsonText As String

    On Error GoTo SyncError

    mTeamsJsonLastError = vbNullString
    mTeamsJsonLastStatus = 0
    mTeamsJsonLastUrl = vbNullString
    mTeamsJsonLastTargetPath = vbNullString

    sourceUrl = TeamsJsonSourceUrl()
    If Len(sourceUrl) = 0 Then
        Err.Raise vbObjectError + 2401, , _
                  "Teams JSON URLが未設定です。"
    End If

    requestUrl = TeamsJsonNormalizeDownloadUrl(sourceUrl)
    mTeamsJsonLastUrl = requestUrl

    jsonText = TeamsJsonDownloadUtf8(requestUrl)
    TeamsJsonValidateDownloadedText jsonText

    targetPath = JsonTemplateFilePath()
    mTeamsJsonLastTargetPath = targetPath
    TeamsJsonEnsureParentFolder targetPath

    tempPath = targetPath & ".download.tmp"
    TeamsJsonDeleteIfExists tempPath
    TeamsJsonWriteUtf8Text tempPath, jsonText

    TeamsJsonReplaceFileSafely tempPath, targetPath

    TeamsJsonSync = True

    If showResult Then
        MsgBox _
            "Teams / SharePointからJSONを同期しました。" & vbCrLf & vbCrLf & _
            "HTTP: " & CStr(mTeamsJsonLastStatus) & vbCrLf & _
            "保存先: " & targetPath, _
            vbInformation, _
            "Teams JSON同期"
    End If
    Exit Function

SyncError:
    mTeamsJsonLastError = Err.Description
    TeamsJsonDeleteIfExists tempPath
    TeamsJsonSync = False

    If showResult Then
        MsgBox _
            "Teams / SharePointからJSONを同期できませんでした。" & vbCrLf & vbCrLf & _
            mTeamsJsonLastError & vbCrLf & vbCrLf & _
            TeamsJsonAuthenticationAdvice(), _
            vbExclamation, _
            "Teams JSON同期"
    End If
End Function

' ============================================================
' HTTP取得
' ============================================================
Private Function TeamsJsonDownloadUtf8(ByVal requestUrl As String) As String
    Dim httpObject As Object
    Dim responseBody As Variant
    Dim firstError As String
    Dim secondError As String

    On Error GoTo XmlHttpError

    Set httpObject = CreateObject("MSXML2.XMLHTTP.6.0")
    With httpObject
        .Open "GET", requestUrl, False
        .setRequestHeader "Accept", "application/json,text/plain,*/*"
        .setRequestHeader "Cache-Control", "no-cache"
        .send

        mTeamsJsonLastStatus = CLng(.Status)
        If mTeamsJsonLastStatus < 200 Or mTeamsJsonLastStatus >= 300 Then
            Err.Raise vbObjectError + 2410, , _
                      "HTTP " & CStr(mTeamsJsonLastStatus) & " が返されました。"
        End If

        responseBody = .responseBody
    End With

    TeamsJsonDownloadUtf8 = TeamsJsonBytesToUtf8(responseBody)
    Exit Function

XmlHttpError:
    firstError = Err.Description
    Err.Clear
    Set httpObject = Nothing

    On Error GoTo WinHttpError

    Set httpObject = CreateObject("WinHttp.WinHttpRequest.5.1")
    With httpObject
        .Open "GET", requestUrl, False
        .SetTimeouts HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS
        .SetAutoLogonPolicy 0
        .SetRequestHeader "Accept", "application/json,text/plain,*/*"
        .SetRequestHeader "Cache-Control", "no-cache"
        .Send

        mTeamsJsonLastStatus = CLng(.Status)
        If mTeamsJsonLastStatus < 200 Or mTeamsJsonLastStatus >= 300 Then
            Err.Raise vbObjectError + 2411, , _
                      "HTTP " & CStr(mTeamsJsonLastStatus) & " が返されました。"
        End If

        responseBody = .ResponseBody
    End With

    TeamsJsonDownloadUtf8 = TeamsJsonBytesToUtf8(responseBody)
    Exit Function

WinHttpError:
    secondError = Err.Description

    If mTeamsJsonLastStatus = 401 Or mTeamsJsonLastStatus = 403 Then
        Err.Raise vbObjectError + 2412, , _
                  "SharePoint側でMicrosoft 365認証が必要です（HTTP " & _
                  CStr(mTeamsJsonLastStatus) & "）。"
    End If

    Err.Raise vbObjectError + 2413, , _
              "URLからJSONを取得できませんでした。" & vbCrLf & _
              "XMLHTTP: " & firstError & vbCrLf & _
              "WinHTTP: " & secondError
End Function

Private Function TeamsJsonBytesToUtf8(ByVal responseBody As Variant) As String
    Dim streamObject As Object
    Dim loadedText As String

    Set streamObject = CreateObject("ADODB.Stream")

    With streamObject
        .Type = 1
        .Open
        .Write responseBody
        .Position = 0
        .Type = 2
        .Charset = "utf-8"
        loadedText = .ReadText(-1)
        .Close
    End With

    If Len(loadedText) > 0 Then
        If AscW(Left$(loadedText, 1)) = &HFEFF Then
            loadedText = Mid$(loadedText, 2)
        End If
    End If

    TeamsJsonBytesToUtf8 = loadedText
End Function

Private Function TeamsJsonNormalizeDownloadUrl(ByVal sourceUrl As String) As String
    Dim normalizedUrl As String

    normalizedUrl = Trim$(sourceUrl)
    normalizedUrl = Replace(normalizedUrl, "&amp;", "&", 1, -1, vbTextCompare)

    normalizedUrl = Replace(normalizedUrl, "?web=1", "?download=1", 1, -1, vbTextCompare)
    normalizedUrl = Replace(normalizedUrl, "&web=1", "&download=1", 1, -1, vbTextCompare)

    If InStr(1, normalizedUrl, "download=1", vbTextCompare) = 0 Then
        If InStr(1, normalizedUrl, "?", vbBinaryCompare) > 0 Then
            normalizedUrl = normalizedUrl & "&download=1"
        Else
            normalizedUrl = normalizedUrl & "?download=1"
        End If
    End If

    TeamsJsonNormalizeDownloadUrl = normalizedUrl
End Function

Private Sub TeamsJsonValidateDownloadedText(ByVal jsonText As String)
    Dim trimmedText As String
    Dim lowerText As String

    trimmedText = Trim$(jsonText)

    If Len(trimmedText) = 0 Then
        Err.Raise vbObjectError + 2420, , _
                  "取得したファイルが空です。"
    End If

    lowerText = LCase$(Left$(trimmedText, 500))

    If Left$(trimmedText, 1) = "<" _
       Or InStr(1, lowerText, "<!doctype", vbTextCompare) > 0 _
       Or InStr(1, lowerText, "<html", vbTextCompare) > 0 _
       Or InStr(1, lowerText, "login.microsoftonline.com", vbTextCompare) > 0 Then
        Err.Raise vbObjectError + 2421, , _
                  "JSONではなくMicrosoft 365のサインイン画面またはWebページが返されました。"
    End If

    If Left$(trimmedText, 1) <> "[" Or Right$(trimmedText, 1) <> "]" Then
        Err.Raise vbObjectError + 2422, , _
                  "取得内容がテンプレートJSONの配列形式ではありません。"
    End If
End Sub

' ============================================================
' ファイル保存
' ============================================================
Private Sub TeamsJsonWriteUtf8Text(ByVal filePath As String, ByVal sourceText As String)
    Dim streamObject As Object

    Set streamObject = CreateObject("ADODB.Stream")

    With streamObject
        .Type = 2
        .Charset = "utf-8"
        .Open
        .WriteText sourceText
        .SaveToFile filePath, 2
        .Close
    End With
End Sub

Private Sub TeamsJsonReplaceFileSafely(ByVal tempPath As String, ByVal targetPath As String)
    Dim backupPath As String

    backupPath = targetPath & ".bak"

    On Error GoTo ReplaceError

    If TeamsJsonFileExists(targetPath) Then
        TeamsJsonDeleteIfExists backupPath
        FileCopy targetPath, backupPath
    End If

    TeamsJsonDeleteIfExists targetPath
    Name tempPath As targetPath
    Exit Sub

ReplaceError:
    On Error Resume Next
    If Not TeamsJsonFileExists(targetPath) And TeamsJsonFileExists(backupPath) Then
        FileCopy backupPath, targetPath
    End If
    On Error GoTo 0

    Err.Raise vbObjectError + 2430, , _
              "取得したJSONを保存先へ反映できませんでした。" & vbCrLf & _
              targetPath
End Sub

Private Sub TeamsJsonEnsureParentFolder(ByVal filePath As String)
    Dim fso As Object
    Dim parentFolder As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    parentFolder = fso.GetParentFolderName(filePath)

    If Len(parentFolder) = 0 Or Not fso.FolderExists(parentFolder) Then
        Err.Raise vbObjectError + 2431, , _
                  "JSONの保存フォルダーが見つかりません。" & vbCrLf & _
                  parentFolder
    End If
End Sub

Private Function TeamsJsonFileExists(ByVal filePath As String) As Boolean
    If Len(Trim$(filePath)) = 0 Then Exit Function

    On Error Resume Next
    TeamsJsonFileExists = _
        (Len(Dir$(filePath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) > 0)
    On Error GoTo 0
End Function

Private Sub TeamsJsonDeleteIfExists(ByVal filePath As String)
    If Len(Trim$(filePath)) = 0 Then Exit Sub

    On Error Resume Next
    If TeamsJsonFileExists(filePath) Then Kill filePath
    On Error GoTo 0
End Sub

Private Function TeamsJsonAuthenticationAdvice() As String
    If mTeamsJsonLastStatus = 401 Or mTeamsJsonLastStatus = 403 _
       Or InStr(1, mTeamsJsonLastError, "サインイン", vbTextCompare) > 0 _
       Or InStr(1, mTeamsJsonLastError, "認証", vbTextCompare) > 0 Then
        TeamsJsonAuthenticationAdvice = _
            "Microsoft 365認証が必要なリンクでは、VBA単体のURL取得が失敗する場合があります。" & vbCrLf & _
            "その場合はTeamsのチャネルフォルダーをOneDriveで同期し、従来のローカル読込方式を使用してください。"
    Else
        TeamsJsonAuthenticationAdvice = _
            "URL、ネットワーク接続、SharePoint側のアクセス権を確認してください。"
    End If
End Function
