param(
    [Parameter(Mandatory=$true)]
    [string]$PendingPath,

    [string]$CsvFolder = "",

    [int]$RetryMinMilliseconds = 450,

    [int]$RetryMaxMilliseconds = 1250,

    [int]$MaxRetrySeconds = 300
)

$ErrorActionPreference = "Stop"

function Show-Balloon {
    param(
        [string]$Title,
        [string]$Text,
        [ValidateSet("Info","Warning","Error")]
        [string]$Kind = "Info"
    )

    try {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        $notify = New-Object System.Windows.Forms.NotifyIcon
        $notify.Icon = [System.Drawing.SystemIcons]::Information
        $notify.Visible = $true
        $notify.BalloonTipTitle = $Title
        $notify.BalloonTipText = $Text

        switch ($Kind) {
            "Warning" { $notify.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Warning }
            "Error"   { $notify.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Error }
            default   { $notify.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info }
        }

        $notify.ShowBalloonTip(5000)
        Start-Sleep -Seconds 5
        $notify.Visible = $false
        $notify.Dispose()
    }
    catch {
        # 通知失敗は送信結果へ影響させない
    }
}

function Write-LocalLog {
    param(
        [string]$Status,
        [string]$RequestId,
        [string]$BaseName,
        [string]$Message = ""
    )

    try {
        $root = Join-Path $env:LOCALAPPDATA "MentionRequest\Logs"

        if (-not (Test-Path -LiteralPath $root)) {
            New-Item -ItemType Directory -Path $root -Force | Out-Null
        }

        $path = Join-Path $root ("worker_" + (Get-Date -Format "yyyyMMdd") + ".log")
        $line = "{0}`t{1}`t{2}`t{3}`t{4}" -f `
            (Get-Date -Format "yyyy/MM/dd HH:mm:ss.fff"),
            $Status,
            $RequestId,
            $BaseName,
            $Message

        Add-Content -LiteralPath $path -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
    }
    catch {}
}

function Decode-Value([string]$Text) {
    return [System.Uri]::UnescapeDataString($Text)
}

function Read-Pending {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Pendingファイルが見つかりません。"
    }

    $text = Get-Content -LiteralPath $Path -Raw
    $rows = $text -replace "`r`n","`n" -replace "`r","`n" -split "`n"

    $format = ""
    $baseName = ""
    $requestId = ""
    $created = ""
    $pendingCsvFolder = ""
    $lines = New-Object System.Collections.Generic.List[string]

    if ($rows.Count -gt 0 -and $rows[0] -eq "MENTION_REQUEST_PENDING_V1") {
        $format = "MENTION_REQUEST_PENDING_V1"
    }

    foreach ($row in $rows) {
        if ($row.StartsWith("FORMAT=")) {
            $format = $row.Substring(7)
        }
        elseif ($row.StartsWith("BASE=")) {
            $baseName = Decode-Value $row.Substring(5)
        }
        elseif ($row.StartsWith("REQUEST_ID=")) {
            $requestId = Decode-Value $row.Substring(11)
        }
        elseif ($row.StartsWith("CREATED=")) {
            $created = Decode-Value $row.Substring(8)
        }
        elseif ($row.StartsWith("CSV_FOLDER=")) {
            $pendingCsvFolder = Decode-Value $row.Substring(11)
        }
        elseif ($row.StartsWith("LINE=")) {
            [void]$lines.Add((Decode-Value $row.Substring(5)))
        }
    }

    if ($format -ne "MENTION_REQUEST_PENDING_V1" -and
        $format -ne "MENTION_REQUEST_PENDING_V2" -and
        $format -ne "MENTION_REQUEST_PENDING_V3") {
        throw "Pending形式が不正です。"
    }

    if ([string]::IsNullOrWhiteSpace($baseName)) {
        throw "拠点情報がありません。"
    }

    if ([string]::IsNullOrWhiteSpace($requestId)) {
        throw "RequestIDがありません。"
    }

    if ($lines.Count -eq 0) {
        throw "CSV書き込みデータがありません。"
    }

    return [PSCustomObject]@{
        Format     = $format
        BaseName   = $baseName
        RequestId  = $requestId
        Created    = $created
        CsvFolder  = $pendingCsvFolder
        Lines      = $lines
    }
}

function Ensure-Folder {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Get-SafeBaseName {
    param([string]$Text)

    foreach ($c in [IO.Path]::GetInvalidFileNameChars()) {
        $Text = $Text.Replace([string]$c, "_")
    }

    return $Text
}

function Get-DateKeyFromPending {
    param([string]$Created)

    if ($Created -match '^(\d{4})/(\d{2})/(\d{2})') {
        return ($Matches[1] + $Matches[2] + $Matches[3])
    }

    return (Get-Date -Format "yyyyMMdd")
}

function Get-CsvPath {
    param(
        [string]$Folder,
        [string]$BaseName,
        [string]$Created
    )

    $day = Get-DateKeyFromPending $Created
    $safe = Get-SafeBaseName $BaseName

    return Join-Path $Folder ($safe + "_" + $day + ".csv")
}

function Get-RetryDelay {
    $min = [Math]::Max(50,$RetryMinMilliseconds)
    $max = [Math]::Max($min + 1,$RetryMaxMilliseconds + 1)

    return Get-Random -Minimum $min -Maximum $max
}

function Acquire-Lock {
    param([string]$LockPath)

    $waitStarted = Get-Date
    $attemptLimitSeconds = [Math]::Min(5,[Math]::Max(1,$MaxRetrySeconds))

    while ($true) {
        try {
            # OpenOrCreate + FileShare.None：
            # lockファイルが残っていても、前プロセスのhandleが解放されていれば取得可能。
            return [System.IO.File]::Open(
                $LockPath,
                [System.IO.FileMode]::OpenOrCreate,
                [System.IO.FileAccess]::ReadWrite,
                [System.IO.FileShare]::None
            )
        }
        catch {
            if (((Get-Date) - $waitStarted).TotalSeconds -ge $attemptLimitSeconds) {
                throw "CSVの排他ロックを取得できませんでした。"
            }

            Start-Sleep -Milliseconds (Get-RetryDelay)
        }
    }
}

function New-CsvParserFromPath {
    param([string]$Path)

    Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction SilentlyContinue

    $parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser(
        $Path,
        [System.Text.Encoding]::GetEncoding(932)
    )

    $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
    $parser.SetDelimiters(",")
    $parser.HasFieldsEnclosedInQuotes = $true

    return $parser
}

function Parse-CsvRecord {
    param([string]$Record)

    Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction SilentlyContinue

    $reader = New-Object System.IO.StringReader($Record)
    $parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($reader)
    $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
    $parser.SetDelimiters(",")
    $parser.HasFieldsEnclosedInQuotes = $true

    try {
        return $parser.ReadFields()
    }
    finally {
        $parser.Close()
        $reader.Dispose()
    }
}

function Get-ExistingRequestNumbers {
    param(
        [string]$CsvPath,
        [string]$RequestId
    )

    $numbers = @{}

    if (-not (Test-Path -LiteralPath $CsvPath)) {
        return $numbers
    }

    $parser = New-CsvParserFromPath $CsvPath

    try {
        $isHeader = $true

        while (-not $parser.EndOfData) {
            $fields = $parser.ReadFields()

            if ($isHeader) {
                $isHeader = $false
                continue
            }

            if ($fields -and $fields.Length -ge 5) {
                if ([string]$fields[1] -eq $RequestId) {
                    $requestNo = [string]$fields[4]

                    if (-not [string]::IsNullOrWhiteSpace($requestNo)) {
                        $numbers[$requestNo] = $true
                    }
                }
            }
        }
    }
    finally {
        $parser.Close()
    }

    return $numbers
}

function Get-PendingRequestLineMap {
    param(
        [System.Collections.Generic.List[string]]$Lines
    )

    $map = @{}

    foreach ($line in $Lines) {
        $fields = Parse-CsvRecord $line

        if (-not $fields -or $fields.Length -lt 5) {
            throw "Pending内CSV行の解析に失敗しました。"
        }

        $requestNo = [string]$fields[4]

        if ([string]::IsNullOrWhiteSpace($requestNo)) {
            throw "依頼番号を取得できません。"
        }

        $map[$requestNo] = $line
    }

    return $map
}

function Write-CsvWithLock {
    param(
        [string]$Folder,
        [string]$BaseName,
        [string]$RequestId,
        [string]$Created,
        [System.Collections.Generic.List[string]]$Lines
    )

    Ensure-Folder $Folder

    $csvPath = Get-CsvPath `
        -Folder $Folder `
        -BaseName $BaseName `
        -Created $Created

    $lockPath = $csvPath + ".lock"
    $lockStream = $null

    try {
        $lockStream = Acquire-Lock $lockPath

        $exists = Test-Path -LiteralPath $csvPath
        $existingNumbers = Get-ExistingRequestNumbers `
            -CsvPath $csvPath `
            -RequestId $RequestId

        $pendingMap = Get-PendingRequestLineMap -Lines $Lines
        $missingLines = New-Object System.Collections.Generic.List[string]

        foreach ($requestNo in $pendingMap.Keys) {
            if (-not $existingNumbers.ContainsKey($requestNo)) {
                [void]$missingLines.Add([string]$pendingMap[$requestNo])
            }
        }

        if ($missingLines.Count -eq 0) {
            return "DUPLICATE_SKIP"
        }

        $wasPartial = ($existingNumbers.Count -gt 0)

        $encoding = [System.Text.Encoding]::GetEncoding(932)
        $writer = New-Object System.IO.StreamWriter($csvPath,$true,$encoding)

        try {
            if (-not $exists) {
                $writer.WriteLine("送信日時,RequestID,拠点,依頼者,依頼番号,組織,CA名,代理CA組織,代理CA1,代理CA2,代理CA3,メールメモ,処理日時,期日,時短,至急,タイプ")
            }

            foreach ($line in $missingLines) {
                $writer.WriteLine($line)
            }

            $writer.Flush()

            try {
                $writer.BaseStream.Flush($true)
            }
            catch {
                # ネットワーク共有等でFlush(bool)が使えない場合は通常Flush済みを採用
            }
        }
        finally {
            $writer.Dispose()
        }

        # 同じlockを保持したまま、依頼番号単位で全件書けたことを再確認
        $verifiedNumbers = Get-ExistingRequestNumbers `
            -CsvPath $csvPath `
            -RequestId $RequestId

        foreach ($requestNo in $pendingMap.Keys) {
            if (-not $verifiedNumbers.ContainsKey($requestNo)) {
                throw ("CSV書き込み確認に失敗しました。依頼番号：" + $requestNo)
            }
        }

        if ($wasPartial) {
            return "RECOVERED_PARTIAL"
        }

        return "SUCCESS"
    }
    finally {
        if ($lockStream) {
            $lockStream.Dispose()
        }

        # handle解放後に削除。別workerが先に取得した場合は削除失敗するので安全。
        Remove-Item -LiteralPath $lockPath -Force -ErrorAction SilentlyContinue
    }
}

function Remove-PendingAfterSuccess {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        return $true
    }

    for ($i=0; $i -lt 5; $i++) {
        try {
            Remove-Item -LiteralPath $Path -Force -ErrorAction Stop
            return $true
        }
        catch {
            Start-Sleep -Milliseconds (Get-RetryDelay)
        }
    }

    return $false
}

# ============================================================
# Main
# ============================================================
$pending = $null

try {
    $pending = Read-Pending $PendingPath

    # v9以降はアプリ側の固定CSV保存先を優先。
    # 旧Pending互換のため、引数が空の場合のみPending内の保存先をfallbackに使う。
    $effectiveCsvFolder = [string]$CsvFolder

    if ([string]::IsNullOrWhiteSpace($effectiveCsvFolder)) {
        $effectiveCsvFolder = [string]$pending.CsvFolder
    }

    if ([string]::IsNullOrWhiteSpace($effectiveCsvFolder)) {
        throw "CSV格納先が設定されていません。"
    }

    Write-LocalLog `
        -Status "START" `
        -RequestId $pending.RequestId `
        -BaseName $pending.BaseName `
        -Message $effectiveCsvFolder

    $startedAt = Get-Date
    $writeResult = ""

    # 共有先一時断・Excel使用中等では一定時間バックグラウンド自動再送。
    # 上限を超えた場合はPendingを残したまま通知して終了する。
    while ($true) {
        try {
            $writeResult = Write-CsvWithLock `
                -Folder $effectiveCsvFolder `
                -BaseName $pending.BaseName `
                -RequestId $pending.RequestId `
                -Created $pending.Created `
                -Lines $pending.Lines

            break
        }
        catch {
            $lastError = $_.Exception.Message
            $elapsedSeconds = ((Get-Date) - $startedAt).TotalSeconds

            Write-LocalLog `
                -Status "RETRY" `
                -RequestId $pending.RequestId `
                -BaseName $pending.BaseName `
                -Message $lastError

            if ($elapsedSeconds -ge $MaxRetrySeconds) {
                Write-LocalLog `
                    -Status "RETRY_EXHAUSTED" `
                    -RequestId $pending.RequestId `
                    -BaseName $pending.BaseName `
                    -Message $lastError

                Show-Balloon `
                    -Title "メンション依頼｜未送信" `
                    -Text "共有CSVへ書き込めませんでした。未送信データはデスクトップに保持しています。フォームを開いて「未送信を再送」を押してください。" `
                    -Kind "Error"

                exit 2
            }

            Start-Sleep -Milliseconds (Get-RetryDelay)
        }
    }

    $pendingRemoved = Remove-PendingAfterSuccess -Path $PendingPath

    # Pendingフォルダが空なら削除
    try {
        $pendingFolder = Split-Path -Parent $PendingPath

        if ((Test-Path -LiteralPath $pendingFolder) -and
            -not (Get-ChildItem -LiteralPath $pendingFolder -Force | Select-Object -First 1)) {
            Remove-Item -LiteralPath $pendingFolder -Force
        }
    }
    catch {}

    Write-LocalLog `
        -Status $writeResult `
        -RequestId $pending.RequestId `
        -BaseName $pending.BaseName `
        -Message ("PendingRemoved=" + $pendingRemoved)

    if ($writeResult -eq "DUPLICATE_SKIP") {
        Show-Balloon `
            -Title "メンション依頼｜送信済み確認" `
            -Text ("同じ依頼がすでにCSVへ反映済みのため、重複書き込みせず完了しました。 RequestID: " + $pending.RequestId) `
            -Kind "Info"
    }
    elseif ($writeResult -eq "RECOVERED_PARTIAL") {
        Show-Balloon `
            -Title "メンション依頼｜送信完了" `
            -Text ("途中まで反映済みの依頼を補完して、CSVへの書き込みが完了しました。 RequestID: " + $pending.RequestId) `
            -Kind "Info"
    }
    else {
        Show-Balloon `
            -Title "メンション依頼｜送信完了" `
            -Text ("CSVへの書き込みが完了しました。 RequestID: " + $pending.RequestId) `
            -Kind "Info"
    }

    exit 0
}
catch {
    if ($pending) {
        Write-LocalLog `
            -Status "FATAL" `
            -RequestId $pending.RequestId `
            -BaseName $pending.BaseName `
            -Message $_.Exception.Message
    }

    Show-Balloon `
        -Title "メンション依頼｜送信処理エラー" `
        -Text "送信処理を開始できませんでした。Pendingデータは削除せず保持しています。" `
        -Kind "Error"

    exit 1
}
