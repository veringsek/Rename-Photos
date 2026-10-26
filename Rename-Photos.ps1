
# Rename-Photo.ps1
# Load from $PROFILE:
# . "D:\Scripts\Rename-Photo.ps1"

$script:RenamePhotoExtensions = @{
    Photos = @(
        '.jpg', '.jpeg', '.jpe', '.jfif',
        '.png', '.gif', '.bmp', '.tif', '.tiff',
        '.webp', '.heic', '.heif', '.avif',
        '.dng', '.raw', '.arw', '.cr2', '.cr3',
        '.nef', '.nrw', '.orf', '.rw2', '.raf',
        '.pef', '.srw', '.ico'
    )
    Videos = @(
        '.mp4', '.m4v', '.mov', '.avi', '.wmv',
        '.mkv', '.webm', '.mpg', '.mpeg', '.mpe',
        '.m2v', '.ts', '.mts', '.m2ts', '.3gp',
        '.3g2', '.flv', '.f4v', '.vob', '.ogv',
        '.ogg', '.mxf', '.asf', '.rm', '.rmvb',
        '.divx'
    )
}

function Get-RenamePhotoDate {
    param(
        [Parameter(Mandatory)]
        [System.IO.FileInfo]$File,

        [Parameter(Mandatory)]
        [string]$According
    )

    foreach ($method in $According.ToCharArray()) {
        switch ($method) {
            'E' {
                $shell = $null
                $folder = $null
                $item = $null

                try {
                    $shell = New-Object -ComObject Shell.Application
                    $folder = $shell.Namespace($File.DirectoryName)
                    $item = $folder.ParseName($File.Name)

                    if ($null -ne $item) {
                        $value = $item.ExtendedProperty(
                            'System.Photo.DateTaken'
                        )

                        $date = [datetime]::MinValue
                        $valid = $false

                        if ($value -is [datetime]) {
                            $date = $value
                            $valid = $true
                        }
                        elseif ($value -is [string]) {
                            $valid = [datetime]::TryParse(
                                $value,
                                [Globalization.CultureInfo]::InvariantCulture,
                                [Globalization.DateTimeStyles]::None,
                                [ref]$date
                            )
                        }

                        if ($valid -and $date -ne [datetime]::MinValue) {
                            return [pscustomobject]@{
                                Date   = $date
                                Source = 'EXIF (E)'
                                Color  = [ConsoleColor]::Green
                            }
                        }
                    }
                }
                catch {
                    # EXIF 無法讀取，繼續下一方案
                }
                finally {
                    foreach ($obj in @($item, $folder, $shell)) {
                        if ($null -ne $obj -and
                            [Runtime.InteropServices.Marshal]::IsComObject($obj)) {
                            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject(
                                $obj
                            )
                        }
                    }
                }
            }

            'N' {
                $pattern = '(P|V|Screenshot|Screen_Recording|IMG|螢幕擷取畫面)?[_ ]?(\d{4})[-.]?(\d\d)[-.]?(\d\d)(_|-| )(\d\d)[-.]?(\d\d)[-.]?(\d\d)(.+)?'

                if ($File.BaseName -match $pattern) {
                    try {
                        $dateText = '{0}-{1}-{2} {3}:{4}:{5}' -f `
                            $Matches[2], $Matches[3], $Matches[4], `
                            $Matches[6], $Matches[7], $Matches[8]

                        $date = [datetime]::ParseExact(
                            $dateText,
                            'yyyy-MM-dd HH:mm:ss',
                            [Globalization.CultureInfo]::InvariantCulture
                        )

                        return [pscustomobject]@{
                            Date   = $date
                            Source = '檔名匹配 (N)'
                            Color  = [ConsoleColor]::Blue
                        }
                    }
                    catch {
                        # 日期或時間不合法，繼續下一方案
                    }
                }
            }

            'M' {
                return [pscustomobject]@{
                    Date   = $File.LastWriteTime
                    Source = '修改日期 (M)'
                    Color  = [ConsoleColor]::DarkYellow
                }
            }
        }
    }

    throw "指定的日期方案均無法取得有效日期：$($File.FullName)"
}

function Rename-Photo {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)]
        [ValidateSet('rename')]
        [string]$Action = 'rename',

        [Parameter(Mandatory)]
        [string]$Path,

        [string]$Output,

        [string]$According = 'ENM',

        [ValidateSet('Skip', 'Numbering', 'Origin')]
        [string]$Duplicate = 'Skip',

        [switch]$DitchOrigin,

        [string]$Ext = 'Photos,Videos',

        [bool]$Recursive = $true,

        [bool]$Immediate = $false
    )

    $ErrorActionPreference = 'Stop'

    # 驗證日期方案
    $According = $According.ToUpperInvariant()

    if ($According -notmatch '^[ENM]+$') {
        throw '-According 只能包含 E、N、M，例如 ENM、EN、MN 或 M。'
    }

    $letters = @($According.ToCharArray())
    $uniqueLetters = @($letters | Select-Object -Unique)

    if ($uniqueLetters.Count -ne $letters.Count) {
        throw '-According 中的 E、N、M 不可重複。'
    }

    # 解析輸入路徑
    $inputItem = Get-Item -LiteralPath $Path -ErrorAction Stop
    $isDirectory = $inputItem -is [System.IO.DirectoryInfo]
    $inPlace = [string]::IsNullOrWhiteSpace($Output)

    # 解析 Ext 列表及別名
    $extensions = [System.Collections.Generic.HashSet[string]]::new(
        [StringComparer]::OrdinalIgnoreCase
    )

    foreach ($entry in $Ext.Split(',')) {
        $token = $entry.Trim()

        if ([string]::IsNullOrWhiteSpace($token)) {
            continue
        }

        switch -Regex ($token) {
            '^Photos$' {
                foreach ($extension in $script:RenamePhotoExtensions.Photos) {
                    [void]$extensions.Add($extension)
                }
                break
            }

            '^Videos$' {
                foreach ($extension in $script:RenamePhotoExtensions.Videos) {
                    [void]$extensions.Add($extension)
                }
                break
            }

            default {
                # 自訂副檔名，例如 jpg、.mp4、.custom
                if ($token -match '[\\/:*?"<>|]') {
                    throw "無效的副檔名：$token"
                }

                if (-not $token.StartsWith('.')) {
                    $token = ".$token"
                }

                [void]$extensions.Add($token)
            }
        }
    }

    # 檔案輸入不檢查副檔名；目錄輸入依 Ext 篩選
    if ($isDirectory) {
        $getChildParams = @{
            LiteralPath = $inputItem.FullName
            File        = $true
            ErrorAction = 'Stop'
        }

        if ($Recursive) {
            $getChildParams['Recurse'] = $true
        }

        $files = @(
            Get-ChildItem @getChildParams |
                Where-Object {
                    $extensions.Contains($_.Extension)
                } |
                Sort-Object FullName
        )

        $sourceRoot = $inputItem.FullName
    }
    elseif ($inputItem -is [System.IO.FileInfo]) {
        $files = @($inputItem)
        $sourceRoot = $inputItem.DirectoryName
    }
    else {
        throw "輸入路徑不是檔案或目錄：$Path"
    }

    if ($files.Count -eq 0) {
        Write-Host '找不到符合條件的檔案。' -ForegroundColor Yellow
        return
    }

    # 準備輸出目錄
    if (-not $inPlace) {
        $outputPath = $ExecutionContext.SessionState.Path.
            GetUnresolvedProviderPathFromPSPath($Output)

        if (Test-Path -LiteralPath $outputPath -PathType Leaf) {
            throw "輸出路徑必須是目錄，不能是檔案：$Output"
        }

        if (-not (Test-Path -LiteralPath $outputPath -PathType Container)) {
            New-Item -ItemType Directory -Path $outputPath -Force |
                Out-Null
        }

        $Output = $outputPath
    }

    # 第一階段：建立完整計畫，不執行檔案異動
    $plan = [System.Collections.Generic.List[object]]::new()
    $reserved = [System.Collections.Generic.HashSet[string]]::new(
        [StringComparer]::OrdinalIgnoreCase
    )

    foreach ($file in $files) {
        $record = [pscustomobject]@{
            Source      = $file.FullName
            Destination = $null
            SourceName  = $file.Name
            NewName     = $null
            DateSource  = ''
            Date        = $null
            Color       = [ConsoleColor]::Gray
            Status      = 'Pending'
            Reason      = ''
        }

        try {
            $photo = Get-RenamePhotoDate `
                -File $file `
                -According $According

            $record.Date = $photo.Date
            $record.DateSource = $photo.Source
            $record.Color = $photo.Color

            $timestamp = $photo.Date.ToString(
                'yyyy.MM.ddTHH.mm.ss',
                [Globalization.CultureInfo]::InvariantCulture
            )

            if ($DitchOrigin) {
                $baseName = "[$timestamp]"
            }
            else {
                $baseName = "[$timestamp][$($file.BaseName)]"
            }

            $extension = $file.Extension

            # Output 模式保留來源相對目錄結構
            if (-not $inPlace -and $isDirectory) {
                $relativeDirectory = [System.IO.Path]::GetRelativePath(
                    $sourceRoot,
                    $file.DirectoryName
                )

                if ($relativeDirectory -eq '.') {
                    $targetDirectory = $Output
                }
                else {
                    $targetDirectory = Join-Path $Output $relativeDirectory
                }
            }
            elseif ($inPlace) {
                $targetDirectory = $file.DirectoryName
            }
            else {
                $targetDirectory = $Output
            }

            $newName = "$baseName$extension"
            $destination = Join-Path $targetDirectory $newName

            # 只有目標重複時才套用 Duplicate 策略
            $isDuplicate = {
                param([string]$Candidate)

                if ($reserved.Contains($Candidate)) {
                    return $true
                }

                if (Test-Path -LiteralPath $Candidate) {
                    # 同一路徑視為自身，不算重複
                    if ($Candidate -eq $file.FullName) {
                        return $false
                    }

                    return $true
                }

                return $false
            }

            if (& $isDuplicate $destination) {
                switch ($Duplicate) {
                    'Numbering' {
                        $number = 1

                        do {
                            $newName = '{0}[{1}]{2}' -f `
                                $baseName, $number, $extension

                            $destination = Join-Path $targetDirectory $newName
                            $number++
                        } while (& $isDuplicate $destination)
                    }

                    'Origin' {
                        if ($DitchOrigin) {
                            $newName = "[$timestamp][$($file.BaseName)]$extension"
                            $destination = Join-Path $targetDirectory $newName
                        }
                    }

                    'Skip' {
                        $record.Status = 'Skip'
                        $record.Reason = '目標檔名已存在或與其他計畫重複'
                    }
                }

                if ($record.Status -ne 'Skip' -and
                    (& $isDuplicate $destination)) {
                    $record.Status = 'Skip'
                    $record.Reason = '套用重複策略後，目標檔名仍然重複'
                }
            }

            $record.NewName = $newName
            $record.Destination = $destination

            if ($record.Status -eq 'Pending') {
                if ($file.FullName -eq $destination) {
                    $record.Status = 'Unchanged'
                    $record.Reason = '檔名已符合規則'
                }
                else {
                    $record.Status = 'Ready'
                    [void]$reserved.Add($destination)
                }
            }
        }
        catch {
            $record.Status = 'Failed'
            $record.Reason = $_.Exception.Message
        }

        $plan.Add($record)
    }

    # 第二階段：展示全部原檔名 / 新檔名配對
    Write-Host ''
    Write-Host '========== 重新命名計畫 ==========' -ForegroundColor Cyan
    Write-Host "日期順序：$According；Duplicate：$Duplicate"
    Write-Host "Recursive：$Recursive；Immediate：$Immediate"
    Write-Host "模式：$(if ($inPlace) { 'In-place' } else { 'Output' })"
    Write-Host ''

    foreach ($record in $plan) {
        $color = $record.Color

        switch ($record.Status) {
            'Skip'      { $color = [ConsoleColor]::DarkYellow }
            'Failed'    { $color = [ConsoleColor]::Red }
            'Unchanged' { $color = [ConsoleColor]::DarkGray }
        }

        Write-Host "[$($record.Status)] [$($record.DateSource)]" `
            -ForegroundColor $color
        Write-Host "  原檔：$($record.SourceName)"
        Write-Host "  新檔：$(if ($record.NewName) { $record.NewName } else { '(無)' })"

        if ($record.Status -eq 'Skip' -or $record.Status -eq 'Failed') {
            Write-Host "  原因：$($record.Reason)" -ForegroundColor $color
        }
    }

    $ready = @($plan | Where-Object { $_.Status -eq 'Ready' })

    Write-Host ''
    Write-Host "計畫：$($plan.Count)；待執行：$($ready.Count)；略過/失敗：$(@($plan | Where-Object { $_.Status -in @('Skip', 'Failed') }).Count)" `
        -ForegroundColor Cyan

    if ($ready.Count -eq 0) {
        Write-Host '沒有可執行的檔案異動。' -ForegroundColor Yellow
        return
    }

    # Immediate=false：先詢問是否繼續
    if (-not $Immediate) {
        $answer = Read-Host '是否繼續執行上述檔名異動？[y/N]'

        if ($answer -notmatch '^(?i:y|yes)$') {
            Write-Host '已取消，沒有執行檔案異動。' -ForegroundColor Yellow
            return
        }
    }

    # 第三階段：執行已確認的計畫
    $success = 0
    $failed = 0
    $renameHistory = [System.Collections.Generic.List[object]]::new()

    foreach ($record in $ready) {
        try {
            $parent = Split-Path -Parent $record.Destination

            if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
                New-Item -ItemType Directory -Path $parent -Force |
                    Out-Null
            }

            # 執行前再次檢查，避免計畫建立後目標被其他程序占用
            if (Test-Path -LiteralPath $record.Destination) {
                throw "目標檔案已存在：$($record.Destination)"
            }

            if ($inPlace) {
                Rename-Item -LiteralPath $record.Source `
                    -NewName $record.NewName

                $renameHistory.Add([pscustomobject]@{
                    OriginalPath = $record.Source
                    NewPath      = $record.Destination
                })
            }
            else {
                Copy-Item -LiteralPath $record.Source `
                    -Destination $record.Destination
            }

            Write-Host "完成：$($record.SourceName) -> $($record.NewName)" `
                -ForegroundColor Green
            $success++
        }
        catch {
            Write-Host "失敗：$($record.SourceName) - $($_.Exception.Message)" `
                -ForegroundColor Red
            $failed++
        }
    }

    Write-Host ''
    Write-Host "執行完成：成功 $success；失敗 $failed" `
        -ForegroundColor Cyan

    # 只對成功的 inplace 異動提供 rollback
    if ($inPlace -and $renameHistory.Count -gt 0) {
        $answer = Read-Host '是否復原本次檔名異動？[Y/n]'

        if ([string]::IsNullOrWhiteSpace($answer) -or
            $answer -match '^(?i:y|yes)$') {

            $rollbackSuccess = 0
            $rollbackFailed = 0

            for ($i = $renameHistory.Count - 1; $i -ge 0; $i--) {
                $record = $renameHistory[$i]

                try {
                    if (-not (Test-Path -LiteralPath $record.NewPath -PathType Leaf)) {
                        throw "找不到新檔案：$($record.NewPath)"
                    }

                    if (Test-Path -LiteralPath $record.OriginalPath) {
                        throw "原始路徑已被占用：$($record.OriginalPath)"
                    }

                    Rename-Item -LiteralPath $record.NewPath `
                        -NewName ([System.IO.Path]::GetFileName(
                            $record.OriginalPath
                        ))

                    Write-Host "已復原：$($record.OriginalPath)" `
                        -ForegroundColor Green
                    $rollbackSuccess++
                }
                catch {
                    Write-Host "復原失敗：$($record.NewPath) - $($_.Exception.Message)" `
                        -ForegroundColor Red
                    $rollbackFailed++
                }
            }

            Write-Host "復原完成：成功 $rollbackSuccess；失敗 $rollbackFailed" `
                -ForegroundColor Cyan
        }
        else {
            Write-Host '保留本次檔名異動。' -ForegroundColor Yellow
        }
    }
}
