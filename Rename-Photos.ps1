param (
    [Parameter(Mandatory=$false)]
    [string]$TargetDir
)

# 確保載入必要的 .NET 元件
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# 若未提供參數，自動跳出資料夾選擇視窗
if ([string]::IsNullOrEmpty($TargetDir)) {
    Write-Host "未偵測到輸入路徑，正在開啟資料夾選取視窗..." -ForegroundColor Cyan
    $FolderBrowser = New-Object -ComObject Shell.Application
    $Folder = $FolderBrowser.BrowseForFolder(0, "請選擇影像所在的資料夾", 0, 0)
    
    if ($Folder) {
        $TargetDir = $Folder.Self.Path
        Write-Host "已選取路徑: $TargetDir" -ForegroundColor Green
    } else {
        Write-Host "已取消選取，腳本結束。" -ForegroundColor Red
        exit
    }
}

if (-not (Test-Path -Path $TargetDir -PathType Container)) {
    Write-Error "找不到指定的資料夾路徑：$TargetDir"
    exit
}

# 支援的影像副檔名
$Extensions = @(".jpg", ".jpeg", ".png", ".heic")
$Files = Get-ChildItem -Path $TargetDir -File

foreach ($File in $Files) {
    if ($Extensions -contains $File.Extension.ToLower()) {
        $TakenDate = $null
        $Image = $null
        
        try {
            $Stream = New-Object System.IO.FileStream($File.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
            $Image = [System.Drawing.Image]::FromStream($Stream)
            
            # 定義 Exif 時間欄位的優先順序：36867 (Original) -> 36868 (Digitized) -> 306 (DateTime)
            $TimeTags = @(36867, 36868, 306)
            $TargetTag = $null
            
            foreach ($Tag in $TimeTags) {
                if ($Image.PropertyIdList -contains $Tag) {
                    $TargetTag = $Tag
                    break
                }
            }
            
            # 如果有找到任一時間標記，開始解析
            if ($TargetTag) {
                $PropertyItem = $Image.GetPropertyItem($TargetTag)
                $DateStringRaw = [System.Text.Encoding]::ASCII.GetString($PropertyItem.Value).Trim()
                
                # 擷取字串中的 "YYYY:MM:DD HH:MM:SS" 部分
                if ($DateStringRaw -match '(\d{4}):(\d{2}):(\d{2})\s(\d{2}):(\d{2}):(\d{2})') {
                    $CleanDateStr = $Matches[0]
                    $TakenDate = [DateTime]::ParseExact($CleanDateStr, "yyyy:MM:dd HH:mm:ss", $null)
                }
            }
        }
        catch {
            Write-Host "讀取 [$($File.Name)] Exif 時發生錯誤: $_" -ForegroundColor Red
        }
        finally {
            if ($Image) { $Image.Dispose() }
            if ($Stream) { $Stream.Close(); $Stream.Dispose() }
        }
        
        # 最終後援方案：若無任何內嵌 Exif 時間標記，才使用檔案修改時間
        if (-not $TakenDate) {
            $TakenDate = $File.LastWriteTime
            Write-Host "檔案 [$($File.Name)] 無法解析內嵌 Exif 拍攝日期，改用檔案修改時間。" -ForegroundColor Yellow
        }
        
        # 依據格式 [YYYY.MM.DDThh.mm.ss][F].E 格式化新檔名 (HH 為 24 小時制)
        $DateString = $TakenDate.ToString("yyyy.MM.ddTHH.mm.ss")
        $BaseName = [System.IO.Path]::GetFileNameWithoutExtension($File.Name)
        
        # 避免重複命名已經有時間前綴的檔案
        if ($BaseName -match '^\[\d{4}\.\d{2}\.\d{2}T\d{2}\.\d{2}\.\d{2}\]') {
            Write-Host "檔案 [$($File.Name)] 已經有時間前綴，跳過。" -ForegroundColor Gray
            continue
        }
        
        $NewName = "[$DateString][$BaseName]$($File.Extension)"
        $NewPath = Join-Path -Path $File.DirectoryName -ChildPath $NewName
        
        if ($File.FullName -ne $NewPath -and (Test-Path $NewPath) -eq $false) {
            Write-Host "重新命名: $($File.Name) -> $NewName" -ForegroundColor Green
            Rename-Item -Path $File.FullName -NewName $NewName
        }
    }
}

Write-Host "影像重新命名作業完成！" -ForegroundColor Cyan