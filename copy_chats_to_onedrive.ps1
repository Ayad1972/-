# نقل المحادثات إلى OneDrive (Windows)
# شغّله عبر: copy_chats_to_onedrive.bat

param(
    [string]$Dest = "",
    [switch]$DryRun
)

$ErrorActionPreference = "Continue"
try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $OutputEncoding = [System.Text.Encoding]::UTF8
} catch {}

function Write-Title($text) {
    Write-Host ""
    Write-Host ("=" * 60) -ForegroundColor Cyan
    Write-Host $text -ForegroundColor Cyan
    Write-Host ("=" * 60) -ForegroundColor Cyan
}
function Write-Ok($text) { Write-Host "[OK] $text" -ForegroundColor Green }
function Write-Info($text) { Write-Host "[..] $text" -ForegroundColor Yellow }
function Write-Warn($text) { Write-Host "[!!] $text" -ForegroundColor Magenta }

function Get-OneDriveFolder {
    foreach ($key in @("OneDrive", "OneDriveConsumer", "OneDriveCommercial")) {
        $value = [Environment]::GetEnvironmentVariable($key, "User")
        if (-not $value) { $value = [Environment]::GetEnvironmentVariable($key, "Process") }
        if ($value -and (Test-Path -LiteralPath $value)) { return $value }
    }
    $home = $env:USERPROFILE
    foreach ($name in @("OneDrive", "OneDrive - Personal")) {
        $candidate = Join-Path $home $name
        if (Test-Path -LiteralPath $candidate) { return $candidate }
    }
    $found = Get-ChildItem -LiteralPath $home -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like "OneDrive*" } |
        Select-Object -First 1
    if ($found) { return $found.FullName }
    try {
        $reg = Get-ItemProperty -Path "HKCU:\Software\Microsoft\OneDrive\Accounts\Personal" -ErrorAction SilentlyContinue
        if ($reg.UserFolder -and (Test-Path -LiteralPath $reg.UserFolder)) { return $reg.UserFolder }
    } catch {}
    return $null
}

function Get-FolderSizeBytes($path) {
    if (-not (Test-Path -LiteralPath $path)) { return 0 }
    try {
        $item = Get-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
        if ($item -and -not $item.PSIsContainer) { return [int64]$item.Length }
        $sum = (Get-ChildItem -LiteralPath $path -Recurse -Force -File -ErrorAction SilentlyContinue |
            Measure-Object -Property Length -Sum).Sum
        if ($sum) { return [int64]$sum } else { return [int64]0 }
    } catch { return [int64]0 }
}

function Format-Size([int64]$num) {
    if ($num -lt 1024) { return "$num بايت" }
    if ($num -lt 1048576) { return ("{0:N1} KB" -f ($num / 1024)) }
    if ($num -lt 1073741824) { return ("{0:N1} MB" -f ($num / 1048576)) }
    return ("{0:N1} GB" -f ($num / 1073741824))
}

function Copy-Safe($src, $dst) {
    if ($DryRun) { return }
    if (Test-Path -LiteralPath $src -PathType Leaf) {
        $parent = Split-Path -Parent $dst
        if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
        Copy-Item -LiteralPath $src -Destination $dst -Force
        return
    }
    if (-not (Test-Path -LiteralPath $dst)) { New-Item -ItemType Directory -Path $dst -Force | Out-Null }
    robocopy $src $dst /E /COPY:DAT /R:1 /W:1 /NFL /NDL /NJH /NJS /NP /XD Cache GPUCache "Code Cache" "Service Worker" Crashpad logs Log Blob_storage | Out-Null
}

Write-Title "نقل المحادثات إلى OneDrive"

$oneDrive = Get-OneDriveFolder
if (-not $Dest) {
    if (-not $oneDrive) {
        Write-Warn "لم يتم العثور على مجلد OneDrive."
        Write-Host "ثبّت تطبيق OneDrive ثم أعد التشغيل، أو اسحب المجلد إلى سطح المكتب."
        exit 2
    }
    $stamp = Get-Date -Format "yyyy-MM-dd_HHmm"
    $Dest = Join-Path (Join-Path $oneDrive "المحادثات") $stamp
}

Write-Host "مجلد OneDrive : $oneDrive"
Write-Host "مجلد الحفظ    : $Dest"
Write-Host "الأصل يبقى في مكانه. هذه نسخة محفوظة."
Write-Host ""

$home = $env:USERPROFILE
$roaming = $env:APPDATA
$local = $env:LOCALAPPDATA
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$candidates = @(
    @{ Title = "Cursor — قاعدة المحادثات العامة"; Path = Join-Path $roaming "Cursor\User\globalStorage"; CursorGlobal = $true },
    @{ Title = "Cursor — محادثات المشاريع"; Path = Join-Path $roaming "Cursor\User\workspaceStorage"; CursorWs = $true },
    @{ Title = "واتساب (سطح المكتب)"; Path = Join-Path $roaming "WhatsApp" },
    @{ Title = "واتساب (Local)"; Path = Join-Path $local "WhatsApp" },
    @{ Title = "تليغرام"; Path = Join-Path $roaming "Telegram Desktop\tdata" },
    @{ Title = "تيمز (القديم)"; Path = Join-Path $roaming "Microsoft\Teams" },
    @{ Title = "أرشيف محادثات هذا المشروع"; Path = Join-Path $scriptDir "conversations" }
)

foreach ($name in @("chats", "ai-tracking", "conversations")) {
    $p = Join-Path (Join-Path $home ".cursor") $name
    if (Test-Path -LiteralPath $p) {
        $candidates += @{ Title = "Cursor — $name"; Path = $p }
    }
}

$found = @()
foreach ($item in $candidates) {
    if ($item.Path -and (Test-Path -LiteralPath $item.Path)) {
        $item.Size = Get-FolderSizeBytes $item.Path
        $found += $item
        Write-Host ("  - {0}  ({1})" -f $item.Title, (Format-Size $item.Size))
    }
}

$storeWa = Get-ChildItem -LiteralPath (Join-Path $local "Packages") -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -like "5319275A.WhatsAppDesktop_*" }
foreach ($pkg in $storeWa) {
    $p = Join-Path $pkg.FullName "LocalState"
    if (Test-Path -LiteralPath $p) {
        $found += @{ Title = "واتساب من المتجر"; Path = $p; Size = (Get-FolderSizeBytes $p) }
        Write-Host ("  - واتساب من المتجر  ({0})" -f (Format-Size (Get-FolderSizeBytes $p)))
    }
}

if ($found.Count -eq 0) {
    Write-Warn "لم يتم العثور على مصادر محادثات."
    exit 3
}

Write-Host ""
if (-not $DryRun) {
    $answer = Read-Host "Enter للنسخ الآن، أو اكتب لا للإلغاء"
    if ($answer -eq "لا" -or $answer -eq "n" -or $answer -eq "N") {
        Write-Host "تم الإلغاء."
        exit 0
    }
    New-Item -ItemType Directory -Path $Dest -Force | Out-Null
}

foreach ($item in $found) {
    $target = Join-Path $Dest $item.Title
    Write-Info ("نسخ: {0}" -f $item.Title)
    if ($item.CursorGlobal) {
        if (-not $DryRun) { New-Item -ItemType Directory -Path $target -Force | Out-Null }
        foreach ($name in @("state.vscdb", "state.vscdb.backup")) {
            $srcFile = Join-Path $item.Path $name
            if (Test-Path -LiteralPath $srcFile) {
                Copy-Safe $srcFile (Join-Path $target $name)
            }
        }
    } elseif ($item.CursorWs) {
        Get-ChildItem -LiteralPath $item.Path -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            $wsDest = Join-Path $target $_.Name
            foreach ($name in @("state.vscdb", "state.vscdb.backup", "workspace.json")) {
                $srcFile = Join-Path $_.FullName $name
                if (Test-Path -LiteralPath $srcFile) {
                    Copy-Safe $srcFile (Join-Path $wsDest $name)
                }
            }
        }
    } else {
        Copy-Safe $item.Path $target
    }
    Write-Ok $item.Title
}

$report = Join-Path $Dest "تقرير_النقل.txt"
$lines = @(
    "تقرير نقل المحادثات إلى OneDrive",
    ("=" * 50),
    ("التاريخ: {0}" -f (Get-Date -Format "yyyy-MM-dd HH:mm")),
    ("مجلد OneDrive: {0}" -f $oneDrive),
    ("مجلد الحفظ: {0}" -f $Dest),
    "",
    "ملاحظة: تم النسخ دون حذف الأصل حتى تبقى البرامج تعمل.",
    ""
)
foreach ($item in $found) {
    $lines += ("- {0}" -f $item.Title)
}
$lines += ""
$lines += "بعد المزامنة ستظهر الملفات على https://onedrive.live.com"
if (-not $DryRun) {
    $utf8 = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllLines($report, $lines, $utf8)
    Write-Ok "التقرير: $report"
    Start-Process explorer.exe $Dest
}

Write-Host ""
Write-Host "انتهى. إذا كان OneDrive يعمل ستبدأ المزامنة تلقائياً."
exit 0
