# -*- coding: utf-8 -*-
# تحسين سرعة الحاسوب - خطوات آمنة لويندوز
# شغّله عبر: speed_up_pc.bat

param(
    [string]$Step = ""
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
function Pause-Step {
    Write-Host ""
    Read-Host "اضغط Enter للمتابعة" | Out-Null
}

function Get-FolderSizeMB($path) {
    if (-not (Test-Path $path)) { return 0 }
    try {
        $bytes = (Get-ChildItem -LiteralPath $path -Recurse -Force -ErrorAction SilentlyContinue |
            Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
        if (-not $bytes) { return 0 }
        return [math]::Round($bytes / 1MB, 1)
    } catch {
        return 0
    }
}

function Step-Diagnose {
    Write-Title "الخطوة 1: فحص حالة الجهاز"
    $os = Get-CimInstance Win32_OperatingSystem
    $cs = Get-CimInstance Win32_ComputerSystem
    $cpu = Get-CimInstance Win32_Processor | Select-Object -First 1
    $totalRam = [math]::Round($cs.TotalPhysicalMemory / 1GB, 1)
    $freeRam = [math]::Round($os.FreePhysicalMemory / 1MB, 1)
    $usedRamPct = 0
    if (($totalRam * 1024) -gt 0) {
        $usedRamPct = [math]::Round((($totalRam * 1024) - $freeRam) / ($totalRam * 1024) * 100, 0)
    }

    Write-Host "الجهاز        : $($cs.Model)"
    Write-Host "المعالج       : $($cpu.Name)"
    Write-Host "الذاكرة الكلية: $totalRam GB"
    Write-Host "الذاكرة الحرة : $freeRam MB"
    Write-Host "استخدام الرام : $usedRamPct%"
    Write-Host "ويندوز        : $($os.Caption) $($os.Version)"

    Write-Host ""
    Write-Host "مساحة الأقراص:"
    Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" | ForEach-Object {
        $freeGB = [math]::Round($_.FreeSpace / 1GB, 1)
        $sizeGB = [math]::Round($_.Size / 1GB, 1)
        $pct = if ($_.Size -gt 0) { [math]::Round(($_.FreeSpace / $_.Size) * 100, 0) } else { 0 }
        $mark = if ($pct -lt 15) { " << ممتلئ تقريباً" } else { "" }
        Write-Host ("  {0}  حر {1} GB من {2} GB ({3}% متاح){4}" -f $_.DeviceID, $freeGB, $sizeGB, $pct, $mark)
    }

    Write-Host ""
    Write-Host "أكثر البرامج استهلاكاً للذاكرة الآن:"
    Get-Process | Sort-Object WorkingSet64 -Descending | Select-Object -First 8 |
        ForEach-Object {
            $mb = [math]::Round($_.WorkingSet64 / 1MB, 0)
            Write-Host ("  {0,-28} {1,6} MB" -f $_.ProcessName, $mb)
        }

    if ($usedRamPct -ge 85) {
        Write-Warn "الذاكرة ممتلئة جداً. أغلق البرامج غير المستخدمة قبل العمل على الفوكس/الأفراد."
    } elseif ($usedRamPct -ge 70) {
        Write-Info "الذاكرة مرتفعة. يفضّل تقليل البرامج المفتوحة مع نظام الأفراد."
    } else {
        Write-Ok "الذاكرة في وضع مقبول حالياً."
    }
    Pause-Step
}

function Step-TempCleanup {
    Write-Title "الخطوة 2: تنظيف الملفات المؤقتة"
    $targets = @(
        $env:TEMP,
        "$env:LOCALAPPDATA\Temp",
        "$env:WINDIR\Temp",
        "$env:LOCALAPPDATA\Microsoft\Windows\INetCache"
    )

    $before = 0
    foreach ($t in $targets) {
        if (Test-Path $t) {
            $sz = Get-FolderSizeMB $t
            $before += $sz
            Write-Host ("  قبل: {0,-55} {1,8} MB" -f $t, $sz)
        }
    }

    Write-Info "جاري الحذف الآمن للملفات المؤقتة (لن تُمس ملفاتك الشخصية)..."
    $deleted = 0
    foreach ($t in $targets) {
        if (-not (Test-Path $t)) { continue }
        Get-ChildItem -LiteralPath $t -Force -ErrorAction SilentlyContinue | ForEach-Object {
            try {
                Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction Stop
                $deleted++
            } catch {
                # ملفات قيد الاستخدام تُتخطى
            }
        }
    }

    if (Get-Command cleanmgr.exe -ErrorAction SilentlyContinue) {
        Write-Info "تشغيل تنظيف القرص التلقائي إن كان مُعداً مسبقاً..."
        try {
            Start-Process -FilePath "cleanmgr.exe" -ArgumentList "/sagerun:1" -WindowStyle Minimized -Wait -ErrorAction SilentlyContinue
        } catch {}
    }

    Write-Ok "تمت محاولة تنظيف $deleted عنصر تقريباً (الحجم السابق المؤقت ≈ $before MB)."
    Write-Host "أعد تشغيل الجهاز لاحقاً لتحرير بقايا الملفات المقفلة."
    Pause-Step
}

function Step-RecycleBin {
    Write-Title "الخطوة 3: تفريغ سلة المحذوفات"
    try {
        Clear-RecycleBin -Force -ErrorAction Stop
        Write-Ok "تم تفريغ سلة المحذوفات."
    } catch {
        try {
            $shell = New-Object -ComObject Shell.Application
            $bin = $shell.NameSpace(0xA)
            if ($bin -ne $null) {
                $bin.Items() | ForEach-Object {
                    Remove-Item -LiteralPath $_.Path -Recurse -Force -ErrorAction SilentlyContinue
                }
            }
            Write-Ok "تم تفريغ سلة المحذوفات (طريقة بديلة)."
        } catch {
            Write-Warn "تعذر تفريغ السلة تلقائياً. فرّغها يدوياً من سطح المكتب."
        }
    }
    Pause-Step
}

function Step-StartupApps {
    Write-Title "الخطوة 4: برامج التشغيل مع ويندوز"
    Write-Host "هذه البرامج تبدأ مع الجهاز وتبطئه. عطّل غير الضروري فقط."
    Write-Host "لا تعطّل برامج الحماية أو تعريفات الصوت/الشبكة إن لم تعرفها."
    Write-Host ""

    $startup = @()
    try {
        $startup = @(Get-CimInstance Win32_StartupCommand | Sort-Object Name)
    } catch {}

    if ($startup.Count -eq 0) {
        Write-Info "تعذر قراءة قائمة التشغيل. سيتم فتح مدير المهام."
    } else {
        $i = 1
        foreach ($s in $startup) {
            Write-Host ("  {0,2}. {1,-30} | {2}" -f $i, $s.Name, $s.Location)
            $i++
        }
    }

    Write-Host ""
    Write-Host "سيتم فتح:"
    Write-Host "  1) مدير المهام -> Startup apps"
    Write-Host "  2) إعدادات تطبيقات بدء التشغيل"
    $open = Read-Host "افتحها الآن؟ (Y/N)"
    if ($open -match '^(Y|y|نعم|ن)$') {
        try { Start-Process "taskmgr.exe" } catch {}
        try { Start-Process "ms-settings:startupapps" } catch {}
        Write-Info "في Startup: عطّل المتصفح الزائد، Teams، Updater، OneDrive إن لم تحتاجه."
        Write-Info "اترك مضاد الفيروسات وخدمات الطابعة/الشبكة إن كنت تستخدمها."
    }
    Pause-Step
}

function Step-PowerPlan {
    Write-Title "الخطوة 5: خطة الطاقة للأداء الأفضل"
    Write-Host "الخطة الحالية:"
    powercfg /GETACTIVESCHEME
    Write-Host ""
    $ans = Read-Host "تفعيل خطة High performance؟ (Y/N)"
    if ($ans -match '^(Y|y|نعم|ن)$') {
        $guid = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"
        powercfg /SETACTIVE $guid 2>$null
        if ($LASTEXITCODE -ne 0) {
            powercfg -duplicatescheme $guid 2>$null | Out-Null
            powercfg /SETACTIVE $guid 2>$null
        }
        Write-Ok "تم طلب تفعيل High performance."
        powercfg /GETACTIVESCHEME
    } else {
        Write-Info "تم التخطي."
    }
    Pause-Step
}

function Step-DiskHealth {
    Write-Title "الخطوة 6: فحص القرص السريع"
    $sysDrive = $env:SystemDrive
    Write-Info "فحص سطحي للقرص $sysDrive (بدون إصلاح تلقائي طويل)..."
    cmd /c "chkdsk $sysDrive" 2>&1 | Select-Object -Last 12

    $disk = Get-PhysicalDisk -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($disk) {
        Write-Host ""
        Write-Host ("نوع القرص: {0} | الحالة: {1}" -f $disk.MediaType, $disk.HealthStatus)
        if ($disk.MediaType -eq "HDD") {
            Write-Warn "قرص HDD ميكانيكي: تجنّب فتح ملفات كبيرة من الشبكة، وضع DBF محلياً يسرّع الفوكس."
            $defrag = Read-Host "تشغيل إلغاء التجزئة الآن؟ (قد يستغرق وقتاً) (Y/N)"
            if ($defrag -match '^(Y|y|نعم|ن)$') {
                defrag $sysDrive /O
            }
        } elseif ($disk.MediaType -eq "SSD") {
            Write-Ok "SSD: لا تحتاج إلغاء تجزئة تقليدي. تأكد أن TRIM مفعّل."
            fsutil behavior query DisableDeleteNotify
        }
    }
    Pause-Step
}

function Step-FoxTips {
    Write-Title "الخطوة 7: تسريع نظام الأفراد / الفوكس / DBF"
    Write-Host @"
نصائح عملية لعمل أسرع مع ملفاتك:

1) ضع MFILE.DBF وملفات Excel على قرص محلي سريع (C: أو SSD)
   وليس على شبكة بطيئة أثناء التعديل الثقيل.

2) أغلق Excel تماماً قبل تشغيل سكربتات النقل إلى DBF
   حتى لا يبقى الملف مقفولاً أو بطيئاً.

3) لا تشغّل فحص فيروسات كاملاً أثناء إدخال البيانات في الفوكس.

4) أبقِ نافذة واحدة فقط من نظام الأفراد مفتوحة إن أمكن.

5) بعد النقل الكبير: أعد فهرسة/إعادة بناء الفهارس من داخل الفوكس
   إن كان النظام يدعم ذلك.

6) استخدم:
     run_update_mfile.bat
   بدل النسخ اليدوي المتكرر.

7) احذف نسخ .bak_* القديمة لـ MFILE إن امتلأ القرص
   بعد التأكد أن النسخة الحالية سليمة.
"@
    Pause-Step
}

function Step-ServicesHint {
    Write-Title "الخطوة 8: خدمات وخلفية ويندوز (يدوي وآمن)"
    Write-Host @"
لا نوقف خدمات النظام تلقائياً حتى لا يتعطل الجهاز.

افحص يدوياً إن أردت:
  - Windows Search indexing: إن كان الجهاز ضعيفاً يمكن تقليل الفهرسة
  - OneDrive / Google Drive / Dropbox: أوقف المزامنة أثناء العمل على DBF
  - تحديثات ويندوز: أجّلها لوقت خارج الدوام إن كانت تبطئ الجهاز الآن

فتح أدوات مفيدة:
"@
    $ans = Read-Host "فتح Resource Monitor و Disk Cleanup؟ (Y/N)"
    if ($ans -match '^(Y|y|نعم|ن)$') {
        try { Start-Process "resmon.exe" } catch {}
        try { Start-Process "cleanmgr.exe" } catch {}
    }
    Pause-Step
}

function Step-AllSafe {
    Step-Diagnose
    Step-TempCleanup
    Step-RecycleBin
    Step-StartupApps
    Step-PowerPlan
    Step-DiskHealth
    Step-FoxTips
    Step-ServicesHint
    Write-Title "اكتملت الخطوات"
    Write-Ok "يُفضّل إعادة تشغيل الجهاز الآن لتثبيت التحسينات."
}

function Show-Menu {
    Clear-Host
    Write-Title "تسريع الحاسوب خطوة بخطوة"
    Write-Host "1) فحص حالة الجهاز"
    Write-Host "2) تنظيف الملفات المؤقتة"
    Write-Host "3) تفريغ سلة المحذوفات"
    Write-Host "4) مراجعة برامج التشغيل مع ويندوز"
    Write-Host "5) تفعيل خطة الطاقة عالية الأداء"
    Write-Host "6) فحص القرص / SSD-HDD"
    Write-Host "7) نصائح تسريع الفوكس والأفراد"
    Write-Host "8) أدوات المراقبة والتنظيف اليدوي"
    Write-Host "9) تنفيذ كل الخطوات الآمنة بالترتيب"
    Write-Host "0) خروج"
    Write-Host ""
}

if ($Step) {
    switch ($Step.ToLower()) {
        "1" { Step-Diagnose }
        "2" { Step-TempCleanup }
        "3" { Step-RecycleBin }
        "4" { Step-StartupApps }
        "5" { Step-PowerPlan }
        "6" { Step-DiskHealth }
        "7" { Step-FoxTips }
        "8" { Step-ServicesHint }
        "all" { Step-AllSafe }
        default { Write-Warn "خطوة غير معروفة: $Step" }
    }
    exit 0
}

do {
    Show-Menu
    $choice = Read-Host "اختر رقم الخطوة"
    switch ($choice) {
        "1" { Step-Diagnose }
        "2" { Step-TempCleanup }
        "3" { Step-RecycleBin }
        "4" { Step-StartupApps }
        "5" { Step-PowerPlan }
        "6" { Step-DiskHealth }
        "7" { Step-FoxTips }
        "8" { Step-ServicesHint }
        "9" { Step-AllSafe }
        "0" { break }
        default { Write-Warn "اختيار غير صالح" ; Start-Sleep -Seconds 1 }
    }
} while ($true)

Write-Host "إلى اللقاء."
