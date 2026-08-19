# مقارنة mm.xlsx مع ss.xls وإضافة الأرقام من mm إلى ss حسب الاسم المشترك
# الاستخدام:
#   powershell -ExecutionPolicy Bypass -File merge_mm_into_ss.ps1
#   powershell -ExecutionPolicy Bypass -File merge_mm_into_ss.ps1 -SsPath "ss.xls" -MmPath "mm.xlsx"

param(
    [string]$SsPath = $(Join-Path $PSScriptRoot "ss.xls"),
    [string]$MmPath = $(Join-Path $PSScriptRoot "mm.xlsx"),
    [string]$OutPath = $(Join-Path $PSScriptRoot "ss_updated.xls")
)

$ErrorActionPreference = "Stop"

function Normalize-Name([object]$value) {
    if ($null -eq $value) { return "" }
    $text = "$value".Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return "" }

    $text = $text.Replace("أ", "ا").Replace("إ", "ا").Replace("آ", "ا")
    $text = $text.Replace("ة", "ه").Replace("ى", "ي")
    $text = $text -replace "\s+", " "
    return $text.ToLowerInvariant()
}

function Get-CellText($sheet, $row, $col) {
    $val = $sheet.Cells.Item($row, $col).Text
    if ($null -eq $val) { return "" }
    return "$val".Trim()
}

function Find-ColumnIndex($sheet, [string[]]$aliases, [int]$headerRow = 1, [int]$maxCols = 50) {
    $usedCols = [Math]::Min([int]$sheet.UsedRange.Columns.Count, $maxCols)
    $headers = @()
    for ($c = 1; $c -le $usedCols; $c++) {
        $headers += (Get-CellText $sheet $headerRow $c)
    }

    for ($c = 0; $c -lt $headers.Count; $c++) {
        $hNorm = Normalize-Name $headers[$c]
        if ([string]::IsNullOrWhiteSpace($hNorm)) { continue }
        foreach ($alias in $aliases) {
            $aNorm = Normalize-Name $alias
            if ($hNorm -eq $aNorm -or $hNorm.Contains($aNorm) -or $aNorm.Contains($hNorm)) {
                return @{ Index = ($c + 1); Header = $headers[$c] }
            }
        }
    }
    return $null
}

function Find-BestNameAndNumberColumns($sheet) {
    $nameAliases = @(
        "الاسم", "اسم", "اسم الموظف", "اسم_الموظف", "الموظف",
        "NAME", "EMP_NAME", "EMPLOYEE", "FULLNAME", "FULL_NAME"
    )
    $numberAliases = @(
        "الرقم", "رقم", "رقم الموظف", "رقم_الموظف", "رقم وظيفي",
        "NUMBER", "NO", "NUM", "EMPNO", "EMP_NO", "ID", "CODE", "الكود", "كود"
    )

    $nameCol = Find-ColumnIndex $sheet $nameAliases
    $numCol = Find-ColumnIndex $sheet $numberAliases

    # إذا لم نجد عناوين واضحة: نفترض العمود 1 اسم والعمود 2 رقم
    if ($null -eq $nameCol -or $null -eq $numCol) {
        $c1 = Get-CellText $sheet 1 1
        $c2 = Get-CellText $sheet 1 2
        $looksHeader = ($c1 -match "اسم|name|موظف") -or ($c2 -match "رقم|number|no|id|كود")
        return @{
            NameIndex   = 1
            NumberIndex = 2
            NameHeader  = $(if ($looksHeader) { $c1 } else { "COL1" })
            NumberHeader= $(if ($looksHeader) { $c2 } else { "COL2" })
            DataStartRow= $(if ($looksHeader) { 2 } else { 1 })
            Assumed     = $true
        }
    }

    return @{
        NameIndex   = [int]$nameCol.Index
        NumberIndex = [int]$numCol.Index
        NameHeader  = [string]$nameCol.Header
        NumberHeader= [string]$numCol.Header
        DataStartRow= 2
        Assumed     = $false
    }
}

Write-Host "============================================"
Write-Host "دمج أرقام mm في ss حسب الاسم المشترك"
Write-Host "============================================"

if (-not (Test-Path -LiteralPath $SsPath)) {
    Write-Host "File not found: $SsPath"
    Write-Host "Put ss.xls in this folder or in data"
    exit 0
}
if (-not (Test-Path -LiteralPath $MmPath)) {
    Write-Host "File not found: $MmPath"
    Write-Host "Put mm.xlsx in this folder or in data"
    exit 0
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $ssBook = $excel.Workbooks.Open($SsPath)
    $mmBook = $excel.Workbooks.Open($MmPath)

    $ssSheet = $ssBook.Worksheets.Item(1)
    $mmSheet = $mmBook.Worksheets.Item(1)

    $ssCols = Find-BestNameAndNumberColumns $ssSheet
    $mmCols = Find-BestNameAndNumberColumns $mmSheet

    Write-Host ""
    Write-Host "ss.xls :"
    Write-Host ("  عمود الاسم = {0} | عمود الرقم = {1}" -f $ssCols.NameHeader, $ssCols.NumberHeader)
    Write-Host "mm.xlsx :"
    Write-Host ("  عمود الاسم = {0} | عمود الرقم = {1}" -f $mmCols.NameHeader, $mmCols.NumberHeader)

    # بناء قاموس الاسم -> الرقم من mm
    $mmMap = @{}
    $mmLastRow = [int]$mmSheet.UsedRange.Rows.Count
    for ($r = [int]$mmCols.DataStartRow; $r -le $mmLastRow; $r++) {
        $name = Get-CellText $mmSheet $r ([int]$mmCols.NameIndex)
        $num  = Get-CellText $mmSheet $r ([int]$mmCols.NumberIndex)
        $key  = Normalize-Name $name
        if ([string]::IsNullOrWhiteSpace($key)) { continue }
        if ([string]::IsNullOrWhiteSpace($num)) { continue }
        if (-not $mmMap.ContainsKey($key)) {
            $mmMap[$key] = $num
        }
    }

    Write-Host ("عدد الأسماء في mm: {0}" -f $mmMap.Count)

    $updated = 0
    $matchedSame = 0
    $notFound = 0
    $ssLastRow = [int]$ssSheet.UsedRange.Rows.Count

    for ($r = [int]$ssCols.DataStartRow; $r -le $ssLastRow; $r++) {
        $name = Get-CellText $ssSheet $r ([int]$ssCols.NameIndex)
        $key  = Normalize-Name $name
        if ([string]::IsNullOrWhiteSpace($key)) { continue }

        if ($mmMap.ContainsKey($key)) {
            $newNum = $mmMap[$key]
            $oldNum = Get-CellText $ssSheet $r ([int]$ssCols.NumberIndex)
            if ($oldNum -ne $newNum) {
                $ssSheet.Cells.Item($r, [int]$ssCols.NumberIndex).Value2 = $newNum
                $updated++
                Write-Host ("تحديث: {0} | من [{1}] إلى [{2}]" -f $name, $oldNum, $newNum)
            } else {
                $matchedSame++
            }
        } else {
            $notFound++
        }
    }

    # حفظ النتيجة
    if (Test-Path -LiteralPath $OutPath) {
        Remove-Item -LiteralPath $OutPath -Force
    }

    # 56 = xlExcel8 (.xls)
    $ssBook.SaveAs($OutPath, 56)
    $ssBook.Close($false)
    $mmBook.Close($false)

    Write-Host "--------------------------------------------"
    Write-Host ("تم التحديث          : {0}" -f $updated)
    Write-Host ("متطابق بدون تغيير   : {0}" -f $matchedSame)
    Write-Host ("اسم غير موجود في mm : {0}" -f $notFound)
    Write-Host ("الملف الناتج        : {0}" -f $OutPath)
    Write-Host "تم إنشاء ملف جديد ولم يتم تعديل ss.xls الأصلي."
}
finally {
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
}
