$ErrorActionPreference = "Stop"
$srcDir = "C:\Users\maqi\Desktop\方案"
$dstDir = "c:\Users\maqi\Desktop\trae\word备注解答工具设计\pdfs"
if (-not (Test-Path $dstDir)) { New-Item -ItemType Directory -Path $dstDir | Out-Null }

try {
    $ppt = New-Object -ComObject PowerPoint.Application
    Write-Host "PPT_OK"
} catch {
    Write-Host "PPT_FAIL: $_"
    exit 1
}

# 优化文件名映射，生成英文文件名避免编码问题
$mapping = @{
    "12.12市场秩序-水电气收费监管解决方案V5【待评审】.pptx" = "01_water_gas_electricity.pdf"
    "【新】智慧行政审批标准化解决方案.pptx" = "02_smart_admin_approval.pdf"
    "【新】非现场监管-食品全产业链溯源监管解决方案1.pptx" = "03_food_traceability.pdf"
    "【新】非现场监管-食品安全互联网+AI监管解决方案(地市版).pptx" = "04_food_ai_safety.pdf"
    "【新】非现场监管——成品油综合监管解决方案（v2）.pptx" = "05_refined_oil.pdf"
    "市场监管行政执法解决方案-0315 V6.pptx" = "06_administrative_enforcement.pdf"
    "海南市场监管局方案_海南省工业产品质量安全监管信息系统建设方案 (备注).pptx" = "07_industrial_product_quality.pdf"
    "网络直播监测解决方案7.23.pdf" = "08_livestream_monitoring.pdf"
    "药品监管-智慧药检解决方案V1.pptx" = "09_drug_inspection.pdf"
    "质量监管"一站式"服务平台解决方案(培训)-0702-中移集成解决方案部市场监管行业.pdf" = "10_nqi_one_stop.pdf"
}

Get-ChildItem -Path $srcDir -File | ForEach-Object {
    $srcFile = $_.FullName
    $srcName = $_.Name
    $outName = $mapping[$srcName]
    if (-not $outName) {
        # 兜底：用序号
        $ext = if ($srcName -match "\.pptx$") { ".pdf" } else { ".pdf" }
        $outName = "99_" + [System.IO.Path]::GetFileNameWithoutExtension($srcName) + $ext
    }
    $outPath = Join-Path $dstDir $outName

    if (Test-Path $outPath) {
        Write-Host "SKIP: $outName (already exists)"
        return
    }

    Write-Host "Converting: $srcName -> $outName"

    if ($srcName -match "\.pptx$") {
        try {
            $pres = $ppt.Presentations.Open($srcFile, $true, $false, $false)
            # 32 = ppFixedFormatTypePDF
            $pres.SaveAs($outPath, 32)
            $pres.Close()
            Write-Host "  DONE"
        } catch {
            Write-Host "  ERROR: $_"
        }
    } elseif ($srcName -match "\.pdf$") {
        try {
            Copy-Item -Path $srcFile -Destination $outPath -Force
            Write-Host "  COPIED (already PDF)"
        } catch {
            Write-Host "  COPY ERROR: $_"
        }
    }
}

$ppt.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
Write-Host "ALL_DONE"
