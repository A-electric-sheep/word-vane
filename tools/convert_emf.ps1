# EMF/WMF -> PNG 高分光栅化脚本(可选, 独立于主转换器)
# 说明: 主转换器 word2md.py 已内置 Pillow 三倍放大光栅化路径。
# 本脚本提供基于 .NET System.Drawing 的替代方案, 输出分辨率更可控。
#
# [DESENSITIZED] 原脚本硬编码了内部项目的绝对目录, 现将输入/输出目录参数化。
#
# 用法: 依次传入 <输入目录> <输出目录>
#   powershell -ExecutionPolicy Bypass -File tools/convert_emf.ps1 <input_dir> <output_dir>

param(
    [Parameter(Mandatory = $true)]
    [string]$InDir,
    [Parameter(Mandatory = $true)]
    [string]$OutDir
)

Add-Type -AssemblyName System.Drawing
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }
if (-not (Test-Path $InDir)) { throw "输入目录不存在: $InDir" }

$files = Get-ChildItem -Path $InDir -Include *.emf,*.wmf -File
$ok = 0; $fail = 0; $failedList = @()
foreach ($f in $files) {
    try {
        $img = [System.Drawing.Image]::FromFile($f.FullName)
        # 物理尺寸(0.01mm) -> 英寸 -> 按 scale 放大
        $scale = 3.0
        $w = [int]([math]::Ceiling($img.Width  / $img.HorizontalResolution * 96 * $scale))
        $h = [int]([math]::Ceiling($img.Height / $img.VerticalResolution   * 96 * $scale))
        $bmp = New-Object System.Drawing.Bitmap($w, $h)
        $bmp.SetResolution(96, 96)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.Clear([System.Drawing.Color]::White)
        $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $g.DrawImage($img, 0, 0, $w, $h)
        $g.Dispose(); $img.Dispose()
        $outName = [System.IO.Path]::GetFileNameWithoutExtension($f.Name) + '.png'
        $bmp.Save((Join-Path $OutDir $outName), [System.Drawing.Imaging.ImageFormat]::Png)
        $bmp.Dispose()
        $ok++
    } catch {
        $fail++; $failedList += ($f.Name + ' : ' + $_.Exception.Message)
    }
}
Write-Output ("OK=$ok FAIL=$fail")
$failedList | ForEach-Object { Write-Output "FAILED: $_" }