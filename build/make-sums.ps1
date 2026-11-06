param(
    [Parameter(Mandatory = $true)][string]$OutputDir
)

# 生成并复核 SHA256SUMS.txt
# 1) 对输出目录下的 .exe / .zip 计算 SHA256，写入 SHA256SUMS.txt（ASCII，两空格分隔，兼容 sha256sum）
# 2) 立刻重新计算一遍并与文件内容逐条比对，不一致则退出码 1
# 放在独立脚本里，避免批处理内联 PowerShell 时的引号/转义/编码问题。

$ErrorActionPreference = 'Stop'

if (-not (Test-Path $OutputDir)) {
    Write-Host "输出目录不存在: $OutputDir"
    exit 1
}

$files = Get-ChildItem -Path $OutputDir -File |
    Where-Object { $_.Extension -eq '.exe' -or $_.Extension -eq '.zip' } |
    Sort-Object Name

if (-not $files -or $files.Count -eq 0) {
    Write-Host "输出目录中没有产物: $OutputDir"
    exit 1
}

$sumsPath = Join-Path $OutputDir 'SHA256SUMS.txt'
$lines = foreach ($file in $files) {
    $hash = (Get-FileHash $file.FullName -Algorithm SHA256).Hash.ToLower()
    '{0}  {1}' -f $hash, $file.Name
}

$lines | Set-Content -Path $sumsPath -Encoding ASCII

# ---- 复核 ----
$bad = 0
foreach ($line in (Get-Content $sumsPath)) {
    $parts = $line -split '\s+', 2
    if ($parts.Count -lt 2) { continue }

    $expected = $parts[0].Trim()
    $name = $parts[1].Trim()
    $file = Join-Path $OutputDir $name

    if (-not (Test-Path $file)) {
        Write-Host "  缺失: $name"
        $bad++
        continue
    }

    $actual = (Get-FileHash $file -Algorithm SHA256).Hash.ToLower()
    if ($actual -ne $expected) {
        Write-Host "  校验不一致: $name"
        Write-Host "    文件值: $expected"
        Write-Host "    实际值: $actual"
        $bad++
    }
    else {
        Write-Host ("  {0}  {1}" -f $name, $expected)
    }
}

if ($bad -gt 0) {
    Write-Host "WARNING: SHA256SUMS.txt 与实际文件不一致"
    exit 1
}

Write-Host "  SHA256SUMS 复核通过"
exit 0
