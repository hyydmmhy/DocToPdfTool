param(
    [Parameter(Mandatory = $true)][string]$SourceDir,
    [Parameter(Mandatory = $true)][string]$DestinationPath
)

# 把目录打包成 zip（等价于 Compress-Archive，但错误处理更明确）。
# 单独成脚本，避免批处理内联 PowerShell 的引号问题。

$ErrorActionPreference = 'Stop'

if (-not (Test-Path $SourceDir)) {
    Write-Host "源目录不存在: $SourceDir"
    exit 1
}

$parent = Split-Path -Parent $DestinationPath
if ($parent -and -not (Test-Path $parent)) {
    New-Item -ItemType Directory -Path $parent -Force | Out-Null
}

if (Test-Path $DestinationPath) {
    Remove-Item $DestinationPath -Force
}

try {
    Compress-Archive -Path $SourceDir -DestinationPath $DestinationPath -CompressionLevel Optimal -Force
}
catch {
    Write-Host ("压缩失败: " + $_.Exception.Message)
    exit 1
}

if (-not (Test-Path $DestinationPath)) {
    Write-Host "压缩后未生成文件: $DestinationPath"
    exit 1
}

$size = (Get-Item $DestinationPath).Length
Write-Host ("  zip 已生成: {0}  ({1:N0} 字节)" -f (Split-Path -Leaf $DestinationPath), $size)
exit 0
