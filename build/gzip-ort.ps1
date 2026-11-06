param(
    [Parameter(Mandatory = $true)][string]$Source,
    [Parameter(Mandatory = $true)][string]$Target
)

# 将 OnnxRuntime 原生库压缩为 .gz，供单文件发布内嵌使用。
# 原生 DLL 用 Deflate/GZip 可压到约 38%，比直接嵌入省 6~7 MB。

if (-not (Test-Path $Source)) {
    Write-Error "源文件不存在: $Source"
    exit 1
}

$sourceInfo = Get-Item $Source
if ((Test-Path $Target) -and ((Get-Item $Target).LastWriteTimeUtc -ge $sourceInfo.LastWriteTimeUtc)) {
    Write-Host "已是最新: $Target"
    exit 0
}

$targetDir = Split-Path -Parent $Target
if ($targetDir -and -not (Test-Path $targetDir)) {
    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
}

$input = [System.IO.File]::OpenRead($Source)
try {
    $output = [System.IO.File]::Create($Target)
    try {
        $gzip = New-Object System.IO.Compression.GZipStream($output, [System.IO.Compression.CompressionLevel]::Optimal)
        try { $input.CopyTo($gzip) } finally { $gzip.Dispose() }
    } finally { $output.Dispose() }
} finally { $input.Dispose() }

$ratio = (Get-Item $Target).Length / $sourceInfo.Length
Write-Host ("已压缩 {0} → {1} ({2:P0})" -f $sourceInfo.Name, (Split-Path -Leaf $Target), $ratio)
