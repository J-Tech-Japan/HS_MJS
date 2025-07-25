param(
    [string]$Action = ""
)

function Test-FileLock {
    param([string]$Path)
    try {
        $stream = [System.IO.File]::Open($Path, 'Open', 'ReadWrite', 'None')
        $stream.Close()
        return $false  # ロックされていない
    } catch {
        return $true   # ロックされている
    }
}

$path = Join-Path $PSScriptRoot "Properties\AssemblyInfo.cs"
$pattern = '(\[assembly: AssemblyVersion\(")(\d+)\.(\d+)\.(\d+)\.0("\)\])'
$filePattern = '(\[assembly: AssemblyFileVersion\(")(\d+)\.(\d+)\.(\d+)\.0("\)\])'

# ロック解除待ち（最大5回リトライ、1秒ごと）
$maxRetry = 0
$retry = 0
while (Test-FileLock $path) {
    if ($retry -ge $maxRetry) {
        Write-Error "$path は他のプロセスによってロックされています。スクリプトを中断します。"
        exit 1
    }
    Start-Sleep -Seconds 1
    $retry++
}

if ($Action -eq "reset") {
    (Get-Content $path) | ForEach-Object {
        if ($_ -match $pattern) {
            '[assembly: AssemblyVersion("3.1.0.0")]'
        } elseif ($_ -match $filePattern) {
            '[assembly: AssemblyFileVersion("3.1.0.0")]'
        } else {
            $_
        }
    } | Set-Content $path
    Write-Host "バージョンを 3.1.0.0 にリセットしました。"
} else {
    # 通常のインクリメント（従来通り）
    $incPattern = '(\[assembly: Assembly(File)?Version\("(\d+)\.(\d+)\.)(\d+)(\.0"\)\])'
    (Get-Content $path) | ForEach-Object {
        if ($_ -match $incPattern) {
            $majorMinor = "$($matches[3]).$($matches[4])"
            $build = [int]$matches[5] + 1
            $_ -replace "(\d+)\.(\d+)\.\d+\.0", "$majorMinor.$build.0"
        } else {
            $_
        }
    } | Set-Content $path
}