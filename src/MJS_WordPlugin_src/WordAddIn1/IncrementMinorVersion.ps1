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
$tempPath = "$path.tmp"
$pattern = '(\[assembly: AssemblyVersion\(")(\d+)\.(\d+)\.(\d+)\.0("\)\])'
$filePattern = '(\[assembly: AssemblyFileVersion\(")(\d+)\.(\d+)\.(\d+)\.0("\)\])'

# ロックされている場合は即終了
if (Test-FileLock $path) {
    Write-Error "$path は他のプロセスによってロックされています。スクリプトを中断します。"
    exit 1
}

try {
    if ($Action -eq "reset") {
        (Get-Content $path) | ForEach-Object {
            if ($_ -match $pattern) {
                '[assembly: AssemblyVersion("3.1.0.0")]'
            } elseif ($_ -match $filePattern) {
                '[assembly: AssemblyFileVersion("3.1.0.0")]'
            } else {
                $_
            }
        } | Set-Content $tempPath
        Move-Item -Force $tempPath $path
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
        } | Set-Content $tempPath
        Move-Item -Force $tempPath $path
    }
} catch {
    Write-Error "$path の更新中にエラーが発生しました: $_"
    if (Test-Path $tempPath) { Remove-Item $tempPath -Force }
    exit 1
}