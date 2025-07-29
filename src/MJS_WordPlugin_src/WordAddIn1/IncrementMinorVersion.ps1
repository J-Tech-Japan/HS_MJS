# IncrementMinorVersion.ps1

param(
    [string]$Action = ""
)

# ファイルがロックされているかどうかを判定する関数
function Test-FileLock {
    param([string]$Path)
    try {
        # ファイルをReadWriteモードで開けるか試す
        $stream = [System.IO.File]::Open($Path, 'Open', 'ReadWrite', 'None')
        $stream.Close()
        return $false  # ロックされていない
    } catch {
        return $true   # ロックされている
    }
}

# AssemblyInfo.csのパスを取得
$path = Join-Path $PSScriptRoot "Properties\AssemblyInfo.cs"
$tempPath = "$path.tmp"

# AssemblyVersionとAssemblyFileVersionのパターン定義
$pattern = '(\[assembly: AssemblyVersion\(")(\d+)\.(\d+)\.(\d+)\.0("\)\])'
$filePattern = '(\[assembly: AssemblyFileVersion\(")(\d+)\.(\d+)\.(\d+)\.0("\)\])'

# ファイルがロックされている場合はエラーを出して終了
if (Test-FileLock $path) {
    Write-Error "$path は他のプロセスによってロックされています。スクリプトを中断します。"

    # 即時終了し、終了コード 1 (エラー) を返す
    exit 1
}

try {
    if ($Action -eq "reset") {
        # バージョンを 3.1.0.0 にリセットする処理
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
        # バージョンのビルド番号（第3項）をインクリメントする処理
        $incPattern = '(\[assembly: Assembly(File)?Version\("(\d+)\.(\d+)\.)(\d+)(\.0"\)\])'
        (Get-Content $path) | ForEach-Object {
            # 現在の行がバージョン情報の行であれば
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
    # エラー発生時の処理
    Write-Error "$path の更新中にエラーが発生しました: $_"
    if (Test-Path $tempPath) { Remove-Item $tempPath -Force }
    exit 1
}