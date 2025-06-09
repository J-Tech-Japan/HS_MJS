# IncrementMinorVersion.ps1
$path = "WordAddIn1\Properties\AssemblyInfo.cs"
$versionPattern = '\[assembly: Assembly(File)?Version\("3\.0\.(\d+)\.0"\)\]'
(Get-Content $path) | ForEach-Object {
    if ($_ -match $versionPattern) {
        $current = [int]$matches[2]
        $next = $current + 1
        $_ -replace "3\.0\.\d+\.0", "3.0.$next.0"
    } else {
        $_
    }
} | Set-Content $path