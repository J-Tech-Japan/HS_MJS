# WordAddIn1\Properties\AssemblyInfo.cs のビルド番号（3桁目）を+1する
$path = "WordAddIn1\Properties\AssemblyInfo.cs"
$pattern = '(\[assembly: Assembly(File)?Version\("3\.1\.)(\d+)(\.0"\)\])'
(Get-Content $path) | ForEach-Object {
    if ($_ -match $pattern) {
        $newBuild = [int]$matches[3] + 1
        $_ -replace "3\.1\.\d+\.0", "3.1.$newBuild.0"
    } else {
        $_
    }
} | Set-Content $path