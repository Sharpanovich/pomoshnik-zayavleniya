# Sborka programmy «Pomoschnik po zayavleniyam»
$ErrorActionPreference = 'Stop'
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Poisk csc.exe
$csc = @(
    "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe",
    "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\csc.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $csc) {
    Write-Host "Compiler csc.exe not found (.NET Framework 4.x required)" -ForegroundColor Red
    exit 1
}

$frameworkDir = Split-Path -Parent $csc
$exeName = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String('0J/QvtC80L7RidC90LjQuiDQv9C+INC30LDRj9Cy0LvQtdC90LjRj9C8LmV4ZQ=='))
$out = Join-Path $dir $exeName

& $csc /nologo /target:winexe /codepage:65001 `
    /out:$out /win32icon:"$(Join-Path $dir 'app.ico')" `
    /r:"$(Join-Path $frameworkDir 'System.IO.Compression.dll')" `
    /r:"$(Join-Path $frameworkDir 'System.IO.Compression.FileSystem.dll')" `
    /r:System.dll /r:System.Core.dll /r:System.Xml.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll `
    "$(Join-Path $dir 'App.cs')"

if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Host "Gotovo! Programma sobrana: $out" -ForegroundColor Green
} else {
    Write-Host "Oshibka sborki." -ForegroundColor Red
    exit 1
}

$uploaderCs = Join-Path $dir 'GitHubUploader.cs'
if (Test-Path $uploaderCs) {
    $uploaderExeName = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String('0JfQsNCz0YDRg9C30LrQsCDQvdCwIEdpdEh1Yi5leGU='))
    $uploaderOut = Join-Path $dir $uploaderExeName
    & $csc /nologo /target:winexe /codepage:65001 /out:$uploaderOut /win32icon:"$(Join-Path $dir 'app.ico')" `
        /r:"$(Join-Path $frameworkDir 'System.IO.Compression.dll')" `
        /r:"$(Join-Path $frameworkDir 'System.IO.Compression.FileSystem.dll')" `
        /r:System.Web.Extensions.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll $uploaderCs
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Gotovo! Uploader sobran: $uploaderOut" -ForegroundColor Green
    }
}
