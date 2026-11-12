# Dong goi "HT CREATIVE AGENT" thanh bo cai Windows (HTCreativeAgent_Setup.exe) va ban portable (.zip).
# Chay bang DONG_GOI.bat. May dong goi can co mang; KHONG can cai Python.
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$PyVer   = '3.12.10'
$Here    = $PSScriptRoot
$Src     = Split-Path $Here -Parent
$Build   = Join-Path $Here 'build'
$AppRoot = Join-Path $Build 'TaiVideo'
$PyDir   = Join-Path $AppRoot 'python'
$AppDir  = Join-Path $AppRoot 'app'
$Py      = Join-Path $PyDir 'python.exe'

function Buoc($t) { Write-Host ""; Write-Host "==> $t" -ForegroundColor Cyan }

Buoc "1/6 Chuan bi thu muc"
if (Test-Path $Build) { Remove-Item $Build -Recurse -Force }
New-Item -ItemType Directory -Force -Path $PyDir, $AppDir | Out-Null

Buoc "2/6 Tai Python $PyVer (ban nhung, ~11 MB)"
$zip = Join-Path $Build 'python-embed.zip'
Invoke-WebRequest "https://www.python.org/ftp/python/$PyVer/python-$PyVer-embed-amd64.zip" -OutFile $zip -UseBasicParsing
Expand-Archive $zip -DestinationPath $PyDir -Force
Remove-Item $zip
$pth = Get-ChildItem $PyDir -Filter 'python*._pth' | Select-Object -First 1
if (-not $pth) { throw "Khong thay file ._pth trong ban Python nhung" }
$zipName = (Get-ChildItem $PyDir -Filter 'python*.zip' | Select-Object -First 1).Name
Set-Content -Path $pth.FullName -Encoding ASCII -Value @($zipName, '.', '..\app', 'import site')

Buoc "3/6 Cai pip"
$getpip = Join-Path $Build 'get-pip.py'
Invoke-WebRequest 'https://bootstrap.pypa.io/get-pip.py' -OutFile $getpip -UseBasicParsing
& $Py $getpip --no-warn-script-location -q
if ($LASTEXITCODE -ne 0) { throw "Cai pip that bai" }
Remove-Item $getpip

Buoc "4/6 Cai thu vien (yt-dlp, playwright, curl_cffi, ffmpeg...) - mat vai phut"
& $Py -m pip install -q --no-warn-script-location --disable-pip-version-check "yt-dlp[default,curl-cffi]" openpyxl imageio-ffmpeg playwright
if ($LASTEXITCODE -ne 0) { throw "Cai thu vien that bai" }
& $Py -c "import yt_dlp, openpyxl, imageio_ffmpeg, playwright, curl_cffi; print('  yt-dlp', yt_dlp.version.__version__, '- cac thu vien OK')"
if ($LASTEXITCODE -ne 0) { throw "Kiem tra thu vien that bai" }

Buoc "5/6 Chep file cua tool"
foreach ($f in 'tai_video.py', 'giao_dien.py', 'giao_dien.html', 'HUONG_DAN.txt', 'icon.ico') {
    Copy-Item (Join-Path $Src $f) $AppDir
}
New-Item -ItemType Directory -Force -Path (Join-Path $AppDir 'cookies') | Out-Null
Set-Content -Path (Join-Path $AppRoot 'Mo HT Creative Agent.bat') -Encoding ASCII -Value @(
    '@echo off', 'cd /d "%~dp0app"', 'start "" "%~dp0python\pythonw.exe" "%~dp0app\giao_dien.py" %*')
# Thu nhanh: tool khoi dong duoc trong ban dong goi
$env:TAIVIDEO_NO_BROWSER = '1'; $env:TAIVIDEO_NO_UPDATE = '1'
& $Py -c "import sys; sys.argv=['x']; import giao_dien, tai_video; print('  Tool nap duoc trong ban dong goi: OK')"
if ($LASTEXITCODE -ne 0) { throw "Tool khong nap duoc trong ban dong goi" }
Remove-Item Env:TAIVIDEO_NO_BROWSER, Env:TAIVIDEO_NO_UPDATE
Get-ChildItem $AppRoot -Recurse -Directory -Filter '__pycache__' | Remove-Item -Recurse -Force

Buoc "6/6 Tao bo cai"
$portable = Join-Path $Here 'HTCreativeAgent_Portable.zip'
if (Test-Path $portable) { Remove-Item $portable }
Compress-Archive -Path $AppRoot -DestinationPath $portable
Write-Host "  Ban portable: $portable"

function TimISCC {
    foreach ($p in @("${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe", "$env:ProgramFiles\Inno Setup 6\ISCC.exe",
                     "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe")) {
        if ($p -and (Test-Path $p)) { return $p }
    }
    return $null
}
$iscc = TimISCC
if (-not $iscc -and (Get-Command winget -ErrorAction SilentlyContinue)) {
    Write-Host "  Chua co Inno Setup -> dang cai bang winget..."
    try { winget install -e --id JRSoftware.InnoSetup --silent --accept-package-agreements --accept-source-agreements | Out-Null } catch {}
    $iscc = TimISCC
}
if ($iscc) {
    & $iscc /Q (Join-Path $Here 'TaiVideo.iss')
    if ($LASTEXITCODE -ne 0) { throw "Inno Setup bao loi khi tao bo cai" }
    Write-Host ""
    Write-Host "XONG. Bo cai: $(Join-Path $Here 'HTCreativeAgent_Setup.exe')" -ForegroundColor Green
} else {
    Write-Host ""
    Write-Host "Da tao ban portable, CHUA tao duoc HTCreativeAgent_Setup.exe vi may chua co Inno Setup." -ForegroundColor Yellow
    Write-Host "Cai Inno Setup (mien phi) tai https://jrsoftware.org/isdl.php roi chay lai DONG_GOI.bat."
}
