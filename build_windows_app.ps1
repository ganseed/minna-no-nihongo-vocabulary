$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$BuildRoot = Join-Path $ProjectRoot "build"
$ReleaseRoot = Join-Path $ProjectRoot "dist\windows"

python -m pip install -r (Join-Path $ProjectRoot "requirements.txt") pyinstaller
python -m PyInstaller --noconfirm --clean --onefile --windowed `
  --name "大家的日语词库生成器" `
  --distpath $ReleaseRoot `
  --workpath (Join-Path $BuildRoot "windows") `
  --specpath $BuildRoot `
  (Join-Path $ProjectRoot "desktop_app.py")

New-Item -ItemType Directory -Force (Join-Path $ReleaseRoot "data"), (Join-Path $ReleaseRoot "output") | Out-Null
Copy-Item (Join-Path $ProjectRoot "data\vocabulary.json") (Join-Path $ReleaseRoot "data\vocabulary.json") -Force
Copy-Item (Join-Path $ProjectRoot "应用使用说明.txt") (Join-Path $ReleaseRoot "使用说明.txt") -Force
Write-Host "Windows 应用已生成：$ReleaseRoot"
