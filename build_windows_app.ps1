$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$BuildRoot = Join-Path $ProjectRoot "build\windows"
$DistRoot = Join-Path $ProjectRoot "dist\windows"
$PackageRoot = Join-Path $DistRoot "大家的日语词库-Windows"

function Invoke-Python {
  param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Arguments)
  & python @Arguments
  if ($LASTEXITCODE -ne 0) {
    throw "Python 命令执行失败，退出代码：$LASTEXITCODE"
  }
}

if (Test-Path -LiteralPath $PackageRoot) {
  Remove-Item -LiteralPath $PackageRoot -Recurse -Force
}

Invoke-Python -m pip install --upgrade pip
Invoke-Python -m pip install -r (Join-Path $ProjectRoot "requirements.txt") pyinstaller

Invoke-Python -m PyInstaller --noconfirm --clean --onefile --windowed `
  --name "大家的日语词库工具" `
  --distpath $PackageRoot `
  --workpath $BuildRoot `
  --specpath $BuildRoot `
  (Join-Path $ProjectRoot "desktop_app.py")

$Executable = Join-Path $PackageRoot "大家的日语词库工具.exe"
if (-not (Test-Path -LiteralPath $Executable)) {
  throw "构建失败：未找到 $Executable"
}

New-Item -ItemType Directory -Force `
  (Join-Path $PackageRoot "data"), `
  (Join-Path $PackageRoot "output") | Out-Null

Copy-Item `
  (Join-Path $ProjectRoot "data\vocabulary.json") `
  (Join-Path $PackageRoot "data\vocabulary.json") `
  -Force
Copy-Item (Join-Path $ProjectRoot "README.md") (Join-Path $PackageRoot "README.md") -Force

Write-Host "Windows 发布目录：$PackageRoot"
