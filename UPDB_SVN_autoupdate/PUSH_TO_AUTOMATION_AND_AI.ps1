# 將本專案 (UPDV_SVN_autoupdate) 上傳到 AutomationAndAI 儲存庫內的 UPDV_SVN_autoupdate 資料夾
# 請在「已安裝 Git 且 git 在 PATH 中」的 PowerShell 裡執行此腳本

$ErrorActionPreference = "Stop"
$repoUrl = "https://github.com/a5512561-creator/AutomationAndAI.git"
$projectRoot = Split-Path -Parent $PSScriptRoot   # 例如 d:\CursorProject
$cloneDir = Join-Path $projectRoot "AutomationAndAI"
$sourceDir = $PSScriptRoot                         # 本專案目錄 UPDV_SVN_autoupdate
$targetDir = Join-Path $cloneDir "UPDV_SVN_autoupdate"

Write-Host "來源: $sourceDir"
Write-Host "目標儲存庫: $cloneDir"
Write-Host ""

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    Write-Host "錯誤: 找不到 git 指令，請確認 Git 已安裝並在 PATH 中。" -ForegroundColor Red
    exit 1
}

# 若尚未 clone，先 clone
if (-not (Test-Path $cloneDir)) {
    Write-Host "正在 clone $repoUrl ..."
    Set-Location $projectRoot
    git clone $repoUrl
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

Set-Location $cloneDir
git pull

# 建立目標資料夾並複製檔案（覆蓋）
New-Item -ItemType Directory -Force -Path $targetDir | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $targetDir "docs") | Out-Null

Copy-Item (Join-Path $sourceDir "README.md") -Destination $targetDir -Force
Copy-Item (Join-Path $sourceDir ".gitignore") -Destination $targetDir -Force
Copy-Item (Join-Path $sourceDir "docs\PLAN.md") -Destination (Join-Path $targetDir "docs") -Force
Copy-Item (Join-Path $sourceDir "docs\INPUT_FORMAT.md") -Destination (Join-Path $targetDir "docs") -Force

Write-Host "已複製檔案到 $targetDir"

git add "UPDV_SVN_autoupdate"
$status = git status --short "UPDV_SVN_autoupdate"
if ([string]::IsNullOrWhiteSpace($status)) {
    Write-Host "沒有變更需要提交。"
    exit 0
}
Write-Host "變更: $status"
git commit -m "Add UPDV_SVN_autoupdate: docs and project layout for UPDB batch add"
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
Write-Host "正在 push..."
git push
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
Write-Host "完成。已上傳至 $repoUrl 的 UPDV_SVN_autoupdate 資料夾。" -ForegroundColor Green
