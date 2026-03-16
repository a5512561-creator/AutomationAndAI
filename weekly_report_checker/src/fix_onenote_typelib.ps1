# 修正 OneNote TypeLib 的 win64 註冊路徑（Click-to-Run 已知問題）
# 寫入 HKCU\Software\Classes（不需要系統管理員權限）

$regPath = "HKCU:\Software\Classes\TypeLib\{0EA692EE-BB50-4E3C-AEF0-356D91732725}\1.1\0\win64"
$correctValue = "C:\Program Files\Microsoft Office\Root\Office16\ONENOTE.EXE\3"

# 建立完整路徑（如果不存在）
if (-not (Test-Path $regPath)) {
    New-Item -Path $regPath -Force | Out-Null
    Write-Output "Created registry path: $regPath"
}

Set-ItemProperty -Path $regPath -Name '(default)' -Value $correctValue
$result = (Get-ItemProperty -Path $regPath).'(default)'
Write-Output "Registry value set to: $result"

if ($result -eq $correctValue) {
    Write-Output "Done! TypeLib path fixed successfully."
} else {
    Write-Output "ERROR: Value mismatch."
}
