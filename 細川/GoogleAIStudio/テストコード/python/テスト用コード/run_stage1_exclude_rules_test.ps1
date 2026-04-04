#Requires -Version 5.1
<#
.SYNOPSIS
  段階1 + 「設定_配台不要工程」の COM 保存テスト（任意で E 列に "1234"）

.DESCRIPTION
  TASK_INPUT_WORKBOOK と EXCLUDE_RULES_TEST_E1234 を設定して task_extract_stage1.py を実行します。
  Python は既定で L:\anaconda3\python.exe（-Python で変更可）。

.PARAMETER Workbook
  マクロブックのフルパス（「加工計画DATA」シート必須）

.PARAMETER TestE1234
  指定すると EXCLUDE_RULES_TEST_E1234=1（E 列テスト行に "1234"）

.PARAMETER E1234Row
  テスト書き込み行（既定 9）

.PARAMETER Python
  python.exe のフルパス
#>
param(
    [string] $Workbook = "",
    [switch] $TestE1234,
    [int] $E1234Row = 9,
    [string] $Python = "L:\anaconda3\python.exe"
)

$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$parent = Split-Path -Parent $here
if (Test-Path -LiteralPath (Join-Path $parent "planning_core.py")) {
    $repoRoot = $parent
} else {
    $repoRoot = $here
}
Set-Location $repoRoot

if (-not $Workbook) {
    $Workbook = Join-Path $repoRoot "生産管理_AI配台テスト.xlsm"
}
if (-not (Test-Path -LiteralPath $Workbook)) {
    Write-Error "ブックが見つかりません: $Workbook"
}
if (-not (Test-Path -LiteralPath $Python)) {
    Write-Error "Python が見つかりません: $Python （-Python で指定）"
}

$env:TASK_INPUT_WORKBOOK = (Resolve-Path -LiteralPath $Workbook).Path
Remove-Item Env:EXCLUDE_RULES_TEST_E1234 -ErrorAction SilentlyContinue
Remove-Item Env:EXCLUDE_RULES_TEST_E1234_ROW -ErrorAction SilentlyContinue

if ($TestE1234) {
    $env:EXCLUDE_RULES_TEST_E1234 = "1"
    $env:EXCLUDE_RULES_TEST_E1234_ROW = "$E1234Row"
    Write-Host "[TEST] EXCLUDE_RULES_TEST_E1234=1, ROW=$E1234Row" -ForegroundColor Yellow
} else {
    Write-Host "[RUN] E列テストなし（通常の段階1）" -ForegroundColor Cyan
}

Write-Host "TASK_INPUT_WORKBOOK=$($env:TASK_INPUT_WORKBOOK)"
Write-Host "Python=$Python"
& $Python (Join-Path $repoRoot "task_extract_stage1.py")
exit $LASTEXITCODE
