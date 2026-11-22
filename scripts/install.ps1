param([string[]]$PreviewRoots, [string]$LegacyWordPython)

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$targetPluginDir = Join-Path $env:APPDATA "BetterGravity\plugins\doc-preview"
$settingsPath = Join-Path $env:APPDATA "BetterGravity\settings.json"

Push-Location $projectRoot
try {
  if (-not (Test-Path (Join-Path $projectRoot "node_modules"))) {
    Write-Host "[1/3] 安装 npm 依赖..."
    npm install
    if ($LASTEXITCODE -ne 0) { throw "安装依赖失败。" }
  }
  Write-Host "[2/3] 构建插件产物..."
  node scripts/build.mjs
  if ($LASTEXITCODE -ne 0) { throw "构建插件失败。" }
} finally {
  Pop-Location
}

Write-Host "[3/3] 同步插件到 $targetPluginDir ..."
New-Item -ItemType Directory -Force -Path (Join-Path $targetPluginDir "styles") | Out-Null
Copy-Item -Force (Join-Path $projectRoot "dist\plugin.json") (Join-Path $targetPluginDir "plugin.json")
Copy-Item -Force (Join-Path $projectRoot "dist\styles\doc-preview.css") (Join-Path $targetPluginDir "styles\doc-preview.css")
$legacyWordArgs = @()
if ($LegacyWordPython) { $legacyWordArgs = @("--legacy-word-python", $LegacyWordPython) }
node (Join-Path $PSScriptRoot "install-large-preview.mjs") $targetPluginDir @legacyWordArgs @PreviewRoots
if ($LASTEXITCODE -ne 0) { throw "安装插件失败。" }

if (Test-Path -LiteralPath $settingsPath) {
  $settings = Get-Content -LiteralPath $settingsPath -Raw -Encoding UTF8 | ConvertFrom-Json
} else {
  $settings = [PSCustomObject]@{
    schemaVersion = 1
    themes = [PSCustomObject]@{ enabled = @() }
    plugins = [PSCustomObject]@{ developerMode = $false; enabled = @() }
    reapplyAfterHostUpdate = $true
  }
}

if ($null -eq $settings.plugins) {
  $settings | Add-Member -Force -NotePropertyName plugins -NotePropertyValue (
    [PSCustomObject]@{ developerMode = $false; enabled = @() }
  )
}
$enabled = @($settings.plugins.enabled | Where-Object { $_ })
if ($enabled -notcontains "doc-preview") {
  $settings.plugins | Add-Member -Force -NotePropertyName enabled -NotePropertyValue @($enabled + "doc-preview")
  $json = $settings | ConvertTo-Json -Depth 100
  [System.IO.File]::WriteAllText($settingsPath, $json + [Environment]::NewLine)
}

if ($settings.plugins.developerMode -eq $true) {
  Write-Host "已安装并登记 Doc Preview，BetterGravity 会自动热重载。"
} else {
  Write-Warning "已复制并登记插件文件。要启用 Doc Preview，请在 BetterGravity 设置里开启开发者模式。"
}
if ($PreviewRoots.Count -gt 0 -or $LegacyWordPython) {
  Write-Host "重启本地预览组件或 Antigravity 以加载预览配置。"
}
