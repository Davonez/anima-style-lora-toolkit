#requires -Version 5.1
<#
  InstallPreset —— 把 presets\ 里的训练预设装进 Next-Trainer
  ============================================================================
  用法：
    双击 安装预设.bat
    或 powershell -NoProfile -ExecutionPolicy Bypass -File .\InstallPreset.ps1 -NextTrainer "D:\Next-Trainer"

  会做三件事：
    1. 找到 <Next-Trainer>\config\presets 目录；
    2. 把本仓库 presets\*.toml 复制进去（同名覆盖，会先备份成 *.bak）；
    3. 打印结果和后续操作。
#>
[CmdletBinding()]
param(
    [string]$NextTrainer,
    [switch]$Yes
)

$ErrorActionPreference = 'Stop'
$here = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$presetSrc = Join-Path $here 'presets'
$interactive = ([Environment]::UserInteractive -and -not [Console]::IsInputRedirected)

Write-Host ''
Write-Host '=============================================' -ForegroundColor Cyan
Write-Host '  安装训练预设 → Next-Trainer' -ForegroundColor Cyan
Write-Host '=============================================' -ForegroundColor Cyan

if (-not (Test-Path -LiteralPath $presetSrc)) {
    Write-Host "找不到 presets 目录：$presetSrc" -ForegroundColor Red
    Write-Host '请把本脚本和 presets\ 放在同一目录下。' -ForegroundColor Red
    exit 1
}
$presetFiles = @(Get-ChildItem -LiteralPath $presetSrc -Filter '*.toml' -File)
if ($presetFiles.Count -eq 0) { Write-Host 'presets 目录里没有 .toml。' -ForegroundColor Red; exit 1 }
Write-Host ("待安装：{0} 个预设" -f $presetFiles.Count)
foreach ($f in $presetFiles) { Write-Host ("    " + $f.Name) }

function Resolve-PresetDir([string]$p) {
    if (-not $p) { return $null }
    $p = $p.Trim().TrimEnd('\').TrimEnd('/')
    if (-not (Test-Path -LiteralPath $p)) { return $null }
    foreach ($cand in @((Join-Path $p 'config\presets'), $p)) {
        if (Test-Path -LiteralPath (Join-Path $cand '.')) {
            if ((Split-Path $cand -Leaf) -eq 'presets') { return $cand }
        }
    }
    return $null
}

$dest = $null
if ($NextTrainer) {
    $dest = Resolve-PresetDir $NextTrainer
    if (-not $dest) {
        Write-Host "`n这个目录里没有 config\presets：$NextTrainer" -ForegroundColor Red
        Write-Host '请选 Next-Trainer 的根目录（里面能看到 config、sd-models、mikazuki 这些文件夹）。' -ForegroundColor Yellow
        exit 1
    }
}

if (-not $dest) {
    if (-not $interactive -or $Yes) {
        Write-Host "请用 -NextTrainer 指定 Next-Trainer 根目录。" -ForegroundColor Red
        exit 1
    }
    Write-Host ''
    Write-Host '请选择 Next-Trainer 的根目录（里面能看到 config、sd-models、mikazuki 文件夹）' -ForegroundColor Cyan
    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = '选择 Next-Trainer 根目录（包含 config\presets）'
    $dlg.ShowNewFolderButton = $false
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { Write-Host '已取消。' -ForegroundColor Yellow; exit 0 }
    $dest = Resolve-PresetDir $dlg.SelectedPath
    if (-not $dest) {
        Write-Host "`n选中的目录里没有 config\presets。请重选 Next-Trainer 根目录。" -ForegroundColor Red
        exit 1
    }
}

Write-Host ''
Write-Host ("目标目录：$dest") -ForegroundColor Green

$ok = 0
foreach ($f in $presetFiles) {
    $target = Join-Path $dest $f.Name
    try {
        if (Test-Path -LiteralPath $target) {
            # 备份必须放在 presets 目录之外：Next-Trainer 会把这个目录里的每个文件
            # 都当预设解析，留 .bak 会多出一条重复预设。
            $bakDir = Join-Path (Split-Path $dest -Parent) 'presets_backup'
            if (-not (Test-Path -LiteralPath $bakDir)) { New-Item -ItemType Directory -Path $bakDir -Force | Out-Null }
            $bak = Join-Path $bakDir ($f.Name + '.' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.bak')
            Copy-Item -LiteralPath $target -Destination $bak -Force
            Write-Host ("    已备份旧同名文件 → {0}" -f $bak) -ForegroundColor DarkYellow
        }
        Copy-Item -LiteralPath $f.FullName -Destination $target -Force
        Write-Host ("    ✓ " + $f.Name) -ForegroundColor Green
        $ok++
    } catch {
        Write-Host ("    ✗ " + $f.Name + "：" + $_.Exception.Message) -ForegroundColor Red
    }
}

Write-Host ''
Write-Host ("安装完成：{0}/{1}" -f $ok, $presetFiles.Count) -ForegroundColor Green
Write-Host ''
Write-Host '下一步：' -ForegroundColor Cyan
Write-Host '  1) 回到 Next-Trainer（不用重启，打开「训练预设」时会自动重新读取）'
Write-Host '  2) 模型选 Anima DiT、引擎 Kohya-ss、目标 LoRA'
Write-Host '  3) 工具栏「训练预设」→ 选 "Anima 画风 LoRA v3..."'
Write-Host '  4) 应用后只改两个框：训练数据集路径、模型保存名称（后者务必改名）'
Write-Host ''
Write-Host '没看到预设？确认 [metadata] train_type = "sd3-lora"（GUI 用它过滤当前页面）。' -ForegroundColor DarkGray
Write-Host ''
