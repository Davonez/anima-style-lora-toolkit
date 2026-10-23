#requires -Version 5.1
<#
  风格对照 v1.0 —— 把「像不像」变成数字
  ============================================================================
  对应《画风 LoRA 实战经验》第一、四节：训练前必须能回答「这版比上版更像吗」，
  靠的不是感觉，是三个量：
      平均亮度（两个口径：通道均值 / luma709）
      平均饱和度（对风格最敏感）
      同 seed 像素差（衡量风格与色调差，不是贴合度）

  用法（三种）：
    1. 双击 一键对照.bat → 选文件夹 → 出报告
    2. 单目录体检（只看数据集本身有多灰/多暗）：
         powershell -NoProfile -ExecutionPolicy Bypass -File .\StyleCompare.ps1 -A "D:\datasets\my_artist"
    3. 两组对照（训练原图 vs 出图，或 零权重 vs 挂 LoRA）：
         powershell -NoProfile -ExecutionPolicy Bypass -File .\StyleCompare.ps1 `
             -A "D:\datasets\my_artist\originals" -B "D:\datasets\my_artist\gen_w0.8"

  配对规则：默认按同名文件配对；两边文件数不一致或名字对不上时，退化为「按文件名排序后逐位配对」并提示。
#>
[CmdletBinding()]
param(
    [string]$A,
    [string]$B,
    [string]$Out,
    [int]$Max = 400,
    [ValidateSet('name', 'order')]
    [string]$PairBy = 'name',
    [switch]$NoDiff,
    [switch]$NoDialog
)

$ErrorActionPreference = 'Stop'
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$script:Interactive = ([Environment]::UserInteractive -and -not [Console]::IsInputRedirected)
$imgExt = @('.png', '.jpg', '.jpeg', '.webp', '.bmp', '.gif', '.jfif')

function Select-Folder([string]$desc) {
    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = $desc
    $dlg.ShowNewFolderButton = $false
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
    return $dlg.SelectedPath
}

function Get-ImgStat([string]$path) {
    try {
        $bmp = [System.Drawing.Bitmap]::FromFile($path)
        $w0 = $bmp.Width; $h0 = $bmp.Height
        $w = 64
        $h = [int]($h0 * 64 / $w0)
        if ($h -lt 1) { $h = 64 }
        $th = $bmp.GetThumbnailImage($w, $h, $null, [IntPtr]::Zero)
        $sumL = 0.0; $sumM = 0.0; $sumS = 0.0; $n = 0; $grey = 0
        for ($y = 0; $y -lt $th.Height; $y += 2) {
            for ($x = 0; $x -lt $th.Width; $x += 2) {
                $c = $th.GetPixel($x, $y)
                $r = $c.R / 255.0; $g = $c.G / 255.0; $b = $c.B / 255.0
                $mx = [Math]::Max($r, [Math]::Max($g, $b))
                $mn = [Math]::Min($r, [Math]::Min($g, $b))
                $sat = if ($mx -gt 0) { ($mx - $mn) / $mx } else { 0 }
                $sumL += (0.2126 * $r + 0.7152 * $g + 0.0722 * $b)
                $sumM += (($r + $g + $b) / 3.0)
                $sumS += $sat
                $n++
                if ($sat -lt 0.08) { $grey++ }
            }
        }
        # 32x32 缩略（用于像素差；构图被 seed 锁死时它衡量风格与色调差）
        $dth = $bmp.GetThumbnailImage(32, 32, $null, [IntPtr]::Zero)
        $px = New-Object 'double[]' 3072
        $k = 0
        for ($y = 0; $y -lt 32; $y++) {
            for ($x = 0; $x -lt 32; $x++) {
                $c = $dth.GetPixel($x, $y)
                $px[$k] = $c.R; $px[$k + 1] = $c.G; $px[$k + 2] = $c.B
                $k += 3
            }
        }
        $bmp.Dispose()
        return [pscustomobject]@{
            W       = $w0
            H       = $h0
            Luma709 = [Math]::Round($sumL / $n, 4)
            Mean255 = [Math]::Round(255.0 * $sumM / $n, 1)
            Sat     = [Math]::Round($sumS / $n, 4)
            GrayPct = [Math]::Round(100.0 * $grey / $n, 1)
            Pixels  = $px
        }
    } catch { return $null }
}

function Get-Files([string]$dir) {
    return @(Get-ChildItem -LiteralPath $dir -File -ErrorAction SilentlyContinue |
            Where-Object { $imgExt -contains $_.Extension.ToLowerInvariant() } | Sort-Object Name)
}

function Compare-Pixels($a, $b) {
    if (-not $a.Pixels -or -not $b.Pixels) { return $null }
    $sum = 0.0
    for ($i = 0; $i -lt 3072; $i++) { $sum += [Math]::Abs($a.Pixels[$i] - $b.Pixels[$i]) }
    return [Math]::Round($sum / 3072.0, 2)
}

# ---------------------------------------------------------------- 入口
Write-Host ''
Write-Host '=============================================' -ForegroundColor Cyan
Write-Host '  风格对照 v1.0   亮度 / 饱和 / 像素差' -ForegroundColor Cyan
Write-Host '=============================================' -ForegroundColor Cyan

Add-Type -AssemblyName System.Drawing | Out-Null

if (-not $A) {
    if (-not $script:Interactive -or $NoDialog) {
        Write-Host '请用 -A 指定目录（先看的那组，通常是训练原图或零权重出图）。' -ForegroundColor Red
        exit 1
    }
    $A = Select-Folder '选择 A 组图片（例如训练原图 或 零权重出图）'
    if (-not $A) { Write-Host '已取消。' -ForegroundColor Yellow; exit 0 }
}
if (-not (Test-Path -LiteralPath $A)) { Write-Host "目录不存在：$A" -ForegroundColor Red; exit 1 }
$A = (Resolve-Path -LiteralPath $A).Path

if (-not $B -and -not $NoDiff -and $script:Interactive -and -not $NoDialog) {
    Write-Host ''
    Write-Host '要对照的另一组在哪？（例如挂 LoRA 的出图；直接取消 = 只做单目录体检）' -ForegroundColor Cyan
    $B = Select-Folder '选择 B 组图片（例如挂 LoRA 的出图）；取消 = 单目录体检'
}
if ($B) {
    if (-not (Test-Path -LiteralPath $B)) { Write-Host "目录不存在：$B" -ForegroundColor Red; exit 1 }
    $B = (Resolve-Path -LiteralPath $B).Path
}

$filesA = Get-Files $A
if ($filesA.Count -eq 0) { Write-Host "A 组没有图片：$A" -ForegroundColor Red; exit 1 }
$filesB = @()
if ($B) {
    $filesB = Get-Files $B
    if ($filesB.Count -eq 0) { Write-Host "B 组没有图片：$B" -ForegroundColor Red; exit 1 }
}

Write-Host ''
Write-Host ("A 组：{0} 张   {1}" -f $filesA.Count, $A)
if ($B) { Write-Host ("B 组：{0} 张   {1}" -f $filesB.Count, $B) }
Write-Host ''

# ---------------------------------------------------------------- 逐张测量
$statA = @($filesA | Select-Object -First $Max | ForEach-Object {
        $s = Get-ImgStat $_.FullName
        if ($s) { $s | Add-Member -NotePropertyName Name -NotePropertyValue $_.Name -Force; $s }
    })
if ($statA.Count -eq 0) { Write-Host 'A 组没有读成功的图片。' -ForegroundColor Red; exit 1 }

$pairs = @()
if ($B) {
    $statB = @($filesB | Select-Object -First $Max | ForEach-Object {
            $s = Get-ImgStat $_.FullName
            if ($s) { $s | Add-Member -NotePropertyName Name -NotePropertyValue $_.Name -Force; $s }
        })
    $byNameB = @{}
    foreach ($x in $statB) { $byNameB[$x.Name.ToLowerInvariant()] = $x }
    $byStemB = @{}
    foreach ($x in $statB) { $byStemB[([System.IO.Path]::GetFileNameWithoutExtension($x.Name)).ToLowerInvariant()] = $x }

    $mode = $PairBy
    $probe = @($statA | Where-Object { $byNameB.ContainsKey($_.Name.ToLowerInvariant()) })
    if ($mode -eq 'name' -and $probe.Count -lt [Math]::Min($statA.Count, $statB.Count) * 0.5) {
        $mode = 'order'
        Write-Host '同名文件对不上，改为按文件名排序逐位配对。' -ForegroundColor Yellow
    }
    for ($i = 0; $i -lt $statA.Count; $i++) {
        $xa = $statA[$i]
        $xb = $null
        if ($mode -eq 'name') {
            $k1 = $xa.Name.ToLowerInvariant()
            $k2 = [System.IO.Path]::GetFileNameWithoutExtension($xa.Name).ToLowerInvariant()
            if ($byNameB.ContainsKey($k1)) { $xb = $byNameB[$k1] }
            elseif ($byStemB.ContainsKey($k2)) { $xb = $byStemB[$k2] }
        } else {
            if ($i -lt $statB.Count) { $xb = $statB[$i] }
        }
        if (-not $xb) { continue }
        $pairs += [pscustomobject]@{
            Name    = $xa.Name
            AL      = $xa.Luma709; BL = $xb.Luma709
            AM      = $xa.Mean255; BM = $xb.Mean255
            AS      = $xa.Sat; BS = $xb.Sat
            AG      = $xa.GrayPct; BG = $xb.GrayPct
            Diff    = $(if ($NoDiff) { $null } else { Compare-Pixels $xa $xb })
        }
    }
}

# ---------------------------------------------------------------- 汇总
function Avg($list, [string]$prop) {
    $v = @($list | ForEach-Object { $_.$prop } | Where-Object { $null -ne $_ })
    if ($v.Count -eq 0) { return 0 }
    return ($v | Measure-Object -Average).Average
}
$aL = Avg $statA 'Luma709'; $aM = Avg $statA 'Mean255'; $aS = Avg $statA 'Sat'; $aG = Avg $statA 'GrayPct'
$grayA = @($statA | Where-Object { $_.Sat -lt 0.08 }).Count
Write-Host ''
Write-Host '---------- A 组 ----------' -ForegroundColor Cyan
Write-Host ("  平均亮度   luma709 = {0:N3}      通道均值 = {1:N1}/255" -f $aL, $aM)
Write-Host ("  平均饱和   {0:N3}     近黑白图 {1}/{2}（{3:N0}%）" -f $aS, $grayA, $statA.Count, (100.0 * $grayA / $statA.Count))
Write-Host '  （口径提醒：luma709 与通道均值能差几十点，报告里写清用的是哪个）'

$md = New-Object System.Text.StringBuilder
[void]$md.AppendLine("# 风格对照报告 $(Get-Date -Format 'yyyy-MM-dd HH:mm')")
[void]$md.AppendLine('')
[void]$md.AppendLine("## A 组：``$A``")
[void]$md.AppendLine('')
[void]$md.AppendLine('| 指标 | 值 |')
[void]$md.AppendLine('|---|---|')
[void]$md.AppendLine(("| 张数 | {0} |" -f $statA.Count))
[void]$md.AppendLine(("| 平均亮度 luma709 | **{0:N3}** |" -f $aL))
[void]$md.AppendLine(("| 平均亮度 通道均值 | {0:N1}/255 |" -f $aM))
[void]$md.AppendLine(("| 平均饱和 | **{0:N3}** |" -f $aS))
[void]$md.AppendLine(("| 近黑白图（饱和<0.08） | {0} 张（{1:N0}%） |" -f $grayA, (100.0 * $grayA / $statA.Count)))
[void]$md.AppendLine('')

if ($pairs.Count -gt 0) {
    $bL = Avg $pairs 'BL'; $bM = Avg $pairs 'BM'; $bS = Avg $pairs 'BS'; $bG = Avg $pairs 'BG'
    $dL = $bL - $aL; $dS = $bS - $aS
    $diffArr = @($pairs | ForEach-Object { $_.Diff } | Where-Object { $null -ne $_ })
    $avgDiff = if ($diffArr.Count) { ($diffArr | Measure-Object -Average).Average } else { 0 }
    $grayB = @($pairs | Where-Object { $_.BS -lt 0.08 }).Count

    Write-Host ''
    Write-Host '---------- B 组（配对 ' -NoNewline -ForegroundColor Cyan
    Write-Host ($pairs.Count) -NoNewline -ForegroundColor Cyan
    Write-Host ' 对）----------' -ForegroundColor Cyan
    Write-Host ("  平均亮度   luma709 = {0:N3}      Δ = {1:+0.000;-0.000;0.000}（{2:N1} 点/255）" -f $bL, $dL, (255 * $dL))
    Write-Host ("  平均饱和   {0:N3}      Δ = {1:+0.000;-0.000;0.000}" -f $bS, $dS)
    Write-Host ("  近黑白图   {0}/{1}" -f $grayB, $pairs.Count)
    if (-not $NoDiff) { Write-Host ("  同像素差   平均 {0:N2}/255（越低越接近）" -f $avgDiff) }

    [void]$md.AppendLine("## B 组：``$B``")
    [void]$md.AppendLine('')
    [void]$md.AppendLine('| 指标 | A | B | Δ(B−A) |')
    [void]$md.AppendLine('|---|---|---|---|')
    [void]$md.AppendLine(("| luma709 | {0:N3} | {1:N3} | {2:+0.000;-0.000;0.000}（{3:N1} 点） |" -f $aL, $bL, $dL, (255 * $dL)))
    [void]$md.AppendLine(("| 通道均值 | {0:N1} | {1:N1} | {2:+0.0;-0.0;0.0} |" -f $aM, $bM, ($bM - $aM)))
    [void]$md.AppendLine(("| 平均饱和 | {0:N3} | {1:N3} | {2:+0.000;-0.000;0.000} |" -f $aS, $bS, $dS))
    [void]$md.AppendLine(("| 近黑白图 | {0} | {1} | — |" -f $grayA, $grayB))
    if (-not $NoDiff) { [void]$md.AppendLine(("| 同像素差 | — | — | 平均 {0:N2}/255 |" -f $avgDiff)) }
    [void]$md.AppendLine('')

    # 判读
    $verdict = @()
    $lumaPts = [Math]::Abs(255 * $dL)
    if ($lumaPts -ge 15) { $verdict += ("亮度差 {0:N0} 点（≥15）→ 差距明显；先怀疑标注口径/提示词，再怀疑参数" -f $lumaPts) }
    elseif ($lumaPts -ge 7) { $verdict += ("亮度差 {0:N0} 点（7–15）→ 有差距，方向对了但没到位" -f $lumaPts) }
    else { $verdict += ("亮度差 {0:N0} 点（<7）→ 明暗已经对上了" -f $lumaPts) }
    if ([Math]::Abs($dS) -ge 0.05) { $verdict += ("饱和差 {0:N3} → 色彩浓度明显不同（饱和对风格最敏感）" -f $dS) }
    if ($grayB -gt $pairs.Count * 0.3 -and $grayA -lt $pairs.Count * 0.3) { $verdict += 'B 组出现大量近黑白图 → 出图发黑白，检查负向词/权重/过训练' }
    if ($statA.Count -gt 0 -and $aS -lt 0.12) { $verdict += ("A 组本身平均饱和只有 {0:N3}（低饱和数据集）→ B 组发灰是「学到了」，不是 bug" -f $aS) }
    $verdict += '单 seed 只够看趋势；权重 0.6/0.8/1.0 三档都要出，别只测一档'
    Write-Host ''
    Write-Host '---------- 判读 ----------' -ForegroundColor Yellow
    foreach ($v in $verdict) { Write-Host ('  · ' + $v) }
    [void]$md.AppendLine('## 判读')
    [void]$md.AppendLine('')
    foreach ($v in $verdict) { [void]$md.AppendLine('- ' + $v) }
    [void]$md.AppendLine('')
    [void]$md.AppendLine('## 逐张明细')
    [void]$md.AppendLine('')
    [void]$md.AppendLine('| 文件 | luma709 A→B | 通道均值 A→B | 饱和 A→B | 像素差 |')
    [void]$md.AppendLine('|---|---|---|---|---|')
    foreach ($p in $pairs) {
        [void]$md.AppendLine(("| {0} | {1:N3}→{2:N3} | {3:N1}→{4:N1} | {5:N3}→{6:N3} | {7} |" -f
                $p.Name, $p.AL, $p.BL, $p.AM, $p.BM, $p.AS, $p.BS,
                $(if ($null -ne $p.Diff) { ('{0:N2}' -f $p.Diff) } else { '—' })))
    }
} else {
    [void]$md.AppendLine('## 逐张明细')
    [void]$md.AppendLine('')
    [void]$md.AppendLine('| 文件 | 尺寸 | luma709 | 通道均值 | 饱和 | 近灰像素% |')
    [void]$md.AppendLine('|---|---|---|---|---|---|')
    foreach ($x in $statA) {
        [void]$md.AppendLine(("| {0} | {1}×{2} | {3:N3} | {4:N1} | {5:N3} | {6:N1} |" -f
                $x.Name, $x.W, $x.H, $x.Luma709, $x.Mean255, $x.Sat, $x.GrayPct))
    }
}

# ---------------------------------------------------------------- 落盘
if (-not $Out) {
    $parent = Split-Path $A -Parent
    $Out = Join-Path $parent ('_风格对照_{0}.md' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
}
$outDir = Split-Path $Out -Parent
if ($outDir -and -not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
[System.IO.File]::WriteAllText($Out, $md.ToString(), $utf8NoBom)
Write-Host ''
Write-Host ("报告：{0}" -f $Out) -ForegroundColor Green
Write-Host ''
