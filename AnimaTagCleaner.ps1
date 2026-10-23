#requires -Version 5.1
<#
  AnimaTagCleaner v2.0.0 —— 画风 LoRA 打标清洗器
  ============================================================================
  和 v1.1 的根本区别：v1 只会删「垃圾 tag」，v2 先问一个 v1 没问的问题 ——
  「这条标注是在帮我学画风，还是在替模型解释画风？」

  依据《给本地扩散模型训练画师风格 LoRA 的实战经验》：
    §1 失败原因 = 标注替模型解释了画风（写了光照/明暗/色彩分级/渲染/流派）
    §2 缩放比 alpha/rank 决定「用不用得出」
    §3 标注语言必须与文本编码器主语言一致（跨语言实测 175 vs 226）
    §4 评测纪律：零权重对照、中立提示词、两套亮度口径
    §5 int8 评测够用

  v2 新增的五个能力：
    1. 画风泄漏词表（媒介/笔触、光照/明暗、色彩分级、风格流派、主观修饰）
    2. 数据集体检：覆盖度、语言、景别/姿势分布、互斥冲突、疑似角色名
    3. 图片色彩实测：直接回答「为什么出图发黑白/发灰」
    4. 训练缓存排雷：_anima_te.npz 不比对 caption，改了标注等于没改
    5. 生成中立评测提示词（不拿训练 caption 当测试提示词）

  用法：
      双击 一键清洗.bat                      交互：选目录 → 预演 → 确认
      命令行：
        powershell -NoProfile -ExecutionPolicy Bypass -File .\AnimaTagCleaner.ps1 `
            -Folder "D:\datasets\my_artist" -Trigger "@myartist"

  常用开关（完整列表见 docs\说明.md）：
      -Trigger <tag>      统一触发词（Anima 要求画师 tag 带 @），强制第一位
      -DryRun             只预演，绝不写盘
      -Yes                跳过确认
      -Recurse            含子目录
      -TrainConfig <toml> 顺手体检训练配置（alpha/rank 缩放、keep_tokens、缓存）
                          不给时会自动找 Next-Trainer 里同名数据集的配置，
                          找不到就设环境变量 NEXT_TRAINER_HOME=<Next-Trainer 目录>
      -CleanCache         自动挪走失效的 _anima_te.npz 文本编码器缓存
      -KeepLighting       保留光照/明暗 tag（不推荐）
      -KeepMedium         保留媒介/笔触/流派 tag（不推荐）
      -KeepColorTreatment 强制保留 monochrome/limited palette 这类 tag
      -DropCJK            删掉中文/日文 tag（编码器主语言为英文时用）
      -KeepSignature      保留 signature / watermark 这类签名 tag
      -DropAllText        连画面内文字 tag 一起删
      -KeepPose           不做姿势/景别互斥合并
      -MaxTags <n>        单图 tag 上限（默认不限）
#>
[CmdletBinding()]
param(
    [string]$Folder,
    [string]$Trigger = '',
    [string]$DropTags = '',
    [string]$Extensions = '.txt,.caption',
    [int]$MaxTags = 0,
    [switch]$Recurse,
    [switch]$DryRun,
    [switch]$Yes,
    [switch]$KeepLighting,
    [switch]$KeepMedium,
    [switch]$KeepSignature,
    [switch]$DropAllText,
    [switch]$KeepColorTreatment,
    [switch]$KeepPose,
    [switch]$KeepHairColor,
    [switch]$KeepEyeColor,
    [switch]$KeepParensEscape,
    [switch]$NoSort,
    [switch]$NoLowercase,
    [switch]$DropCJK,
    [switch]$CleanCache,
    [switch]$KeepCache,
    [switch]$SkipImageScan,
    [string]$ConfigDir,
    [string]$BackupRoot,
    [string]$TrainConfig
)

$ErrorActionPreference = 'Stop'
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$script:Warn = New-Object System.Collections.Generic.List[string]
$script:Crit = New-Object System.Collections.Generic.List[string]
$extList = @($Extensions -split ',' | ForEach-Object { $_.Trim().ToLowerInvariant() } | Where-Object { $_ })
$script:Interactive = ([Environment]::UserInteractive -and -not [Console]::IsInputRedirected)

# ============================================================================
# 0. 选文件夹 / 触发词
# ============================================================================
if (-not $Folder) {
    try {
        Add-Type -AssemblyName System.Windows.Forms | Out-Null
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = '选择要清洗标注的文件夹（图片 + txt 所在目录）'
        $dlg.ShowNewFolderButton = $false
        if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { Write-Host '已取消。' -ForegroundColor Yellow; exit 0 }
        $Folder = $dlg.SelectedPath
    } catch {
        Write-Host '无法弹出文件夹选择框，请用 -Folder 指定目录。' -ForegroundColor Red
        exit 1
    }
}
if (-not (Test-Path -LiteralPath $Folder)) { Write-Host "文件夹不存在：$Folder" -ForegroundColor Red; exit 1 }
$Folder = (Resolve-Path -LiteralPath $Folder).Path
$folderName = Split-Path $Folder -Leaf
if (-not $folderName) { $folderName = 'dataset' }

# 已有画师 tag 探查（帮用户决定触发词）
$existingArtists = @{}
if (-not $Trigger -and -not $Yes) {
    $probe = @(Get-ChildItem -LiteralPath $Folder -File -ErrorAction SilentlyContinue |
            Where-Object { $extList -contains $_.Extension.ToLowerInvariant() } | Select-Object -First 40)
    foreach ($pf in $probe) {
        $txt = ''
        try { $txt = [System.IO.File]::ReadAllText($pf.FullName, [System.Text.Encoding]::UTF8) } catch { }
        foreach ($m in [regex]::Matches($txt, '(?i)(^|,\s*)(@[^\s,]+)')) {
            $a = $m.Groups[2].Value
            if ($existingArtists.ContainsKey($a)) { $existingArtists[$a]++ } else { $existingArtists[$a] = 1 }
        }
    }
}

Write-Host ''
Write-Host '=============================================' -ForegroundColor Cyan
Write-Host '  Anima 打标清洗器  v2.0   (画风 LoRA 用)' -ForegroundColor Cyan
Write-Host '=============================================' -ForegroundColor Cyan
Write-Host "文件夹 : $Folder"
if ($existingArtists.Count -gt 0) {
    $top = $existingArtists.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 5
    Write-Host ('已有画师tag : ' + (($top | ForEach-Object { "$($_.Key)($($_.Value))" }) -join ', '))
}
if (-not $Trigger -and -not $Yes -and $script:Interactive) {
    Write-Host ''
    Write-Host '要统一成哪个画风触发词？（Anima 要求画师 tag 带 @；回车 = 不改动画师 tag）' -ForegroundColor Cyan
    $ans = Read-Host '触发词'
    $ans = if ($ans) { $ans.Trim() } else { '' }
    if ($ans) {
        if (-not $ans.StartsWith('@')) { $ans = '@' + $ans }
        $Trigger = $ans
    }
} elseif (-not $Trigger -and $existingArtists.Count -gt 0) {
    $Trigger = ($existingArtists.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1).Key
}
if ($Trigger) {
    $how = if ($script:Interactive -and -not $Yes) { '强制第一位' } else { '自动取出现最多的画师 tag，强制第一位' }
    Write-Host "触发词 : $Trigger  ($how)"
} else { Write-Host '触发词 : 未指定（不自动添加）' }
Write-Host ''

# ============================================================================
# 1. 规则表
# ============================================================================
function New-Set([string[]]$items) {
    $h = @{}
    foreach ($i in $items) { if ($i) { $h[$i.Trim().ToLowerInvariant()] = $true } }
    return $h
}

# ---- 1.1 元数据 / 平台 / 来源 / meme（不是画面元素）
$MetaList = @(
    'virtual youtuber', 'vtuber', 'original', 'commentary', 'commentary request', 'translation request',
    'partially translated', 'english commentary', 'tagme', 'artist request', 'character request', 'source request',
    'highres', 'absurdres', 'lowres', 'huge filesize', 'revision', 'edit', 'bad id', 'bad pixiv id', 'bad twitter id',
    'bad source', 'bad link', 'sample', 'sample watermark', 'third-party watermark', 'third-party edit',
    'ai-generated', 'ai assisted', 'ai-generated art', 'generated', 'novelty', 'name connection', 'refer to',
    'how to', 'web address', 'twitter username', 'username', 'artist name', 'character name', 'dated',
    'parody', 'meme', 'meme attire', 'crossover', 'alternate costume', 'alternate outfit', 'alternate hairstyle',
    'spoilers', 'commission', 'commissioned', 'request', 'repost', 'do not reupload', 'please do not reupload',
    'artist collaboration', 'cover', 'album cover', 'scan', 'magazine scan', 'border', 'black border',
    'white border', 'letterboxed', 'column lineup', 'reference sheet', 'character sheet', 'character profile',
    'variant set', 'set of 2', 'set of 3'
)
# ---- 1.2 版权 / 系列（画风 LoRA 不能把作品绑进触发词）
$SeriesList = @(
    'touhou', 'arknights', 'fate (series)', 'fate/grand order', 'blue archive', 'kantai collection', 'azur lane',
    'girls'' frontline', 'genshin impact', 'honkai impact 3rd', 'honkai: star rail', 'zenless zone zero',
    'pokemon', 'pokemon (series)', 'vocaloid', 'utau', 'umamusume', 'idolmaster', 'love live!', 'nijisanji',
    'hololive', 'project sekai', 'bang dream!', 'nikke', 'granblue fantasy', 'league of legends', 'overwatch',
    'princess connect!', 'mahjong soul', 'ensky', 'guilty gear', 'dragon ball', 'naruto', 'one piece', 'boruto',
    'bleach', 'jujutsu kaisen', 'chainsaw man', 'my hero academia', 'kimetsu no yaiba', 'frieren', 'oshi no ko',
    'neon genesis evangelion', 'touken ranbu', 'shadowverse', 'magia record', 'puella magi madoka magica',
    'girls band cry', 'lycoris recoil', 'bocchi the rock!', 'spy x family', 'shingeki no kyojin', 'dungeon meshi',
    'granblue', 'kancolle', 'fategrandorder', 'genshin', 'star rail'
)
# ---- 1.3 质量 / 评分 / 主观形容词（画风 LoRA 会被这类词绑死）
$QualityList = @(
    'masterpiece', 'best quality', 'good quality', 'normal quality', 'low quality', 'worst quality', 'high quality',
    'very aesthetic', 'aesthetic', 'the more the better', 'the_more_the_better', '8k', '4k', '2k', 'hd', 'uhd',
    'ultra detailed', 'highly detailed', 'detailed', 'extremely detailed', 'intricate', 'sharp focus',
    'high resolution', 'best', 'amazing', 'breathtaking', 'gorgeous', 'stunning', 'beautiful', 'pretty',
    'professional', 'award winning', 'masterful', 'top quality', 'premium', 'refined', 'polished', 'vibrant',
    'vivid', 'rich colors', 'bad anatomy', 'bad hands', 'extra digits', 'fewer digits', 'missing fingers',
    'mutated hands', 'bad proportions', 'poorly drawn', 'bad art', 'worst face', 'low saturation'
)
# ---- 1.4 打标伪影 / 相机后期（留着会把画风整体学糊）
$ArtifactList = @(
    'blurry', 'blurry background', 'blurry foreground', 'heavy blur', 'depth of field', 'bokeh', 'focus blur',
    'out of focus', 'jpeg artifacts', 'compression artifacts', 'motion blur', 'chromatic aberration', 'film grain',
    'noise', 'digital noise', 'vignetting', 'vignette', 'scanlines', 'dithering', 'pixelated', 'aliasing',
    'oversharpened', 'distorted', 'glitch', 'glitch art', 'light leak', 'double exposure', 'motion lines',
    'speed lines', 'action lines', 'impact lines', 'focus lines', 'bloom', 'lens flare', 'glow', 'glowing',
    'glowing light', 'light particles', 'light shafts', 'sparkle', 'sparkles', 'glitter', 'shiny'
)
# ---- 1.5 水印 / 签名（图上有，但学它只会画假签名）
$SigList = @(
    'watermark', 'signature', 'artist name', 'artist logo', 'twitter username', 'web address', 'logo', 'dated',
    'patreon logo', 'webtoon logo', 'studio logo', 'publisher logo', 'copyright notice', 'all rights reserved',
    'character name', 'third-party watermark', 'sample watermark', 'username', 'watermarked'
)
# ---- 1.6 画面内文字（笔记要求「画面内文字」算内容 → 默认保留）
$TextList = @(
    'text', 'english text', 'chinese text', 'japanese text', 'korean text', 'romaji text', 'speech bubble',
    'thought bubble', 'spoken heart', 'spoken question mark', 'text focus', 'sign', 'writing', 'letter'
)
# ---- 1.6b 表情符号噪声（纯符号的已被正则拦截，这里补几个带字母的）
$EmoteList = @(
    ':d', ':o', ':p', ':3', ':)', ';)', ':/', ':(', 'x3', 'uwu', 'owo', 'o o', 'a a'
)
# ---- 1.7 ★核心★ 画风泄漏词表
$LeakMedium = @(
    'watercolor (medium)', 'watercolor', 'oil painting (medium)', 'oil painting', 'acrylic (medium)', 'gouache',
    'ink', 'ink wash', 'sumi-e', 'pencil (medium)', 'colored pencil', 'graphite (medium)', 'charcoal (medium)',
    'pastel (medium)', 'crayon', 'chalk', 'marker (medium)', 'copic', 'digital media', 'traditional media',
    'mixed media', 'vector art', 'pixel art', '3d', '3d model', 'cgi', 'photo (medium)', 'photograph', 'photo',
    'realistic', 'photorealistic', 'semi-realistic', 'hyperrealistic', 'anime style', 'anime coloring', 'anime',
    'cartoon', 'comic', 'cel shading', 'cel shaded', 'flat color', 'flat colour', 'flat colours',
    'gradient shading', 'shading', 'screentone', 'halftone', 'hatching', 'cross-hatching', 'stippling',
    'airbrush', 'spray paint', 'paint splatter', 'impasto', 'lineart', 'line art', 'clean lineart',
    'thick lineart', 'thin lineart', 'colored lineart', 'sketch', 'rough sketch', 'unfinished', 'sketchbook',
    'concept art', 'illustration', 'official art', 'promotional art', 'cover art', 'poster (medium)', 'artbook'
)
$LeakLight = @(
    'lighting', 'backlighting', 'backlight', 'rim lighting', 'sidelighting', 'side lighting', 'volumetric lighting',
    'cinematic lighting', 'dramatic lighting', 'soft lighting', 'hard lighting', 'ambient lighting', 'ambient light',
    'ambient occlusion', 'dappled sunlight', 'god rays', 'light rays', 'light beam', 'sunbeam', 'spotlight',
    'chiaroscuro', 'high contrast', 'low contrast', 'contrast', 'shadow', 'cast shadow', 'drop shadow',
    'subsurface scattering', 'overexposure', 'underexposure', 'overexposed', 'underexposed', 'silhouette',
    'backlit', 'sunlit', 'moonlight', 'sunlight', 'moody lighting', 'natural lighting', 'studio lighting',
    'diffuse light', 'soft light', 'hard light', 'highlight', 'specular', 'rim light', 'darkness', 'brightness',
    'low-key', 'high-key', 'low key', 'high key', 'light and shadow'
)
# 彩色/黑白这一组由数据集实测决定（见 Get-DropReason）
$LeakColorConditional = @(
    'monochrome', 'greyscale', 'grayscale', 'greyscale with colored background',
    'grayscale with colored background', 'sepia', 'limited palette', 'spot color', 'duotone', 'selective color'
)
$LeakColor = @(
    'colorized', 'pastel colors', 'vibrant colors', 'muted colors', 'desaturated', 'high saturation', 'saturation',
    'color grading', 'color contrast', 'complementary colors', 'warm colors', 'cool colors', 'colorful',
    'gradient', 'pale color', 'flat tones', 'color palette', 'color scheme', 'monochromatic', 'pastel tone',
    'soft colors', 'vivid colors', 'neon colors', 'cool tone', 'warm tone'
)
$LeakStyle = @(
    'style', 'art style', 'in the style of', 'style of', 'impressionism', 'post-impressionism', 'pop art',
    'ukiyo-e', 'art nouveau', 'art deco', 'baroque', 'rococo', 'surrealism', 'expressionism', 'cubism',
    'pointillism', 'minimalism', 'abstract', 'geometric', 'psychedelic', 'retro', 'retro artstyle', 'vaporwave',
    'synthwave', 'y2k', '80s (style)', '90s (style)', '2000s (style)', '1970s (style)', '60s (style)',
    'aesthetic style', 'flat design', 'chibi', 'super deformed', 'anime screencap', 'fake screenshot',
    'screenshot', 'animated', 'animation', 'loop', 'gif'
)
$LeakSubjective = @(
    'detailed background', 'detailed eyes', 'detailed face', 'expressive', 'atmospheric', 'moody', 'elegant',
    'cute', 'pretty', 'lovely', 'gentle', 'cozy', 'dreamy', 'ethereal', 'cinematic', 'epic', 'dramatic',
    'mysterious', 'melancholic', 'nostalgic', 'wholesome', 'sleek', 'glossy', 'matte', 'smooth'
)
# ---- 1.7b 常用内容词（只用来判断「这个 tag 里有没有一个描述画面的词」，
#           用于把角色名/误标从待复核清单里挑出来，不参与删除）
$BodySceneWords = @(
    'hair', 'eyes', 'eye', 'ears', 'ear', 'tail', 'horns', 'horn', 'wings', 'wing', 'head', 'headwear', 'face',
    'cheek', 'cheeks', 'nose', 'mouth', 'lips', 'lip', 'tongue', 'teeth', 'tooth', 'fang', 'skin', 'neck',
    'shoulder', 'shoulders', 'collarbone', 'chest', 'breast', 'breasts', 'waist', 'hip', 'hips', 'navel',
    'stomach', 'back', 'thigh', 'thighs', 'knee', 'knees', 'leg', 'legs', 'foot', 'feet', 'toe', 'toes',
    'finger', 'fingers', 'hand', 'hands', 'arm', 'arms', 'elbow', 'wrist', 'groin', 'ass', 'butt', 'cleavage',
    'midriff', 'sideboob', 'underboob', 'nude', 'naked', 'topless', 'barefoot', 'bare', 'skin',
    'bangs', 'braid', 'braids', 'ponytail', 'twintails', 'ahoge', 'sidelocks', 'bun', 'buns', 'strand',
    'sitting', 'standing', 'lying', 'kneeling', 'squatting', 'crouching', 'walking', 'running', 'jumping',
    'holding', 'holding', 'looking', 'leaning', 'bent', 'reaching', 'hugging', 'grabbing', 'touching',
    'open', 'closed', 'crossed', 'spread', 'spread', 'raised', 'up', 'down', 'over', 'under', 'front', 'side',
    'behind', 'above', 'below', 'viewer', 'from', 'toward', 'towards', 'between', 'against',
    'girl', 'girls', 'boy', 'boys', 'solo', 'group', 'multiple', 'other', 'others', 'couple', 'duo',
    'background', 'indoors', 'outdoors', 'simple', 'flower', 'flowers', 'tree', 'trees', 'grass', 'plant',
    'plants', 'leaf', 'leaves', 'petal', 'rose', 'cherry', 'blossom', 'sky', 'cloud', 'clouds', 'sun', 'moon',
    'star', 'stars', 'night', 'day', 'sunset', 'sunrise', 'rain', 'snow', 'snowing', 'fog', 'wind',
    'water', 'ocean', 'sea', 'beach', 'river', 'lake', 'wave', 'waves', 'sand', 'ruins', 'mountain', 'forest',
    'field', 'garden', 'path', 'street', 'road', 'city', 'cityscape', 'building', 'buildings', 'architecture',
    'room', 'floor', 'wall', 'walls', 'ceiling', 'window', 'windows', 'door', 'curtain', 'curtains', 'bed',
    'pillow', 'blanket', 'chair', 'stool', 'table', 'desk', 'shelf', 'bookshelf', 'book', 'books', 'paper',
    'letter', 'newspaper', 'magazine', 'box', 'bag', 'backpack', 'basket', 'bottle', 'cup', 'glass', 'mug',
    'teacup', 'teapot', 'plate', 'saucer', 'bowl', 'food', 'cake', 'candy', 'drink', 'coffee', 'tea', 'water',
    'umbrella', 'lantern', 'candle', 'lamp', 'light', 'phone', 'camera', 'sword', 'knife', 'gun', 'weapon',
    'syringe', 'chain', 'rope', 'ribbon', 'string', 'balloon', 'doll', 'toy', 'plush', 'stuffed', 'animal',
    'cat', 'dog', 'bird', 'butterfly', 'fish', 'rabbit', 'mouse', 'bear', 'fox', 'wolf', 'dragon', 'horse',
    'sheep', 'cow', 'deer', 'squirrel', 'panda', 'penguin', 'witch', 'angel', 'devil', 'demon', 'ghost',
    'knight', 'maid', 'nurse', 'police', 'office', 'lady', 'school', 'student', 'teacher', 'idol', 'magic',
    'circle', 'pattern', 'scrunchie', 'bandaid', 'bandage', 'tattoo', 'tattoos', 'piercing', 'glasses',
    'sunglasses', 'eyewear', 'veil', 'mask', 'helmet', 'armor', 'harness', 'belt', 'suspenders', 'capelet',
    'cloak', 'shawl', 'towel', 'blanket', 'reflection', 'mirror', 'shadow', 'blood', 'sweat', 'tears',
    'smile', 'smiling', 'laughing', 'crying', 'blush', 'blushing', 'angry', 'sad', 'surprised', 'sleeping',
    'yawn', 'yawning', 'singing', 'dancing', 'eating', 'reading', 'writing', 'playing', 'holding',
    'chewing', 'gum', 'bubble', 'blowing', 'squirting', 'cum', 'nipples', 'pussy', 'sex', 'hetero',
    'yuri', 'kiss', 'hug', 'panties', 'bra', 'underwear', 'leotard', 'swimsuit', 'kimono', 'yukata',
    'uniform', 'serafuku', 'sailor', 'shirt', 'skirt', 'shorts', 'pants', 'dress', 'coat', 'jacket',
    'socks', 'thighhighs', 'pantyhose', 'stockings', 'boots', 'shoes', 'sandals', 'slippers', 'sneakers',
    'gloves', 'hat', 'cap', 'beret', 'beanie', 'hood', 'scarf', 'necktie', 'bowtie', 'bow', 'choker',
    'collar', 'neckerchief', 'apron', 'vest', 'sweater', 'cardigan', 'hoodie', 'blouse', 'camisole',
    'bodysuit', 'bikini', 'earrings', 'necklace', 'bracelet', 'ring', 'jewelry', 'nails', 'nail', 'makeup',
    'lipstick', 'eyeshadow', 'tiles', 'tile', 'onsen', 'bath', 'bathtub', 'shower', 'pool', 'fan', 'electric',
    'aircon', 'tv', 'monitor', 'screen', 'computer', 'game', 'keyboard', 'piano', 'violin', 'guitar',
    'racket', 'ball', 'tennis', 'basketball', 'sport', 'sports', 'sportswear', 'sneaker', 'mop', 'duster',
    'broom', 'bucket', 'cleaning', 'torii', 'shrine', 'temple', 'church', 'castle', 'bridge', 'fence',
    'railing', 'stair', 'stairs', 'roof', 'chimney', 'clock', 'alarm', 'mirror', 'sink', 'bathtub', 'moon',
    'crescent', 'full', 'half', 'new', 'shooting', 'galaxy', 'nebula', 'rainbow', 'aurora', 'lightning',
    'thunder', 'snowing', 'raining', 'breeze', 'summer', 'winter', 'spring', 'autumn', 'fall', 'festival',
    'hanabi', 'fireworks', 'parasol', 'bikini', 'sarong', 'scales', 'tentacles', 'feathers', 'halo',
    'pasties', 'keyhole', 'restrained', 'bound', 'leash', 'collar', 'garter', 'strap', 'straps', 'cuff',
    'cuffs', 'paw', 'paws', 'pose', 'poses',
    # 服装结构 / 细节（这些是内容，不是角色名）
    'sleeves', 'sleeve', 'frills', 'frill', 'frilled', 'footwear', 'hairband', 'headband', 'legwear',
    'thighhigh', 'thighhighs', 'cutout', 'cutouts', 'cape', 'capelet', 'cuff', 'trim', 'embroidery',
    'embroidered', 'lace', 'ruffle', 'ruffles', 'hem', 'neckline', 'lapel', 'pleat', 'pleats', 'pattern',
    'patterns', 'motif', 'motifs', 'ornament', 'ornaments', 'accessory', 'accessories', 'item', 'items',
    'prop', 'props', 'miniskirt', 'mini', 'crop', 'heels', 'heel', 'fishnet', 'fishnets', 'sole', 'soles',
    'forehead', 'tights', 'formal', 'habit', 'cross', 'heart', 'striped', 'plaid', 'polka', 'dot', 'dots',
    'stripe', 'stripes', 'hooded', 'shot', 'cowboy', 'full', 'close', 'upper', 'lower', 'body', 'roll',
    'rolled', 'puffy', 'wide', 'strapless', 'sleeveless', 'toeless', 'asymmetrical', 'center', 'torn',
    'single', 'detached', 'see', 'through', 'highleg', 'wariza', 'seiza', 'yokozuwari', 'hagoromo',
    'yin', 'yang', 'fours', 'cleft', 'venus', 'pubic', 'genital', 'genitals', 'crotch', 'anus',
    'classroom', 'class', 'chalkboard', 'blackboard', 'whiteboard', 'hallway', 'corridor',
    'clothes', 'clothing', 'outfit', 'costume', 'casual', 'sportswear', 'swimwear', 'lingerie',
    'bedroom', 'kitchen', 'bathroom', 'balcony', 'rooftop', 'cafe', 'restaurant', 'library', 'station',
    'heterochromia', 'multicolored', 'streaked', 'two-tone',
    # 补充常见内容词（压掉待复核清单的误报）
    'evening', 'twilight', 'morning', 'afternoon', 'watch', 'strawberry', 'military', 'instrument',
    'controller', 'pipe', 'painting', 'power', 'lines', 'facial', 'mark', 'public', 'nudity', 'censoring',
    'convenient', 'mechanical', 'parts', 'pole', 'horizon', 'mountainous', 'lolita', 'fashion', 'ship',
    'abyssal', 'tentacle', 'pit', 'buruma', 'uwabaki', 'easel', 'paintbrush', 'stove', 'petite', 'camouflage',
    'breath', 'pipe', 'smoking', 'utility', 'bar', 'censor', 'censored', 'mosaic', 'lens', 'signboard',
    'poster', 'banner', 'flag', 'paint', 'palette', 'canvas', 'frame', 'picture', 'photo', 'album', 'cd',
    'headphones', 'earphones', 'microphone', 'stereo', 'speaker', 'radio', 'television', 'laptop', 'tablet',
    'cable', 'wire', 'wires', 'plug', 'socket', 'switch', 'button', 'keyboard', 'mouse', 'console', 'joystick'
)

$SetMeta = New-Set $MetaList
$SetSeries = New-Set $SeriesList
$SetQuality = New-Set $QualityList
$SetArtifact = New-Set $ArtifactList
$SetSig = New-Set $SigList
$SetText = New-Set $TextList
$SetEmote = New-Set $EmoteList
$SetColorConditional = New-Set $LeakColorConditional
$SetLeak = @{}
foreach ($pair in @(
        @{ Cat = '媒介/笔触'; Bag = $LeakMedium },
        @{ Cat = '光照/明暗'; Bag = $LeakLight },
        @{ Cat = '色彩分级'; Bag = $LeakColor },
        @{ Cat = '风格流派'; Bag = $LeakStyle },
        @{ Cat = '主观修饰'; Bag = $LeakSubjective }
    )) {
    foreach ($w in $pair.Bag) { if (-not $SetLeak.ContainsKey($w)) { $SetLeak[$w] = $pair.Cat } }
}

# ---- 1.8 拼写/写法归一（口径统一）
$Synonym = @{
    'catgirl'         = 'cat girl'; '1girls' = '1girl'; '1boys' = '1boy'
    'blond hair'      = 'blonde hair'; 'gray hair' = 'grey hair'; 'gray eyes' = 'grey eyes'
    'gray background' = 'grey background'; 'gray' = 'grey'
    'thigh highs'     = 'thighhighs'; 'thighhigh' = 'thighhighs'; 'thigh-highs' = 'thighhighs'
    'hairband'        = 'hair band'; 'shoe' = 'shoes'; 'sock' = 'socks'
    'cellphone'       = 'phone'; 'smartphone' = 'phone'; 'mobile phone' = 'phone'
    'tee shirt'       = 'shirt'; 't-shirt' = 'shirt'
    'panty hose'      = 'pantyhose'
}
# ---- 1.9 同义组（保留最先出现 = tagger 置信度最高）
$SynonymGroups = @(
    @('stuffed toy', 'stuffed animal'), @('bikini', 'bikini top'), @('underwear', 'panties'),
    @('phone', 'smartphone'), @('kimono', 'japanese clothes')
)
# ---- 1.10 父子蕴含：有具体就删通用的（同一件事说两遍）
$Implies = @(
    @{ P = 'animal ears'; C = @('cat ears', 'dog ears', 'fox ears', 'wolf ears', 'rabbit ears', 'mouse ears', 'bear ears', 'cow ears', 'horse ears', 'tiger ears', 'leopard ears', 'panda ears', 'sheep ears', 'deer ears', 'pig ears', 'lion ears', 'fake animal ears', 'animal ear fluff', 'cat hat', 'animal hat', 'cat hood', 'animal hood') },
    @{ P = 'animal hat'; C = @('cat hat', 'dog hat', 'fox hat', 'bunny hat', 'rabbit hat', 'bear hat', 'panda hat', 'cow hat', 'cat hood') },
    @{ P = 'animal hood'; C = @('cat hood', 'dog hood', 'fox hood', 'bunny hood', 'bear hood') },
    @{ P = 'tail'; C = @('cat tail', 'dog tail', 'fox tail', 'cow tail', 'horse tail', 'wolf tail', 'dragon tail', 'demon tail', 'devil tail', 'rabbit tail', 'bear tail', 'tiger tail', 'raccoon tail', 'stuffed tail', 'lion tail') },
    @{ P = 'very long hair'; C = @('absurdly long hair') },
    @{ P = 'sleeves past wrists'; C = @('sleeves past fingers') },
    @{ P = 'short sleeves'; C = @('puffy short sleeves', 'frilled short sleeves') },
    @{ P = 'long sleeves'; C = @('puffy long sleeves', 'frilled long sleeves', 'sleeves rolled up') },
    @{ P = 'puffy sleeves'; C = @('puffy long sleeves', 'puffy short sleeves') },
    @{ P = 'wide sleeves'; C = @('fur-trimmed sleeves') },
    @{ P = 'sleeveless'; C = @('sleeveless dress', 'sleeveless shirt', 'sleeveless top', 'sleeveless sweater', 'sleeveless jacket', 'sleeveless kimono') },
    @{ P = 'strapless'; C = @('strapless dress', 'strapless leotard', 'strapless bra', 'strapless top') },
    @{ P = 'hood'; C = @('hood up', 'hood down') },
    @{ P = 'bra'; C = @('bow bra', 'sports bra', 'strapless bra', 'lace bra', 'front-tie bra', 'cupless bra') },
    @{ P = 'fang out'; C = @('fang') },
    @{ P = 'tongue'; C = @('tongue out') },
    @{ P = 'heart'; C = @('heart balloon', 'heart-shaped box', 'heart hair ornament', 'heart background') },
    @{ P = 'bell'; C = @('jingle bell', 'cowbell') },
    @{ P = 'sky'; C = @('blue sky', 'night sky', 'cloudy sky', 'gradient sky', 'orange sky', 'red sky', 'purple sky', 'dark sky', 'starry sky') },
    @{ P = 'starry sky'; C = @('night sky') },
    @{ P = 'striped'; C = @('striped thighhighs', 'striped socks', 'striped shirt', 'striped dress', 'striped skirt', 'striped shorts', 'striped pants', 'striped bikini', 'striped bow') },
    @{ P = 'breasts'; C = @('small breasts', 'medium breasts', 'large breasts', 'huge breasts') },
    @{ P = 'hair ornament'; C = @('hairclip', 'hair bow', 'hair bobbles', 'hair band', 'hair ribbon', 'hair flower', 'hair stick', 'hair beads', 'hair rings', 'hairpin') },
    @{ P = 'bow'; C = @('hair bow', 'tail bow', 'hat bow', 'bow panties', 'bow bra', 'bowtie') },
    @{ P = 'ribbon'; C = @('hair ribbon', 'waist ribbon') },
    @{ P = 'hat'; C = @('cat hat', 'animal hat', 'beanie', 'beret', 'witch hat', 'conical hat', 'santa hat', 'top hat', 'baseball cap', 'bonnet', 'cabbie hat', 'peaked cap', 'visor cap') },
    @{ P = 'headwear'; C = @('hat', 'cat hat', 'animal hat', 'beanie', 'beret', 'witch hat', 'conical hat', 'santa hat', 'baseball cap', 'bonnet', 'helmet', 'headband', 'hood up', 'hood down', 'cabbie hat', 'peaked cap', 'visor cap', 'nurse cap', 'maid headdress') },
    @{ P = 'footwear'; C = @('sneakers', 'boots', 'slippers', 'sandals', 'high heels', 'loafers', 'mary janes', 'ballet shoes') },
    @{ P = 'shoes'; C = @('sneakers', 'boots', 'slippers', 'sandals', 'high heels', 'loafers', 'mary janes', 'ballet shoes', 'no shoes') },
    @{ P = 'socks'; C = @('striped socks', 'patterned socks', 'frilled socks', 'rainbow socks') },
    @{ P = 'thighhighs'; C = @('striped thighhighs', 'patterned thighhighs', 'frilled thighhighs', 'fishnet thighhighs') },
    @{ P = 'pantyhose'; C = @('patterned pantyhose', 'thighband pantyhose') },
    @{ P = 'gloves'; C = @('fingerless gloves', 'elbow gloves', 'paw gloves', 'frilled gloves', 'animal hands') },
    @{ P = 'underwear'; C = @('panties', 'bra', 'bloomers', 'boxers', 'fundoshi') },
    @{ P = 'panties'; C = @('bow panties', 'patterned panties', 'side-tie panties', 'frilled panties', 'bikini panties') },
    @{ P = 'swimsuit'; C = @('bikini', 'one-piece swimsuit', 'school swimsuit', 'competition swimsuit') },
    @{ P = 'wings'; C = @('demon wings', 'angel wings', 'feathered wings', 'butterfly wings') },
    @{ P = 'horns'; C = @('demon horns', 'dragon horns', 'curled horns') },
    @{ P = 'box'; C = @('gift box', 'cardboard box', 'heart-shaped box') },
    @{ P = 'open clothes'; C = @('open jacket', 'open shirt', 'open dress', 'open coat', 'open vest') },
    @{ P = 'gradient'; C = @('gradient background', 'gradient hair', 'gradient eyes') },
    @{ P = 'balloon'; C = @('heart balloon', 'star balloon') },
    @{ P = 'antlers'; C = @('reindeer antlers', 'fake antlers') },
    @{ P = 'clock'; C = @('alarm clock', 'wall clock') },
    @{ P = 'nail polish'; C = @('black nails', 'white nails', 'red nails', 'blue nails', 'green nails', 'pink nails', 'purple nails', 'yellow nails', 'grey nails', 'brown nails', 'orange nails', 'multicolored nails', 'french nails') },
    @{ P = 'jewelry'; C = @('earrings', 'ring', 'bracelet', 'necklace', 'pendant', 'brooch') },
    @{ P = 'uniform'; C = @('school uniform', 'serafuku', 'military uniform') },
    @{ P = 'school uniform'; C = @('serafuku', 'blazer uniform') },
    @{ P = 'multicolored hair'; C = @('streaked hair', 'two-tone hair', 'gradient hair', 'colored inner hair') },
    @{ P = 'shirt'; C = @('dress shirt', 'collared shirt', 't-shirt', 'naked shirt', 'sleeveless shirt', 'turtleneck') },
    @{ P = 'dress'; C = @('sundress', 'wedding dress', 'maid dress', 'china dress', 'santa dress', 'frilled dress', 'fur-trimmed dress', 'sleeveless dress', 'strapless dress', 'halter dress', 'backless dress', 'evening gown') },
    @{ P = 'leotard'; C = @('strapless leotard', 'black leotard', 'playboy bunny') },
    @{ P = 'shorts'; C = @('short shorts', 'denim shorts', 'plaid shorts') },
    @{ P = 'sitting'; C = @('wariza', 'seiza', 'yokozuwari', 'indian style', 'on bed') },
    @{ P = 'lying'; C = @('on back', 'on stomach', 'on side') },
    @{ P = 'hair between eyes'; C = @('hair over one eye') },
    @{ P = 'looking at viewer'; C = @('facing viewer') }
)
# ---- 1.11 服装/配饰前缀规则：<颜色/修饰语> <名词> 存在时，裸名词删除
$GarmentNouns = @(
    'shirt', 'dress', 'skirt', 'jacket', 'hoodie', 'sweater', 'cardigan', 'coat', 'vest', 'apron', 'robe', 'cloak',
    'cape', 'bikini', 'swimsuit', 'leotard', 'camisole', 'bra', 'panties', 'socks', 'thighhighs', 'pantyhose',
    'gloves', 'shoes', 'boots', 'footwear', 'hat', 'headwear', 'beanie', 'bow', 'ribbon', 'choker', 'scarf',
    'necktie', 'bowtie', 'collar', 'sailor collar', 'nails', 'uniform', 'bloomers', 'bandeau', 'bustier', 'shorts',
    'pants', 'tank top', 'crop top', 'hairband', 'hair band', 'slippers', 'sneakers', 'pullover', 'blouse',
    'stockings', 'garter belt', 'bag', 'armor', 'eyewear', 'heels', 'sandals', 'legwear', 'headband', 'hairclip',
    'belt'
)
$Modifiers = @(
    'white', 'black', 'red', 'blue', 'grey', 'gray', 'green', 'yellow', 'pink', 'purple', 'brown', 'orange', 'aqua',
    'cyan', 'silver', 'gold', 'blonde', 'rainbow', 'multicolored', 'dark', 'light', 'pale', 'bright', 'navy',
    'beige', 'teal', 'short', 'long', 'puffy', 'frilled', 'pleated', 'striped', 'patterned', 'plaid', 'denim',
    'leather', 'open', 'oversized', 'hooded', 'detached', 'wide', 'bare', 'sleeveless', 'strapless', 'halter',
    'backless', 'fur-trimmed', 'see-through', 'translucent', 'cropped', 'naked', 'heart-shaped', 'side-tie',
    'school', 'competition', 'one-piece', 'reindeer', 'fake', 'cat', 'animal', 'frilly', 'layered', 'asymmetrical',
    'torn', 'loose', 'tight', 'high-waist', 'high', 'low', 'no', 'single', 'thigh-high', 'fishnet', 'round',
    'black-framed', 'polka dot', 'floral', 'lace', 'satin', 'silk', 'velvet', 'peaked', 'visor', 'cabbie',
    'nurse', 'maid'
)
$ColorWords = @(
    'white', 'black', 'red', 'blue', 'grey', 'gray', 'green', 'yellow', 'pink', 'purple', 'brown', 'orange',
    'aqua', 'cyan', 'silver', 'gold', 'navy', 'beige', 'teal', 'blonde'
)
$ColorConflictNouns = @(
    'shirt', 'dress', 'skirt', 'jacket', 'hoodie', 'sweater', 'cardigan', 'coat', 'vest', 'apron', 'robe', 'cloak',
    'cape', 'bikini', 'swimsuit', 'leotard', 'camisole', 'bra', 'panties', 'socks', 'thighhighs', 'pantyhose',
    'gloves', 'shoes', 'boots', 'hat', 'beanie', 'bloomers', 'shorts', 'pants', 'uniform', 'scarf', 'necktie',
    'nails', 'sailor collar', 'bandeau', 'bustier', 'pullover', 'blouse', 'stockings', 'slippers', 'eyewear',
    'headwear', 'hairband', 'hair band'
)
# 多角色/多视图：这些图里「同属性不同颜色」是正常的
$MultiCharTags = @(
    '2girls', '3girls', '4girls', '6+girls', 'multiple girls', '2boys', '3boys', 'multiple boys', 'group',
    'multiple views', '1boy', '1other', '2others', 'hetero'
)
# ---- 1.12 互斥组（同一张图只能有一个；保留最先出现 = 置信度最高）
$ConflictGroups = @(
    @{ Name = '人数'; Re = '^(1girl|1boy|1other|2girls|2boys|3girls|3boys|4girls|4boys)$'; MultiChar = $false },
    @{ Name = '发色'; Re = '^(black|brown|blonde|red|blue|green|purple|pink|white|grey|silver|aqua|orange) hair$'; MultiChar = $true; Skip = $KeepHairColor },
    @{ Name = '瞳色'; Re = '^(black|brown|blonde|red|blue|green|purple|pink|yellow|grey|silver|aqua|orange) eyes$'; MultiChar = $true; Skip = $KeepEyeColor },
    @{ Name = '背景色'; Re = '^(white|black|grey|blue|red|green|yellow|pink|purple|brown|orange) background$'; MultiChar = $false },
    @{ Name = '胸围'; Re = '^(small breasts|medium breasts|large breasts|huge breasts|flat chest)$'; MultiChar = $false },
    @{ Name = '姿势'; Re = '^(standing|sitting|lying|all fours|kneeling|squatting|crouching|handstand)$'; MultiChar = $true; Skip = $KeepPose },
    @{ Name = '景别'; Re = '^(portrait|upper body|cowboy shot|full body|close-up)$'; MultiChar = $true; Skip = $KeepPose }
)
# ---- 1.13 排序（Anima 官方顺序：触发词 → 计数 → 安全 → 人物 → 服装 → 姿势 → 场景 → 构图）
$SortRules = @(
    @{ R = 0; P = '^@' },
    @{ R = 1; P = '^(1girl|1boy|1other|2girls|2boys|3girls|3boys|4girls|4boys|multiple girls|multiple boys|solo|group|6\+girls)$' },
    @{ R = 2; P = '^(safe|sensitive|questionable|nsfw|explicit)$' },
    @{ R = 3; P = 'hair|bangs|ahoge|twintails|ponytail|braid|sidelocks|drill|hime cut|bun' },
    @{ R = 4; P = 'eyes|eyebrow|eyelashes|pupil|heterochromia|mole|freckles' },
    @{ R = 5; P = 'ears|tail|horns|wings|^animal |cat girl|dog girl|fox girl|demon|pointy ears|facial mark|^cow$|^cat$' },
    @{ R = 6; P = 'shirt|jacket|hoodie|sweater|cardigan|coat|vest|apron|robe|cloak|cape|dress|uniform|serafuku|sailor|bikini|swimsuit|leotard|bodysuit|camisole|tank top|crop top|bra|necktie|bowtie|collar|scarf|kimono|yukata|bandeau|bustier|naked|sleeves|hood' },
    @{ R = 7; P = 'skirt|shorts|pants|jeans|trousers|leggings|clothes pull' },
    @{ R = 8; P = 'socks|thighhighs|pantyhose|footwear|shoes|boots|slippers|sneakers|sandals|heels|barefoot|bare legs|no shoes|legwear' },
    @{ R = 9; P = 'panties|underwear|bloomers|garter|thigh strap|nipples|navel|cleavage|groin' },
    @{ R = 10; P = 'hair ornament|hairclip|hair bow|hair band|hairband|bobbles|ribbon|^bow|bow$|choker|jewelry|earrings|ring|bracelet|necklace|pendant|nails|brooch|hat|headwear|beanie|beret|gloves|wrist cuffs|ofuda|umbrella|bag|backpack|sack|gift box|box|plate|bottle|book|pillow|stuffed toy|stuffed cat|stuffed bunny|balloon|candy|lollipop|popsicle|food|drink|cake|fur trim|fur-trimmed' },
    @{ R = 11; P = 'sitting|standing|lying|wariza|seiza|all fours|knees up|arms up|hands up|hand up|holding|spread legs|crossed legs|leaning|outstretched|steepled|pose|walking|running|squatting|on back|on side|on stomach|front view|crossed arms|arm up' },
    @{ R = 12; P = 'smile|open mouth|closed mouth|parted lips|blush|^:|^;|^\^|tongue|teeth|expression|looking at viewer|facing viewer|one eye closed|embarrassed|tears|sweatdrop|spoken' },
    @{ R = 13; P = 'background|outdoors|indoors|sky|water|beach|^day$|^night$|cloud|ocean|horizon|forest|room|bed|pillow|floor|window|city|street|snow|rain|sand|railing' },
    @{ R = 14; P = 'upper body|full body|portrait|cowboy shot|feet out of frame|close-up|from side|from behind|from above|dutch angle|simple background' }
)

# ============================================================================
# 2. 自定义词表
# ============================================================================
$cfgDir = if ($ConfigDir) { $ConfigDir } else { if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path } }
function Read-WordList([string]$path) {
    if (-not (Test-Path -LiteralPath $path)) { return @() }
    return @(Get-Content -LiteralPath $path -Encoding UTF8 | ForEach-Object {
            $l = $_.Trim()
            if ($l -and -not $l.StartsWith('#')) { ($l -replace '\s*#.*$', '').Trim().TrimEnd(',').ToLowerInvariant() }
        } | Where-Object { $_ })
}
$extraRemove = @(Read-WordList (Join-Path $cfgDir '自定义_额外删除.txt'))
$forceKeep = @(Read-WordList (Join-Path $cfgDir '自定义_强制保留.txt'))
$leakExtra = @(Read-WordList (Join-Path $cfgDir '自定义_画风词.txt'))
if ($DropTags) { $extraRemove += @($DropTags -split ',' | ForEach-Object { $_.Trim().ToLowerInvariant() } | Where-Object { $_ }) }
foreach ($w in $leakExtra) { if (-not $SetLeak.ContainsKey($w)) { $SetLeak[$w] = '自定义画风词' } }
$script:ForceKeep = New-Set $forceKeep
$script:SetExtra = New-Set $extraRemove

# 常用内容词集合（只在待复核清单里判「这条 tag 像不像名字」，不参与删除）
$contentRaw = @($BodySceneWords) + @($ColorWords) + @($GarmentNouns -split ' ') + @($Modifiers -split ' ')
$ContentWords = New-Set @($contentRaw | Where-Object { $_ -and $_ -notmatch ' ' })

$modAlt = (($Modifiers | Sort-Object Length -Descending | ForEach-Object { [regex]::Escape($_) }) -join '|')
$modifierNounRes = @()
foreach ($n in $GarmentNouns) {
    $modifierNounRes += [pscustomobject]@{ Noun = $n; Re = ('^(?:' + $modAlt + ') ' + [regex]::Escape($n) + '$') }
}
$colorAlt = (($ColorWords | ForEach-Object { [regex]::Escape($_) }) -join '|')
$colorNounRes = @()
foreach ($n in $ColorConflictNouns) {
    $colorNounRes += [pscustomobject]@{ Noun = $n; Re = ('^(?:' + $colorAlt + ') ' + [regex]::Escape($n) + '$') }
}

# ============================================================================
# 3. 工具函数
# ============================================================================
function Split-Tags([string]$text) {
    if ([string]::IsNullOrWhiteSpace($text)) { return @() }
    $t = $text -replace "`r`n", ',' -replace "`n", ',' -replace "`r", ','
    $t = $t -replace [char]0xFF0C, ','                     # 全角逗号
    $t = $t -replace [char]0x3001, ','                     # 顿号
    $t = $t -replace '\s*\|\s*', ','                       # 竖线分隔
    return @($t -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function Normalize-Tag([string]$tag) {
    $t = $tag.Trim()
    if (-not $t) { return '' }
    try { $t = $t.Normalize([System.Text.NormalizationForm]::FormKC) } catch { }
    $t = $t -replace '\s+', ' '
    # 去掉 A1111 / kohya 加权语法：(tag:1.2) ((tag)) [tag] {tag} <lora:x:1>
    $t = $t -replace '<[^>]*>', ''
    if ($t -match '^\(+\s*(.+?)\s*:\s*-?[\d.]+\s*\)+$') { $t = $Matches[1] }
    elseif ($t -match '^\(+\s*(.+?)\s*\)+$') { $t = $Matches[1] }
    elseif ($t -match '^\[+\s*(.+?)\s*\]+$') { $t = $Matches[1] }
    elseif ($t -match '^\{+\s*(.+?)\s*\}+$') { $t = $Matches[1] }
    elseif ($t -match '^(.+?)\s*:\s*-?[\d.]+$' -and $t -notmatch '^[:;^<>|]') { $t = $Matches[1] }
    $t = $t -replace '^\[+|\]+$', ''
    $t = $t -replace '^BREAK$', ''
    if (-not $KeepParensEscape) { $t = $t -replace '\\([()])', '$1' }
    $t = $t -replace '\\', ''
    # 下划线归一：只对含字母/数字的 tag 做（^_^ 这类表情不拆）
    if ($t -notmatch '^@' -and $t -match '[A-Za-z0-9]' -and $t -notmatch '^[\W_]+$') { $t = $t -replace '_', ' ' }
    $t = $t -replace '\s+', ' '
    $t = $t.Trim(' ', ',', '.', ';', '"', "'", '。', '、')
    if (-not $t) { return '' }
    if (-not $NoLowercase -and -not $t.StartsWith('@')) { $t = $t.ToLowerInvariant() }
    $k = $t.ToLowerInvariant()
    if ($Synonym.ContainsKey($k)) { $t = $Synonym[$k] }
    return $t
}

function Get-SortRank([string]$tag) {
    foreach ($rule in $SortRules) { if ($tag -match $rule.P) { return $rule.R } }
    return 99
}

function Add-Reason($list, [string]$tag, [string]$why) {
    $list.Add([pscustomobject]@{ Tag = $tag; Why = $why })
}

function Test-CJK([string]$t) { return ($t -match '[\u3040-\u30FF\u3400-\u4DBF\u4E00-\u9FFF\uF900-\uFAFF\uAC00-\uD7AF]') }

# 返回删除原因（$null = 保留）
function Get-DropReason([string]$t, [double]$grayRatio) {
    $lt = $t.ToLowerInvariant()
    if ($script:ForceKeep.ContainsKey($lt)) { return $null }
    if ($script:SetExtra.ContainsKey($lt)) { return '自定义删除' }
    if ($SetMeta.ContainsKey($lt)) { return '元数据' }
    if ($SetSeries.ContainsKey($lt) -or $lt -match '\(series\)$' -or $lt -match '\(copyright\)$') { return '版权/系列' }
    if ($SetQuality.ContainsKey($lt)) { return '质量/评分' }
    if ($SetArtifact.ContainsKey($lt)) { return '打标伪影' }
    # score_9 在下划线归一后变成 "score 9"，两种写法都要抓
    if ($lt -match '^score[ _]?\d+(?:[ _]up)?$') { return '质量/评分' }
    if ($SetEmote.ContainsKey($lt)) { return '噪声符号' }
    if (-not $KeepSignature -and $SetSig.ContainsKey($lt)) { return '签名/水印' }
    if ($DropAllText -and $SetText.ContainsKey($lt)) { return '画面文字' }
    if ($SetLeak.ContainsKey($lt)) {
        $cat = $SetLeak[$lt]
        if ($KeepMedium -and $cat -eq '媒介/笔触') { return $null }
        if ($KeepLighting -and $cat -eq '光照/明暗') { return $null }
        return ('画风泄漏·' + $cat)
    }
    # 彩色/黑白这类词：数据集里有近黑白图时保留（它是区分信息，不是泄漏）
    if ($SetColorConditional.ContainsKey($lt)) {
        if ($KeepColorTreatment -or $grayRatio -ge 0.05) { return $null }
        return '画风泄漏·色彩分级'
    }
    if ($DropCJK -and (Test-CJK $t)) { return '非英文' }
    if ($t -match '^[^\p{L}\p{N}]+$') { return '噪声符号' }
    if ($t.Length -eq 1 -and $t -notmatch '[\u4E00-\u9FFF]') { return '噪声符号' }
    if ($lt -match '^(?:unknown|none|null|n/?a|undefined|blank|empty)$') { return '噪声符号' }
    return $null
}

# ============================================================================
# 4. 单条 caption 清洗
# ============================================================================
function Clean-Caption([string]$raw, [double]$grayRatio) {
    $removed = New-Object System.Collections.Generic.List[object]
    $notes = New-Object System.Collections.Generic.List[string]
    $tags = @(Split-Tags $raw | ForEach-Object { Normalize-Tag $_ } | Where-Object { $_ })
    $origCount = $tags.Count

    # 4.1 分类删除
    $step = New-Object System.Collections.Generic.List[string]
    foreach ($t in $tags) {
        $why = Get-DropReason $t $grayRatio
        if ($why) { Add-Reason $removed $t $why; continue }
        $step.Add($t)
    }
    $tags = @($step)

    # 4.2 精确去重（转义反斜杠 / 大小写视为同一个）
    $seen = @{}
    $uniq = New-Object System.Collections.Generic.List[string]
    foreach ($t in $tags) {
        $k = ($t.ToLowerInvariant() -replace '\\', '')
        if ($seen.ContainsKey($k)) { Add-Reason $removed $t '重复'; continue }
        $seen[$k] = $true; $uniq.Add($t)
    }
    $tags = @($uniq)

    # 4.3 同义组
    foreach ($grp in $SynonymGroups) {
        $hit = @($tags | Where-Object { $grp -contains $_.ToLowerInvariant() })
        if ($hit.Count -gt 1) {
            $keep = $hit[0]
            foreach ($h in @($hit | Select-Object -Skip 1)) {
                $tags = @($tags | Where-Object { $_ -ne $h })
                Add-Reason $removed $h ('同义，保留 ' + $keep)
            }
        }
    }

    # 4.4 父子蕴含 / 修饰语前缀
    $lower = @($tags | ForEach-Object { $_.ToLowerInvariant() })
    $multiChar = @($lower | Where-Object { $MultiCharTags -contains $_ }).Count -gt 0
    $drop = @{}
    foreach ($rule in $Implies) {
        if ($script:ForceKeep.ContainsKey($rule.P)) { continue }
        if ($lower -notcontains $rule.P) { continue }
        foreach ($c in $rule.C) { if ($lower -contains $c) { $drop[$rule.P] = $c; break } }
    }
    foreach ($mn in $modifierNounRes) {
        if ($lower -notcontains $mn.Noun) { continue }
        if ($script:ForceKeep.ContainsKey($mn.Noun)) { continue }
        $child = @($lower | Where-Object { $_ -match $mn.Re }) | Select-Object -First 1
        if ($child) { $drop[$mn.Noun] = $child }
    }
    if ($lower -contains 'holding' -and -not $script:ForceKeep.ContainsKey('holding')) {
        $child = @($lower | Where-Object { $_ -match '^holding \S' }) | Select-Object -First 1
        if ($child) { $drop['holding'] = $child }
    }
    foreach ($k in @($drop.Keys)) {
        $tags = @($tags | Where-Object { $_.ToLowerInvariant() -ne $k })
        Add-Reason $removed $k ('已被 ' + $drop[$k] + ' 包含')
    }

    # 4.5 互斥合并
    $lower = @($tags | ForEach-Object { $_.ToLowerInvariant() })
    foreach ($cg in $ConflictGroups) {
        if ($cg.Skip) { continue }
        if ($cg.MultiChar -and $multiChar) { continue }
        if ($cg.Name -eq '瞳色' -and ($lower -contains 'heterochromia')) {
            $notes.Add('heterochromia：跳过瞳色合并'); continue
        }
        $hit = @($tags | Where-Object { $_.ToLowerInvariant() -match $cg.Re })
        if ($hit.Count -gt 1) {
            $keep = $hit[0]
            $others = @($hit | Select-Object -Skip 1)
            foreach ($o in $others) {
                $tags = @($tags | Where-Object { $_ -ne $o })
                Add-Reason $removed $o ('互斥(' + $cg.Name + ')，保留 ' + $keep)
            }
            $notes.Add(('{0} 互斥：保留 {1}，删 {2}' -f $cg.Name, $keep, ($others -join ', ')))
        }
        # 人数改完后重算「多角色」标记：1boy 被删掉后，发色/瞳色就该合并了
        if ($cg.Name -eq '人数') {
            $lower = @($tags | ForEach-Object { $_.ToLowerInvariant() })
            $multiChar = @($lower | Where-Object { $MultiCharTags -contains $_ }).Count -gt 0
        }
    }
    # 4.5b 同一件衣服标了两种颜色（多角色图跳过）
    if (-not $multiChar) {
        foreach ($cn in $colorNounRes) {
            $hit = @($tags | Where-Object { $_.ToLowerInvariant() -match $cn.Re })
            if ($hit.Count -gt 1) {
                $keep = $hit[0]
                foreach ($o in @($hit | Select-Object -Skip 1)) {
                    $tags = @($tags | Where-Object { $_ -ne $o })
                    Add-Reason $removed $o ('颜色冲突(' + $cn.Noun + ')，保留 ' + $keep)
                }
                $notes.Add(('{0} 颜色冲突：保留 {1}' -f $cn.Noun, $keep))
            }
        }
    }

    # 4.6 触发词钉在第一位
    if ($Trigger) {
        $tg = $Trigger.Trim()
        $others = @($tags | Where-Object { $_.StartsWith('@') -and $_.ToLowerInvariant() -ne $tg.ToLowerInvariant() })
        foreach ($o in $others) {
            $tags = @($tags | Where-Object { $_ -ne $o })
            Add-Reason $removed $o '画师tag被触发词替换'
        }
        $tags = @($tags | Where-Object { $_.ToLowerInvariant() -ne $tg.ToLowerInvariant() })
        $tags = @($tg) + $tags
        if ($others.Count -gt 0) { $notes.Add('原有画师 tag 已替换为 ' + $tg) }
    }

    # 4.7 排序
    if (-not $NoSort -and $tags.Count -gt 1) {
        $idx = for ($i = 0; $i -lt $tags.Count; $i++) {
            [pscustomobject]@{ T = $tags[$i]; R = (Get-SortRank $tags[$i]); I = $i }
        }
        $tags = @($idx | Sort-Object R, I | ForEach-Object { $_.T })
    }
    # 4.8 上限（触发词不计）
    $keepMax = $MaxTags + $(if ($Trigger) { 1 } else { 0 })
    if ($MaxTags -gt 0 -and $tags.Count -gt $keepMax) {
        $cutTags = @($tags | Select-Object -Last ($tags.Count - $keepMax))
        $tags = @($tags | Select-Object -First $keepMax)
        foreach ($c in $cutTags) { Add-Reason $removed $c '超出上限截断' }
    }

    $cjk = @($tags | Where-Object { Test-CJK $_ })
    return [pscustomobject]@{
        Text      = ($tags -join ', ')
        Tags      = $tags
        Removed   = $removed
        Notes     = $notes
        OrigCount = $origCount
        NewCount  = $tags.Count
        CJK       = $cjk
    }
}

# ============================================================================
# 5. 图片色彩实测（回答「为什么出图发黑白/发灰」）
# ============================================================================
function Get-ImageColorStat([string]$path) {
    try {
        $bmp = [System.Drawing.Bitmap]::FromFile($path)
        $w = 64
        $h = [int]($bmp.Height * 64 / $bmp.Width)
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
        $bmp.Dispose()
        return [pscustomobject]@{
            Luma709 = [Math]::Round($sumL / $n, 4)
            Mean255 = [Math]::Round(255.0 * $sumM / $n, 1)
            Sat     = [Math]::Round($sumS / $n, 4)
            GrayPct = [Math]::Round(100.0 * $grey / $n, 1)
        }
    } catch { return $null }
}

# ============================================================================
# 6. 收集文件
# ============================================================================
$gciArgs = @{ LiteralPath = $Folder; File = $true }
if ($Recurse) { $gciArgs['Recurse'] = $true }
$allFiles = @(Get-ChildItem @gciArgs -ErrorAction SilentlyContinue)
$capFiles = @($allFiles | Where-Object { $extList -contains $_.Extension.ToLowerInvariant() } | Sort-Object FullName)
if ($capFiles.Count -eq 0) {
    Write-Host ("没找到标注文件（{0}）。" -f ($extList -join ' / ')) -ForegroundColor Yellow
    Write-Host '如果还没打标：先用 tagger 出一遍 txt，再回来清洗。' -ForegroundColor Yellow
    exit 1
}
$imgExt = @('.png', '.jpg', '.jpeg', '.webp', '.bmp', '.gif', '.jfif', '.avif')
$imgFiles = @($allFiles | Where-Object { $imgExt -contains $_.Extension.ToLowerInvariant() })
$jsonCaps = @($allFiles | Where-Object { $_.Extension.ToLowerInvariant() -eq '.json' })
Write-Host ("标注文件 : {0} 个    同目录图片 : {1} 个" -f $capFiles.Count, $imgFiles.Count)

$capStems = @{}
foreach ($f in $capFiles) { $capStems[(Split-Path $f.FullName -Parent) + '|' + $f.BaseName] = $true }
$imgStems = @{}
foreach ($f in $imgFiles) { $imgStems[(Split-Path $f.FullName -Parent) + '|' + $f.BaseName] = $true }
$missingImg = @($capStems.Keys | Where-Object { -not $imgStems.ContainsKey($_) } | ForEach-Object { ($_ -split '\|')[-1] })
$missingCap = @($imgStems.Keys | Where-Object { -not $capStems.ContainsKey($_) } | ForEach-Object { ($_ -split '\|')[-1] })
if ($imgFiles.Count -eq 0) { $script:Warn.Add('这个目录里没有图片，请确认没选错文件夹。') }
if ($missingImg.Count -gt 0) { $script:Warn.Add(('{0} 个标注没有同名图片（前几个：{1}）' -f $missingImg.Count, (($missingImg | Select-Object -First 5) -join ', '))) }
if ($missingCap.Count -gt 0) { $script:Warn.Add(('{0} 张图片没有标注（前几个：{1}）—— 没标注的图会被当成空 caption' -f $missingCap.Count, (($missingCap | Select-Object -First 5) -join ', '))) }
if ($jsonCaps.Count -gt 0) { $script:Crit.Add(('检测到 {0} 个 .json 标注。训练配置 prefer_json_caption=true 时会优先吃 json，改 txt 不会生效。' -f $jsonCaps.Count)) }

# ============================================================================
# 7. 图片色彩实测（先做，因为彩色/黑白会影响规则）
# ============================================================================
$imgStats = @()
$grayRatio = 0.0
$imgMeanL = 0.0
$imgMeanS = 0.0
$imgMean255 = 0.0
$grayImages = @()
$lowSatImages = @()
if (-not $SkipImageScan -and $imgFiles.Count -gt 0) {
    $drawOk = $true
    try { Add-Type -AssemblyName System.Drawing -ErrorAction Stop | Out-Null } catch { $drawOk = $false }
    if (-not $drawOk) { $script:Warn.Add('无法加载 System.Drawing，跳过图片色彩实测。') }
    else {
        Write-Host ''
        Write-Host '正在实测图片色彩（这条会直接回答「为什么出图发黑白/发灰」）...' -ForegroundColor DarkGray
        $maxScan = 160
        $step = [Math]::Max(1, [int]($imgFiles.Count / $maxScan))
        for ($i = 0; $i -lt $imgFiles.Count; $i += $step) {
            $st = Get-ImageColorStat $imgFiles[$i].FullName
            if ($st) {
                $st | Add-Member -NotePropertyName Name -NotePropertyValue $imgFiles[$i].Name -Force
                $imgStats += $st
                if ($st.Sat -lt 0.08) { $grayImages += $imgFiles[$i].Name }
                elseif ($st.Sat -lt 0.15) { $lowSatImages += $imgFiles[$i].Name }
            }
        }
        if ($imgStats.Count -gt 0) {
            $imgMeanL = ($imgStats | Measure-Object Luma709 -Average).Average
            $imgMeanS = ($imgStats | Measure-Object Sat -Average).Average
            $imgMean255 = ($imgStats | Measure-Object Mean255 -Average).Average
            $grayRatio = $grayImages.Count / $imgStats.Count
        }
    }
}

# ============================================================================
# 8. 逐文件清洗（先算，不写）
# ============================================================================
$results = New-Object System.Collections.Generic.List[object]
$skipped = New-Object System.Collections.Generic.List[object]
$tagFreq = @{}
$afterFreq = @{}
$reasonFreq = @{}
$leakFreq = @{}
$encFixed = 0
$cjkTotal = 0

foreach ($f in $capFiles) {
    $raw = ''
    try { $raw = [System.IO.File]::ReadAllText($f.FullName, [System.Text.Encoding]::UTF8) } catch { $raw = '' }
    if ($raw -match [char]0xFFFD) {
        try {
            $gbk = [System.Text.Encoding]::GetEncoding(936)
            $raw2 = [System.IO.File]::ReadAllText($f.FullName, $gbk)
            if ($raw2 -notmatch [char]0xFFFD) {
                $raw = $raw2; $encFixed++
                $script:Warn.Add(('{0} 按 GBK 解码（文件不是 UTF-8）' -f $f.Name))
            }
        } catch { }
    }
    if ([string]::IsNullOrWhiteSpace($raw)) { $skipped.Add([pscustomobject]@{ File = $f.Name; Why = '空文件' }); continue }
    $r = Clean-Caption $raw $grayRatio
    if ($r.Tags.Count -eq 0) { $skipped.Add([pscustomobject]@{ File = $f.Name; Why = '清洗后为空' }); continue }
    $results.Add([pscustomobject]@{
            Path    = $f.FullName
            Name    = $f.Name
            OldText = $raw.Trim()
            NewText = $r.Text
            Removed = $r.Removed
            Notes   = $r.Notes
            OldN    = $r.OrigCount
            NewN    = $r.NewCount
            Changed = ($raw.Trim() -ne $r.Text)
            CJK     = $r.CJK.Count
        })
    foreach ($x in $r.Removed) {
        $k = $x.Tag.ToLowerInvariant()
        if ($tagFreq.ContainsKey($k)) { $tagFreq[$k]++ } else { $tagFreq[$k] = 1 }
        if ($reasonFreq.ContainsKey($x.Why)) { $reasonFreq[$x.Why]++ } else { $reasonFreq[$x.Why] = 1 }
        if ($x.Why.StartsWith('画风泄漏')) {
            if ($leakFreq.ContainsKey($k)) { $leakFreq[$k]++ } else { $leakFreq[$k] = 1 }
        }
    }
    foreach ($t in $r.Tags) {
        $k = $t.ToLowerInvariant()
        if ($afterFreq.ContainsKey($k)) { $afterFreq[$k]++ } else { $afterFreq[$k] = 1 }
    }
    $cjkTotal += $r.CJK.Count
}

if ($results.Count -eq 0) {
    Write-Host ''
    Write-Host '所有标注都被跳过（空文件 / 清洗后为空）。检查一下 tagger 是否真的写进了 txt。' -ForegroundColor Red
    exit 1
}

$changed = @($results | Where-Object { $_.Changed })
$removedTotal = ($results | ForEach-Object { $_.Removed.Count } | Measure-Object -Sum).Sum
if (-not $removedTotal) { $removedTotal = 0 }
$oldSum = ($results | ForEach-Object { $_.OldN } | Measure-Object -Sum).Sum
$newSum = ($results | ForEach-Object { $_.NewN } | Measure-Object -Sum).Sum
$fileCount = [Math]::Max(1, $results.Count)

# ============================================================================
# 9. 数据集体检
# ============================================================================
$cov = @()
foreach ($kv in $afterFreq.GetEnumerator()) {
    $cov += [pscustomobject]@{ Tag = $kv.Key; N = $kv.Value; Pct = [Math]::Round(100.0 * $kv.Value / $fileCount, 1) }
}
$cov = @($cov | Sort-Object N -Descending)
$onceTags = @($cov | Where-Object { $_.N -le 1 })
$allTags = @($cov | Where-Object { $_.Pct -ge 99 })

function Get-Coverage([string[]]$patterns) {
    $hit = 0
    $details = @{}
    foreach ($r in $results) {
        $lower = @($r.NewText -split ',' | ForEach-Object { $_.Trim().ToLowerInvariant() })
        $m = @($lower | Where-Object { $tag = $_; @($patterns | Where-Object { $tag -match $_ }).Count -gt 0 })
        if ($m.Count -gt 0) {
            $hit++
            foreach ($x in $m) { if ($details.ContainsKey($x)) { $details[$x]++ } else { $details[$x] = 1 } }
        }
    }
    return [pscustomobject]@{ Files = $hit; Details = $details }
}
$framing = Get-Coverage @('^(portrait|upper body|cowboy shot|full body|close-up|feet out of frame)$')
$pose = Get-Coverage @('^(standing|sitting|lying|all fours|kneeling|squatting|crouching|wariza|seiza|yokozuwari|on back|on stomach|on side|legs up|knees up|bent over|leaning forward)$')
$expr = Get-Coverage @('^(smile|open mouth|closed mouth|blush|looking at viewer|parted lips|fang|tears|embarrassed)$')
$bg = Get-Coverage @('(background|indoors|outdoors|sky|room|city|beach|forest|classroom|street)$')

$counts = @($results | ForEach-Object { $_.NewN })
$avgN = if ($counts.Count) { ($counts | Measure-Object -Average).Average } else { 0 }
$stdN = 0.0
if ($counts.Count -gt 1) {
    $sd = ($counts | ForEach-Object { [Math]::Pow($_ - $avgN, 2) } | Measure-Object -Sum).Sum / ($counts.Count - 1)
    $stdN = [Math]::Sqrt($sd)
}

# 标注形态检测：tag 列表 vs 自然语言。如果是句子，tag 级规则（画风泄漏词表）会失灵
$segTotal = 0
$wordTotal = 0
foreach ($r in $results) {
    foreach ($t in ($r.NewText -split ',')) {
        $segTotal++
        $wordTotal += @(($t.Trim() -split '\s+') | Where-Object { $_ }).Count
    }
}
$avgWords = if ($segTotal) { [Math]::Round($wordTotal / $segTotal, 2) } else { 0 }
$isProse = ($avgWords -ge 4)
if ($isProse) {
    $script:Crit.Add(('标注看起来是自然语言句子（平均每段 {0} 个词）。tag 级清洗规则（画风泄漏词表）基本不会命中——要么改用 tagger 出 tag，要么把句子里的光照/渲染描述手工删掉。' -f $avgWords))
}

# ============================================================================
# 10. 训练配置体检（可选）
# ============================================================================
$cfg = $null
# 没指定配置时，自动找一份同名数据集的主训练配置。
# 查找顺序：环境变量 NEXT_TRAINER_HOME / NEXT_TRAINER_CONFIG → 数据集父目录 → 脚本所在目录。
if (-not $TrainConfig) {
    $cfgRoots = @()
    if ($env:NEXT_TRAINER_HOME) { $cfgRoots += (Join-Path $env:NEXT_TRAINER_HOME 'config') }
    if ($env:NEXT_TRAINER_CONFIG) { $cfgRoots += $env:NEXT_TRAINER_CONFIG }
    $cfgRoots += (Split-Path $Folder -Parent)
    if ($PSScriptRoot) { $cfgRoots += $PSScriptRoot }
    foreach ($root in ($cfgRoots | Select-Object -Unique)) {
        if (-not $root -or -not (Test-Path -LiteralPath $root)) { continue }
        $hit = @(Get-ChildItem -LiteralPath $root -Filter '*.toml' -File -ErrorAction SilentlyContinue |
                Where-Object {
                    $b = $_.BaseName.ToLowerInvariant()
                    $b.StartsWith($folderName.ToLowerInvariant()) -and $b -notlike '*dataset*' -and
                    $b -notlike '*sd-scripts*' -and $b -notlike '*smoke*'
                } | Sort-Object LastWriteTime -Descending)
        if ($hit.Count -ge 1) {
            $TrainConfig = $hit[0].FullName
            Write-Host ("自动带上训练配置：{0}" -f (Split-Path $TrainConfig -Leaf)) -ForegroundColor DarkGray
            break
        }
    }
}
if ($TrainConfig) {
    if (Test-Path -LiteralPath $TrainConfig) {
        $c = @{}
        foreach ($line in (Get-Content -LiteralPath $TrainConfig -Encoding UTF8)) {
            if ($line -match '^\s*([A-Za-z0-9_]+)\s*=\s*(.+?)\s*(?:#.*)?$') {
                $c[$Matches[1]] = $Matches[2].Trim([char[]]@(' ', '"', "'"))
            }
        }
        $cfg = $c
    } else {
        $script:Warn.Add("找不到训练配置：$TrainConfig")
    }
}

# ============================================================================
# 11. 控制台汇总
# ============================================================================
Write-Host ''
Write-Host '---------- 结果预演 ----------' -ForegroundColor Cyan
Write-Host ("要处理 : {0} 个文件    需要改动 : {1} 个    跳过 : {2} 个" -f $results.Count, $changed.Count, $skipped.Count)
Write-Host ("tag 总数 : {0} -> {1}   删除 {2} 个    平均 {3:N1} 个/张" -f $oldSum, $newSum, $removedTotal, $avgN)
if ($reasonFreq.Count -gt 0) {
    Write-Host ''
    Write-Host '删除分类：' -ForegroundColor Cyan
    $reasonFreq.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 12 | ForEach-Object {
        Write-Host ("    {0,-24} {1}" -f $_.Key, $_.Value)
    }
} else {
    Write-Host ''
    Write-Host '没有可删的东西——这批标注已经洗过了（或本来就很干净）。' -ForegroundColor DarkGray
}
if ($leakFreq.Count -gt 0) {
    Write-Host ''
    Write-Host '★ 画风泄漏词（这些原本在替模型解释画风）：' -ForegroundColor Yellow
    $leakFreq.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 15 | ForEach-Object {
        Write-Host ("    {0,-34} {1}" -f $_.Key, $_.Value) -ForegroundColor Yellow
    }
}
if ($imgStats.Count -gt 0) {
    Write-Host ''
    Write-Host '★ 数据集色彩实测' -ForegroundColor Yellow
    Write-Host ("    平均亮度 luma709 = {0:N3}（通道均值 {1:N1}/255）  平均饱和 = {2:N3}" -f $imgMeanL, $imgMean255, $imgMeanS)
    Write-Host ("    近黑白图（饱和<0.08）: {0}/{1}   低饱和图(0.08~0.15): {2}" -f $grayImages.Count, $imgStats.Count, $lowSatImages.Count)
    if ($grayImages.Count -gt 0) {
        Write-Host ('    近黑白图：' + (($grayImages | Select-Object -First 8) -join ', ') + $(if ($grayImages.Count -gt 8) { ' …' } else { '' })) -ForegroundColor DarkYellow
    }
    if ($grayRatio -ge 0.05) {
        Write-Host ('    → 有 {0:N0}% 近黑白图，monochrome/greyscale 这类 tag 予以保留（它们让模型区分彩色/黑白，删了会混成一锅）' -f (100 * $grayRatio)) -ForegroundColor Yellow
    }
}
if ($skipped.Count -gt 0) {
    Write-Host ''
    Write-Host '跳过：' -ForegroundColor Yellow
    $skipped | ForEach-Object { Write-Host ("    {0}  ({1})" -f $_.File, $_.Why) }
}
foreach ($w in $script:Warn) { Write-Host ("[注意] " + $w) -ForegroundColor Yellow }
foreach ($w in $script:Crit) { Write-Host ("[严重] " + $w) -ForegroundColor Red }

# ============================================================================
# 12. 输出目录
# ============================================================================
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
if (-not $BackupRoot) { $BackupRoot = Join-Path (Split-Path $Folder -Parent) '_tag备份' }
$outDir = Join-Path $BackupRoot ("{0}_{1}" -f $folderName, $stamp)
$reportPath = Join-Path $outDir '清洗报告.md'
$checkPath = Join-Path $outDir '数据集体检.md'
$reviewPath = Join-Path $outDir '待复核.txt'
$promptPath = Join-Path $outDir '评测提示词.txt'
$specPath = Join-Path $outDir '标注规范.md'
$fullPath = Join-Path $outDir 'tag全量表.md'

# ============================================================================
# 13. 清洗报告
# ============================================================================
$sb = New-Object System.Text.StringBuilder
function AddW([string]$s) { [void]$sb.AppendLine($s) }

AddW ("# {0} 打标清洗报告 {1}" -f $folderName, $stamp)
AddW ''
AddW ("- 目录：``$Folder``")
AddW ("- 触发词：$(if ($Trigger) { '`' + $Trigger + '`（强制第一位）' } else { '未指定' })")
AddW ("- 文件：{0} 个，改动 {1} 个，跳过 {2} 个" -f $results.Count, $changed.Count, $skipped.Count)
AddW ("- tag：{0} -> {1}（删除 {2}，平均 {3:N1} 个/张）" -f $oldSum, $newSum, $removedTotal, $avgN)
AddW ("- 模式：$(if ($DryRun) { '预演（未写盘）' } else { '已应用' })")
AddW ''
AddW '## 一、删除分类'
AddW ''
AddW '| 原因 | 次数 | 说明 |'
AddW '|---|---|---|'
$reasonDesc = @{
    '元数据'             = '平台/来源/meme，不是画面元素'
    '版权/系列'           = '作品名，会把画风绑到角色上'
    '质量/评分'           = '留质量词只会把画风和质量绑死'
    '打标伪影'            = '模糊/景深/噪点等 tagger 副产品'
    '签名/水印'           = '学它只会画假签名'
    '画面文字'            = '按 -DropAllText 删除'
    '自定义删除'           = '来自 自定义_额外删除.txt'
    '重复'                = '同图内重复'
    '噪声符号'            = '纯符号/单字符垃圾'
    '非英文'              = '按 -DropCJK 删除'
    '超出上限截断'          = '按 -MaxTags 截断'
    '画师tag被触发词替换'    = '统一画师 tag'
}
$reasonFreq.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
    $d = ''
    if ($reasonDesc.ContainsKey($_.Key)) { $d = $reasonDesc[$_.Key] }
    elseif ($_.Key.StartsWith('已被')) { $d = '具体写法已经说清楚了，通用词多余' }
    elseif ($_.Key.StartsWith('互斥')) { $d = '同图互斥，保留第一个（tagger 置信度最高）' }
    elseif ($_.Key.StartsWith('颜色冲突')) { $d = '同件衣物标了两种颜色，多角色图不合并' }
    elseif ($_.Key.StartsWith('同义')) { $d = '同义写法只留一个' }
    elseif ($_.Key.StartsWith('画风泄漏')) { $d = '★ 在替模型解释画风，见下一节' }
    AddW ("| {0} | {1} | {2} |" -f $_.Key, $_.Value, $d)
}
AddW ''
AddW '## 二、画风泄漏词（v2 核心）'
AddW ''
if ($leakFreq.Count -eq 0) {
    AddW '这批标注里没有检测到画风泄漏词。'
    AddW ''
} else {
    AddW '依据笔记 §1：**把光照、明暗、色彩分级、渲染方式写进 caption，等于告诉模型「这些不用你学」。** 下列词已删除：'
    AddW ''
    AddW '| tag | 次数 |'
    AddW '|---|---|'
    $leakFreq.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object { AddW ("| ``{0}`` | {1} |" -f $_.Key, $_.Value) }
    AddW ''
}
AddW '## 三、逐文件明细'
AddW ''
AddW '| 文件 | tag | 删除 | 备注 |'
AddW '|---|---|---|---|'
foreach ($r in $results) {
    $rem = (@($r.Removed | ForEach-Object {
                if ($_.Why -match '泄漏|元数据|质量|伪影|签名|版权') { "**$($_.Tag)**($($_.Why))" } else { "$($_.Tag)($($_.Why))" }
            }) -join ', ')
    if ($rem.Length -gt 400) { $rem = $rem.Substring(0, 400) + ' …' }
    $nt = (@($r.Notes) -join '；')
    AddW ("| {0} | {1}→{2} | {3} | {4} |" -f $r.Name, $r.OldN, $r.NewN, $rem, $nt)
}
if ($skipped.Count -gt 0) {
    AddW ''
    AddW '## 四、跳过'
    AddW ''
    $skipped | ForEach-Object { AddW ("- ``{0}``：{1}" -f $_.File, $_.Why) }
}
AddW ''

# ============================================================================
# 14. 数据集体检
# ============================================================================
$cb = New-Object System.Text.StringBuilder
function AddC([string]$s) { [void]$cb.AppendLine($s) }
function Fmt-Details($d) {
    return (($d.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 6 |
            ForEach-Object { "$($_.Key)×$($_.Value)" }) -join ', ')
}
$conflictFiles = @($results | Where-Object { @($_.Notes | Where-Object { $_ -match '互斥' }).Count -gt 0 })
$cjkFiles = @($results | Where-Object { $_.CJK -gt 0 })
$npzFiles = @($allFiles | Where-Object { $_.Name -like '*_anima_te.npz' })
$latentNpz = @($allFiles | Where-Object { $_.Name -like '*_anima.npz' -and $_.Name -notlike '*_te.npz' })

AddC ("# {0} 数据集体检 {1}" -f $folderName, $stamp)
AddC ''
AddC '> 先量化，再归因。这一页的数字比参数值钱。'
AddC ''
AddC '## 1. 规模'
AddC ''
AddC ("- 标注 {0} 个 / 图片 {1} 张" -f $capFiles.Count, $imgFiles.Count)
AddC ("- 每张 tag 数：平均 {0:N1}，最少 {1}，最多 {2}，标准差 {3:N1}" -f $avgN,
    $(if ($counts.Count) { ($counts | Measure-Object -Minimum).Minimum } else { 0 }),
    $(if ($counts.Count) { ($counts | Measure-Object -Maximum).Maximum } else { 0 }), $stdN)
AddC ''
AddC '## 2. 色彩实测（回答「为什么出图发黑白/发灰」）'
AddC ''
if ($imgStats.Count -eq 0) {
    AddC '- 未做图片实测（-SkipImageScan 或加载失败）。'
} else {
    AddC ("- 采样 {0} 张：平均亮度 **luma709 = {1:N3}**（通道均值 {2:N1}/255，两个口径不能混用）" -f $imgStats.Count, $imgMeanL, $imgMean255)
    AddC ("- 平均饱和 **{0:N3}**（0–1；<0.10 视觉上接近灰）" -f $imgMeanS)
    AddC ("- 近黑白图（每图平均饱和<0.08）：**{0} 张 / {1}（{2:N0}%）**" -f $grayImages.Count, $imgStats.Count, (100 * $grayRatio))
    AddC ("- 低饱和图（0.08–0.15）：{0} 张" -f $lowSatImages.Count)
    AddC ''
    AddC '| 文件 | luma709 | 通道均值 | 饱和 | 近灰像素% |'
    AddC '|---|---|---|---|---|'
    $imgStats | Sort-Object Sat | Select-Object -First 10 | ForEach-Object {
        AddC ("| {0} | {1:N3} | {2:N1} | {3:N3} | {4:N1} |" -f $_.Name, $_.Luma709, $_.Mean255, $_.Sat, $_.GrayPct)
    }
}
AddC ''
AddC '## 3. 覆盖度'
AddC ''
AddC ("- 全数据集都有的 tag（{0} 个，对画风无区分信息但无害）：{1}" -f $allTags.Count, (($allTags | Select-Object -First 12 | ForEach-Object { "``$($_.Tag)``" }) -join ', '))
AddC ("- 只出现 1 次：{0} 个（全量见 tag全量表.md，可疑项见 待复核.txt）" -f $onceTags.Count)
AddC ("- 不同 tag 总数：{0}" -f $cov.Count)
AddC ''
AddC '## 4. 景别 / 姿势 / 表情 / 背景覆盖（「换个动作就崩」的根源）'
AddC ''
AddC '| 类别 | 有标注的文件数 | 覆盖 | 分布 |'
AddC '|---|---|---|---|'
AddC ("| 景别 | {0} | {1:N0}% | {2} |" -f $framing.Files, (100.0 * $framing.Files / $fileCount), (Fmt-Details $framing.Details))
AddC ("| 姿势 | {0} | {1:N0}% | {2} |" -f $pose.Files, (100.0 * $pose.Files / $fileCount), (Fmt-Details $pose.Details))
AddC ("| 表情 | {0} | {1:N0}% | {2} |" -f $expr.Files, (100.0 * $expr.Files / $fileCount), (Fmt-Details $expr.Details))
AddC ("| 背景 | {0} | {1:N0}% | {2} |" -f $bg.Files, (100.0 * $bg.Files / $fileCount), (Fmt-Details $bg.Details))
AddC ''
AddC '## 5. 语言'
AddC ''
AddC ("- 含中文/日文/韩文的 tag：共 {0} 个（涉及 {1} 个文件）" -f $cjkTotal, $cjkFiles.Count)
AddC ("- 标注形态：平均每段 {0} 个词 → {1}" -f $avgWords, $(if ($isProse) { '**自然语言句子**（tag 级规则会失灵，见下方警告）' } else { 'tag 列表 ✓' }))
if ($cjkTotal -gt 0) {
    AddC '- **笔记 §3：标注语言必须与文本编码器主语言一致，跨语言实测同一份权重 175 vs 226（等于没挂 LoRA）。**'
    AddC '- 要么全部译成英文，要么统一中文，不要混用。要直接删掉非英文 tag 就加 `-DropCJK`。'
}
AddC ''
AddC '## 6. 疑似问题'
AddC ''
$issues = @()
if ($grayRatio -ge 0.15) { $issues += ("{0:N0}% 的图是近黑白 → 出图偏黑白不是标注问题，是数据集本身；想要彩色必须混入彩色图或加负向词" -f (100 * $grayRatio)) }
if ($imgMeanS -lt 0.15 -and $imgStats.Count -gt 0) { $issues += ("数据集平均饱和只有 {0:N3}，属于低饱和 → 出图发灰是「学到了」的结果，不是 bug" -f $imgMeanS) }
if ($pose.Files -lt 0.6 * $fileCount) { $issues += ("只有 {0:N0}% 的文件标了姿势 → 模型学不到「文字↔动作」的对应，换动作容易出错" -f (100.0 * $pose.Files / $fileCount)) }
if ($framing.Details.Count -gt 0) {
    $topFraming = ($framing.Details.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1)
    if ($topFraming.Value -ge 0.8 * $fileCount) {
        $issues += ("景别高度单一（{0} 占 {1}/{2}）→ 画风会被绑在这个景别上，要全身/复杂构图时容易崩" -f $topFraming.Key, $topFraming.Value, $fileCount)
    }
}
if ($conflictFiles.Count -gt 0) { $issues += ("{0} 个文件里出现互斥 tag（同一张图标了站/坐、单/双人等）→ 已按「保留第一个」处理，但根子在 tagger，值得回头看这几张图" -f $conflictFiles.Count) }
if ($issues.Count -eq 0) { $issues += '没有触发自动检查项。' }
foreach ($i in $issues) { AddC ('- ' + $i) }
AddC ''
AddC '## 7. 训练缓存排雷（最容易忽略，也最致命）'
AddC ''
if ($npzFiles.Count -gt 0) {
    AddC ("- 目录里有 **{0} 个 ``*_anima_te.npz`` 文本编码器缓存**" -f $npzFiles.Count)
    AddC '- 已核对源码 `sd-scripts/library/strategy_anima.py:171` 的 `is_disk_cached_outputs_expected()`：它只检查 npz 里有没有 `prompt_embeds`/`attn_mask`/`t5_input_ids` 等键，**完全不比对 caption。**'
    AddC '- 也就是说：**你改了 txt，训练器仍然用旧 caption 的 embedding，清洗等于没做。**'
    AddC '- 处理：把这些 `_anima_te.npz` 删掉或挪走再训练（清洗器会问你要不要自动挪走）。'
    AddC ("- 另有 {0} 个 ``*_anima.npz`` 潜变量缓存，只跟图片/分辨率有关，不用动。" -f $latentNpz.Count)
} else {
    AddC '- 没发现文本编码器缓存，改了 txt 会正常生效。'
}
AddC ''
if ($cfg) {
    AddC '## 8. 训练配置体检'
    AddC ''
    $dim = 0; $alpha = 0
    if ($cfg.ContainsKey('network_dim')) { [void][int]::TryParse($cfg['network_dim'], [ref]$dim) }
    if ($cfg.ContainsKey('network_alpha')) { [void][int]::TryParse($cfg['network_alpha'], [ref]$alpha) }
    if ($dim -gt 0) {
        $scale = if ($alpha -gt 0) { $alpha / $dim } else { 1.0 }
        AddC ("- network_dim = **{0}**，network_alpha = **{1}** → 实际缩放 **alpha/rank = {2:N2}**" -f $dim, $alpha, $scale)
        if ($scale -lt 1.0) {
            AddC '  - **笔记 §2.2：缩放 0.5 等于把学到的增量砍半，出图只能靠把权重往上拉来补偿；权重一拉高，构图和动作就开始崩。**'
            AddC '  - 建议 `network_alpha = network_dim × 1.5~2`（如 dim 32 / alpha 64），这样出图权重 0.6–0.8 就有效果。'
            AddC '  - 想更浓优先动 alpha（一个标量，不占显存），不要动 dim（翻倍会显著涨显存）。'
        } else {
            AddC '  - 缩放正常（≥1.0），出图不需要靠拉高权重找风格。'
        }
    }
    if ($cfg.ContainsKey('keep_tokens')) {
        AddC ("- keep_tokens = **{0}**{1}" -f $cfg['keep_tokens'], $(if ($cfg['keep_tokens'] -eq '1') { '（触发词被钉在第一位 ✓）' } else { '（应为 1，否则触发词可能被 shuffle 打散）' }))
    }
    if ($cfg.ContainsKey('max_train_epochs')) {
        $ep = 0; [void][int]::TryParse($cfg['max_train_epochs'], [ref]$ep)
        if ($ep -gt 0) {
            AddC ("- max_train_epochs = **{0}** → 约 {1:N0} 步（{2} 张 × {0} 轮）" -f $ep, ($ep * $fileCount), $fileCount)
            if ($ep -ge 30 -and $fileCount -lt 60) {
                AddC '  - 小数据集跑 30+ 轮属于过训练区：实测后期会脸崩、底色发灰。建议每 4 轮存档，取 **60%–100%** 之间的档。'
            }
        }
    }
    if ($cfg.ContainsKey('cache_text_encoder_outputs') -and $cfg['cache_text_encoder_outputs'] -eq 'true' -and $npzFiles.Count -gt 0) {
        AddC '- `cache_text_encoder_outputs = true` 且目录里已有 `_anima_te.npz` → 见第 7 节，必须先清缓存。'
    }
    AddC ''
}

# ============================================================================
# 15. 待复核 / 评测提示词 / 标注规范
# ============================================================================
function Test-NoContentWord([string]$tag) {
    $ws = @($tag -split '[^a-z0-9]+' | Where-Object { $_.Length -ge 3 })
    if ($ws.Count -eq 0) { return $false }
    foreach ($w in $ws) { if ($ContentWords.ContainsKey($w)) { return $false } }
    return $true
}
# 已经在规则表里的 tag（画风词/质量/元数据等）不再进"疑似名字"清单
$KnownRuleTags = @{}
foreach ($dict in @($SetMeta, $SetSeries, $SetQuality, $SetArtifact, $SetSig, $SetText, $SetEmote, $SetColorConditional, $SetLeak)) {
    foreach ($k in $dict.Keys) { $KnownRuleTags[$k] = $true }
}
$suspName = @($cov | Where-Object {
        $_.N -le 2 -and -not $_.Tag.StartsWith('@') -and -not $KnownRuleTags.ContainsKey($_.Tag) -and @($_.Tag -split ' ').Count -ge 2 -and (Test-NoContentWord $_.Tag)
    })
$suspWord = @($cov | Where-Object {
        $_.N -le 2 -and $_.Tag -notmatch '\d' -and -not $_.Tag.StartsWith('@') -and -not $KnownRuleTags.ContainsKey($_.Tag) -and @($_.Tag -split ' ').Count -eq 1 -and $_.Tag.Length -gt 4 -and (Test-NoContentWord $_.Tag)
    })
$rareOdd = @($cov | Where-Object {
        $_.N -ge 3 -and $_.Pct -lt 15 -and $_.Tag -notmatch '\d' -and -not $_.Tag.StartsWith('@') -and -not $KnownRuleTags.ContainsKey($_.Tag) -and (Test-NoContentWord $_.Tag)
    })
$contentOnce = @($cov | Where-Object { $_.N -le 1 -and -not (Test-NoContentWord $_.Tag) })

$rb = New-Object System.Text.StringBuilder
function Rvw([string]$s) { [void]$rb.AppendLine($s) }
function Fmt-RareList($list, [int]$cap = 15) {
    if (-not $list -or $list.Count -eq 0) { return '（无）' }
    $shown = @($list | Sort-Object N | Select-Object -First $cap | ForEach-Object { "``$($_.Tag)``($($_.N))" })
    $s = ($shown -join ', ')
    if ($list.Count -gt $cap) { $s += (" …（共 {0} 个，其余见 tag全量表.md）" -f $list.Count) }
    return $s
}
Rvw ("# {0} 待复核清单" -f $folderName)
Rvw ''
Rvw '清洗器只做「删 + 排 + 归一」，不猜内容。下面这些需要你对着图看一眼。'
Rvw '判定规则：**tag 里没有任何一个「描述画面的常用词」（颜色/身体/服装/动作/场景），而且出现次数 ≤2** → 大概率是角色名、作品名或 tagger 误标（有误报，人工确认即可）。'
Rvw ''
Rvw ("## A. 可能是角色名/作品名（2 词以上 + 没有内容词 + 出现 ≤2 次）：{0} 个" -f $suspName.Count)
Rvw ''
Rvw (Fmt-RareList $suspName)
Rvw ''
Rvw ("## B. 可能是 tagger 误标的生词（单个词 + 没有内容词 + 出现 ≤2 次）：{0} 个" -f $suspWord.Count)
Rvw ''
Rvw (Fmt-RareList $suspWord)
Rvw ''
Rvw ("## C. 稀有且没有内容词（3–15% 覆盖）：{0} 个" -f $rareOdd.Count)
Rvw ''
Rvw (Fmt-RareList $rareOdd)
Rvw ''
Rvw '## D. 近黑白图（如果这些不是画风的一部分，从数据集里剔掉）'
Rvw ''
if ($grayImages.Count) { Rvw (($grayImages | ForEach-Object { "``$_``" }) -join ', ') } else { Rvw '（无）' }
Rvw ''
Rvw '## E. 没标注 / 没图片'
Rvw ''
if ($missingCap.Count) { Rvw ('没标注的图片：' + (($missingCap | ForEach-Object { "``$_``" }) -join ', ')) }
if ($missingImg.Count) { Rvw ('没图片的标注：' + (($missingImg | ForEach-Object { "``$_``" }) -join ', ')) }
if (-not $missingCap.Count -and -not $missingImg.Count) { Rvw '（无）' }
Rvw ''
Rvw ("## 附：只出现 1 次的常规内容 tag：{0} 个（一般不用管，完整清单见 tag全量表.md）" -f $contentOnce.Count)

$fb = New-Object System.Text.StringBuilder
[void]$fb.AppendLine("# $folderName 全量 tag 覆盖度")
[void]$fb.AppendLine('')
[void]$fb.AppendLine('| tag | 出现次数 | 覆盖率 |')
[void]$fb.AppendLine('|---|---|---|')
foreach ($row in $cov) { [void]$fb.AppendLine(("| ``{0}`` | {1} | {2}% |" -f $row.Tag, $row.N, $row.Pct)) }

$pb = New-Object System.Text.StringBuilder
function AddP([string]$s) { [void]$pb.AppendLine($s) }
$trg = if ($Trigger) { "$Trigger, " } else { '' }
AddP '# 中立评测提示词（不要用训练集的 caption 当测试提示词）'
AddP ''
AddP '笔记 §1：拿训练 caption 当测试提示词，等于一边要求模型画亮、一边拿它和暗图比，量出来的差距是自己造的。'
AddP '所以下面这些是**中立内容**提示词，不含任何画风/光照/色彩词。'
AddP ''
AddP '## 出图设置'
AddP ''
AddP '- seed 固定（例如 42），分辨率与训练一致，步数固定'
AddP '- 每组都出 **三行**：训练原图 / 零权重底模 / 挂 LoRA —— 没有零权重对照，分不清是 LoRA 的功劳还是提示词本来就能画'
AddP '- 权重至少测 0.6 / 0.8 / 1.0 三档；单 seed 只够看趋势'
AddP ''
AddP '## 正向提示词'
AddP ''
$prompts = @(
    @('A 白底半身', '1girl, solo, upper body, looking at viewer, smile, blush, simple background, white background'),
    @('B 全身站姿', '1girl, solo, full body, standing, looking at viewer, arms at sides, white background'),
    @('C 室内坐姿', '1girl, solo, sitting, chair, indoors, looking at viewer, long hair, casual clothes'),
    @('D 双人', '2girls, standing, looking at viewer, indoors, upper body'),
    @('E 复杂背景', '1girl, solo, upper body, outdoors, city, street, looking at viewer'),
    @('F 动作', '1girl, solo, full body, running, looking at viewer, outdoors')
)
foreach ($p in $prompts) {
    AddP ("### $($p[0])")
    AddP ''
    AddP '```'
    AddP ($trg + $p[1])
    AddP '```'
    AddP ''
}
AddP '## 负向提示词'
AddP ''
AddP '```'
AddP 'lowres, worst quality, low quality, jpeg artifacts, blurry, bad anatomy, bad hands, extra digits, fewer digits, watermark, signature'
AddP '```'
AddP ''
if ($imgStats.Count -gt 0 -and $grayRatio -lt 0.05) {
    AddP '## 针对性：出图发黑白/发灰时追加到负向'
    AddP ''
    AddP '```'
    AddP 'monochrome, greyscale, greyscale with colored background, limited palette, spot color, sepia'
    AddP '```'
    AddP ''
    AddP '（你的数据集近黑白图占比很低，所以这些词是**负向**用的。加了还灰，说明是权重过高或过训练：往下调权重、取更早的存档。）'
}

$ib = New-Object System.Text.StringBuilder
function AddS([string]$s) { [void]$ib.AppendLine($s) }
AddS '# 画风 LoRA 标注规范（照这个口径写，别每次一套）'
AddS ''
AddS '## 写什么'
AddS ''
AddS '- **人物**：人数、景别、姿势、视角'
AddS '- **发型/发色/瞳色**：逐层写（长度 → 颜色 → 刘海 → 发饰）'
AddS '- **服装**：逐层写（颜色 → 材质 → 剪裁 → 装饰；每件单品分开写）'
AddS '- **配饰/道具**：头部、手部、身边的物件'
AddS '- **背景结构与母题密度**：室内/外、主要物件、画面里反复出现的母题'
AddS '- **画面内文字**：招牌、书页、对话气泡里的文字（这算内容）'
AddS ''
AddS '## 禁写（写了模型就不学了）'
AddS ''
AddS '- 光源方向、明暗对比、暗部颜色、高光处理'
AddS '- 色彩分级：饱和度、色调、限色、黑白（除非数据集本身混着黑白图）'
AddS '- 渲染方式：笔触、线宽、上色法、媒介、滤镜、景深、噪点'
AddS '- 风格流派 / 年代 / 画师名 / 作品名 / 角色名'
AddS '- 质量词：masterpiece、best quality、8k、detailed'
AddS '- 主观形容词：beautiful、cute、dreamy、atmospheric'
AddS ''
AddS '## 一句话'
AddS ''
AddS '**只写「画面里有什么」，不写「它看起来怎样」。看起来怎样，正是要让 LoRA 去学的东西。**'

# ============================================================================
# 16. 确认 & 应用
# ============================================================================
$apply = $false
if ($DryRun) {
    Write-Host ''
    Write-Host '预演模式，不修改任何文件。' -ForegroundColor Yellow
} elseif ($Yes) {
    $apply = $true
} elseif ($script:Interactive) {
    Write-Host ''
    $ans = Read-Host '确认把清洗结果写入这些 txt？(y = 应用 / 其他 = 只出报告)'
    if ($ans -match '^(y|yes)$') { $apply = $true }
}

New-Item -ItemType Directory -Path $outDir -Force | Out-Null
[System.IO.File]::WriteAllText($reportPath, $sb.ToString(), $utf8NoBom)
[System.IO.File]::WriteAllText($checkPath, $cb.ToString(), $utf8NoBom)
[System.IO.File]::WriteAllText($reviewPath, $rb.ToString(), $utf8NoBom)
[System.IO.File]::WriteAllText($promptPath, $pb.ToString(), $utf8NoBom)
[System.IO.File]::WriteAllText($specPath, $ib.ToString(), $utf8NoBom)
[System.IO.File]::WriteAllText($fullPath, $fb.ToString(), $utf8NoBom)

if ($apply) {
    $backup = Join-Path $outDir 'txt_原文件'
    New-Item -ItemType Directory -Path $backup -Force | Out-Null
    $written = 0
    foreach ($r in $results) {
        $rel = $r.Path.Substring($Folder.Length).TrimStart('\')
        $bpath = Join-Path $backup $rel
        $bdir = Split-Path $bpath -Parent
        if (-not (Test-Path -LiteralPath $bdir)) { New-Item -ItemType Directory -Path $bdir -Force | Out-Null }
        Copy-Item -LiteralPath $r.Path -Destination $bpath -Force
        if (-not $r.Changed) { continue }
        $orig = ''
        try { $orig = [System.IO.File]::ReadAllText($r.Path, [System.Text.Encoding]::UTF8) } catch { }
        $nl = if ($orig.EndsWith("`r`n")) { "`r`n" } elseif ($orig.EndsWith("`n")) { "`n" } else { '' }
        [System.IO.File]::WriteAllText($r.Path, ($r.NewText + $nl), $utf8NoBom)
        $written++
    }
    Write-Host ''
    Write-Host ("已写入 {0} 个文件（共 {1} 个，未改动的不重写）。" -f $written, $results.Count) -ForegroundColor Green
    Write-Host ("原文件备份：{0}" -f $backup) -ForegroundColor Green
    Write-Host ("报告      ：{0}" -f $outDir) -ForegroundColor Green
    $sample = @($changed) | Select-Object -First 1
    if ($sample) {
        Write-Host ''
        Write-Host '示例（第一个改动文件）：' -ForegroundColor Cyan
        Write-Host ('  旧: ' + $sample.OldText) -ForegroundColor DarkGray
        Write-Host ('  新: ' + $sample.NewText) -ForegroundColor White
    }

    # ---- 文本编码器缓存：不处理就等于没清洗 ----
    if ($npzFiles.Count -gt 0) {
        $doClean = $false
        if ($CleanCache) { $doClean = $true }
        elseif ($KeepCache) { $doClean = $false }
        elseif ($script:Interactive) {
            Write-Host ''
            Write-Host ("检测到 {0} 个 _anima_te.npz 文本编码器缓存 —— 它们不比对 caption，会继续用旧标注的 embedding。" -f $npzFiles.Count) -ForegroundColor Yellow
            $a2 = Read-Host '现在把它们挪到备份目录（训练时会自动重建）？(y = 挪走 / 其他 = 保留)'
            if ($a2 -match '^(y|yes)$') { $doClean = $true }
        }
        if ($doClean) {
            $npzBak = Join-Path $outDir 'npz_失效缓存'
            New-Item -ItemType Directory -Path $npzBak -Force | Out-Null
            $moved = 0
            foreach ($nf in $npzFiles) {
                $rel = $nf.FullName.Substring($Folder.Length).TrimStart('\')
                $dest = Join-Path $npzBak $rel
                $ddir = Split-Path $dest -Parent
                try {
                    if (-not (Test-Path -LiteralPath $ddir)) { New-Item -ItemType Directory -Path $ddir -Force | Out-Null }
                    Move-Item -LiteralPath $nf.FullName -Destination $dest -Force
                    $moved++
                } catch { }
            }
            Write-Host ("已挪走 {0} 个文本编码器缓存 → {1}" -f $moved, $npzBak) -ForegroundColor Green
        } else {
            Write-Host '保留了旧缓存。注意：不改缓存的情况下，这次清洗对训练不生效。' -ForegroundColor Yellow
        }
    }
} else {
    Write-Host ''
    Write-Host ("未修改任何文件。报告：{0}" -f $outDir) -ForegroundColor Yellow
}

Write-Host ''
Write-Host '下一步：' -ForegroundColor Cyan
Write-Host '  1) 打开 数据集体检.md，看「色彩实测」和「训练缓存排雷」两节'
Write-Host '  2) 打开 评测提示词.txt，按里面的中立提示词出对照图（记得零权重行）'
Write-Host '  3) 用 一键对照.bat 量出图 vs 训练原图的亮度/饱和/像素差'
Write-Host ''
