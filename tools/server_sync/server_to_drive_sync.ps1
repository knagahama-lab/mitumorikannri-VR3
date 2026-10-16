<#
  server_to_drive_sync.ps1
  会社サーバーに保存された注文書PDFを、Googleドライブ パソコン版のフォルダ
  （マイドライブ\見積管理_自動取込\注文書(試作)／注文書(量産)）へ自動コピーする。
  コピーされたPDFは Googleドライブが自動でアップロードし、見積・注文管理システム（GAS）が
  5分ごとに読み取って登録する（43_server_sync.gs）。

  ・対象：sync_config.json の Sources に書いたフォルダのPDF（例：藤商事注文書\試作注文書\１．西宮）
          ※ Sources が無い場合は SourceRoot 直下の「○○注文書」の「試作注文書」「量産注文書」以下すべて
  ・コピー先のファイル名：[取引先名]元のファイル名.pdf（例：[藤商事]発注書206424.pdf）
  ・一度コピーしたファイルは「取引先＋試作/量産＋ファイル名＋サイズ」で記録し、二度コピーしない
    （担当者フォルダ → 完了 → 済 のように移動しても再コピーしない。内容が変わってサイズが変われば差し替えとして再コピー）
  ・保存直後（1分以内）のファイルは書き込み中の可能性があるため次回に回す

  使い方：
    初回      : .\server_to_drive_sync.ps1 -InitialDays 0      … 既存のPDFは「コピー済み」として記録だけする（以後の新しいPDFだけ送る）
                .\server_to_drive_sync.ps1 -InitialDays 14     … 直近14日以内に保存されたPDFだけ送り、それより前は記録だけ
    確認だけ  : .\server_to_drive_sync.ps1 -WhatIf             … コピーせず、対象ファイルを表示
    定期実行  : install_task.ps1 でタスクスケジューラに登録（見張り＋平日30分おき、画面は出ない）
#>
param(
  [switch]$WhatIf,
  [int]$InitialDays = -1,
  [string]$ConfigPath = (Join-Path $PSScriptRoot 'sync_config.json'),
  [string]$StateDir = (Join-Path $env:LOCALAPPDATA 'MitsumoriServerSync')   # 記録・ログの保存先（テスト用に変更可）
)

$ErrorActionPreference = 'Stop'
# 見張り（watch_server_folders.ps1）と定期実行が同時に動いても記録が壊れないよう、1つずつ実行
$syncMutex = New-Object System.Threading.Mutex($false, 'Local\MitsumoriServerSyncRun')
try {
  if (-not $syncMutex.WaitOne([TimeSpan]::FromMinutes(5))) { exit 0 }
} catch [System.Threading.AbandonedMutexException] { }   # 前の実行が強制終了されていても続ける
$stateDir  = $StateDir
$statePath = Join-Path $stateDir 'state.json'
$logPath   = Join-Path $stateDir 'sync.log'
New-Item -ItemType Directory -Force -Path $stateDir | Out-Null

function Write-Log([string]$msg) {
  $line = '{0}  {1}' -f (Get-Date -Format 'yyyy/MM/dd HH:mm:ss'), $msg
  Write-Host $line
  if (-not $WhatIf) {
    Add-Content -Path $logPath -Value $line -Encoding UTF8
    $lines = Get-Content -Path $logPath -Encoding UTF8 -ErrorAction SilentlyContinue
    if ($lines.Count -gt 3000) { $lines | Select-Object -Last 2000 | Set-Content -Path $logPath -Encoding UTF8 }
  }
}

# ── 設定 ──
$cfg = Get-Content -Path $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
$sourceRoot = $cfg.SourceRoot
$destRoot   = $cfg.DestRoot
$checkPath  = if ($cfg.Sources) { $cfg.Sources[0].Path } else { $sourceRoot }
if (-not (Test-Path -LiteralPath $checkPath)) { Write-Log "会社サーバーに接続できません（スキップ）: $checkPath"; exit 0 }
if (-not (Test-Path -LiteralPath (Split-Path $destRoot -Parent))) { Write-Log "Googleドライブ（パソコン版）が見つかりません。起動・ログインを確認してください: $destRoot"; exit 0 }

# ── 状態（コピー済みファイルの記録） ──
$state = @{}
$initialized = $false
if (Test-Path -LiteralPath $statePath) {
  $raw = Get-Content -Path $statePath -Raw -Encoding UTF8 | ConvertFrom-Json
  $initialized = [bool]$raw.initialized
  if ($raw.files) { $raw.files.PSObject.Properties | ForEach-Object { $state[$_.Name] = [string]$_.Value } }
}
if (-not $initialized -and $InitialDays -lt 0) {
  Write-Log '初回実行です。-InitialDays を指定してください（0 = 既存は記録のみ、14 = 直近14日分だけ送る）。'
  Write-Log ('  [診断] 記録ファイル: {0} 存在={1} 内容={2}' -f $statePath, (Test-Path -LiteralPath $statePath), ((Get-Content -LiteralPath $statePath -Raw -ErrorAction SilentlyContinue) -replace '\s+', ' '))
  exit 1
}

# ── 対象ファイルを集める ──
$now = Get-Date
$targets = @()
if ($cfg.Sources) {
  # 指定フォルダ方式：保存先のフォルダ（担当者フォルダ）だけを見る
  foreach ($s in $cfg.Sources) {
    if (-not (Test-Path -LiteralPath $s.Path)) { Write-Log "フォルダが見つかりません（スキップ）: $($s.Path)"; continue }
    Get-ChildItem -LiteralPath $s.Path -Recurse -File -Filter '*.pdf' -ErrorAction SilentlyContinue | ForEach-Object {
      $targets += [pscustomobject]@{ File = $_; Client = $s.Client; Dest = (Join-Path $destRoot $s.Dest); Cat = $s.Dest }
    }
  }
} else {
Get-ChildItem -LiteralPath $sourceRoot -Directory | Where-Object { $_.Name -like $cfg.ClientFolderPattern } | ForEach-Object {
  $clientDir = $_
  $client = $clientDir.Name -replace $cfg.ClientNameStrip, ''
  foreach ($cat in $cfg.Categories) {
    $src = Join-Path $clientDir.FullName $cat.Source
    if (-not (Test-Path -LiteralPath $src)) { continue }
    Get-ChildItem -LiteralPath $src -Recurse -File -Filter '*.pdf' -ErrorAction SilentlyContinue | ForEach-Object {
      $targets += [pscustomobject]@{ File = $_; Client = $client; Dest = (Join-Path $destRoot $cat.Dest); Cat = $cat.Dest }
    }
  }
}
}

$copied = 0; $baselined = 0; $pending = 0; $failed = 0
foreach ($t in $targets) {
  $f = $t.File
  # 記録のキーは場所ではなく「種別｜取引先｜ファイル名｜サイズ」（フォルダ移動では再コピーしない）
  $key = '{0}|{1}|{2}|{3}' -f $t.Cat, $t.Client, $f.Name.ToLowerInvariant(), $f.Length
  $sig = $f.FullName
  if ($state.ContainsKey($key)) { continue }

  # 初回：InitialDays より前のファイルは「コピー済み」として記録だけ
  if (-not $initialized -and $f.LastWriteTime -lt $now.AddDays(-$InitialDays)) { $state[$key] = $sig; $baselined++; continue }
  # 保存直後は書き込み中かもしれないので次回へ
  if ($f.LastWriteTime -gt $now.AddMinutes(-1)) { $pending++; continue }

  $destName = '[{0}]{1}' -f $t.Client, $f.Name
  $destPath = Join-Path $t.Dest $destName
  $n = 2
  while (Test-Path -LiteralPath $destPath) {
    $destPath = Join-Path $t.Dest ('[{0}]{1}_{2}{3}' -f $t.Client, $f.BaseName, $n, $f.Extension); $n++
  }
  if ($WhatIf) { Write-Host ('[確認] {0}  →  {1}' -f $f.FullName, $destPath); continue }
  try {
    New-Item -ItemType Directory -Force -Path $t.Dest | Out-Null
    Copy-Item -LiteralPath $f.FullName -Destination $destPath -ErrorAction Stop
    $state[$key] = $sig; $copied++
    Write-Log ('コピー: {0}  →  {1}' -f $f.FullName, $destPath)
  } catch {
    $failed++; Write-Log ('失敗: {0}  ({1})' -f $f.FullName, $_.Exception.Message)
  }
}

if (-not $WhatIf) {
  $out = [ordered]@{ initialized = $true; updatedAt = (Get-Date -Format 's'); files = $state }
  $tmp = "$statePath.tmp"
  ($out | ConvertTo-Json -Depth 4) | Set-Content -Path $tmp -Encoding UTF8
  Move-Item -LiteralPath $tmp -Destination $statePath -Force
}
$msg = 'コピー {0}件／既存として記録 {1}件／書き込み待ち {2}件／失敗 {3}件（対象PDF {4}件）' -f $copied, $baselined, $pending, $failed, $targets.Count
if ($copied -or $baselined -or $failed -or $WhatIf) { Write-Log $msg } else { Write-Host $msg }
