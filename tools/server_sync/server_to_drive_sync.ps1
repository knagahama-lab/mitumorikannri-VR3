<#
  server_to_drive_sync.ps1
  会社サーバーに保存された注文書PDFを、Googleドライブ パソコン版のフォルダ
  （マイドライブ\見積管理_自動取込\注文書(試作)／注文書(量産)）へ自動コピーする。
  コピーされたPDFは Googleドライブが自動でアップロードし、見積・注文管理システム（GAS）が
  5分ごとに読み取って登録する（43_server_sync.gs）。

  ・対象：SourceRoot 直下の「○○注文書」フォルダ（例：藤商事注文書、コナミ注文書、エキサイト注文書）の
          「試作注文書」「量産注文書」以下のPDF（サブフォルダも含む）
  ・コピー先のファイル名：[取引先名]元のファイル名.pdf（例：[藤商事]発注書206424.pdf）
  ・一度コピーしたファイルは記録し、二度コピーしない（ファイルが更新された場合は差し替えとして再コピー）
  ・保存直後（1分以内）のファイルは書き込み中の可能性があるため次回に回す

  使い方：
    初回      : .\server_to_drive_sync.ps1 -InitialDays 0      … 既存のPDFは「コピー済み」として記録だけする（以後の新しいPDFだけ送る）
                .\server_to_drive_sync.ps1 -InitialDays 14     … 直近14日以内に保存されたPDFだけ送り、それより前は記録だけ
    確認だけ  : .\server_to_drive_sync.ps1 -WhatIf             … コピーせず、対象ファイルを表示
    定期実行  : install_task.ps1 でタスクスケジューラに登録（10分ごと）
#>
param(
  [switch]$WhatIf,
  [int]$InitialDays = -1,
  [string]$ConfigPath = (Join-Path $PSScriptRoot 'sync_config.json')
)

$ErrorActionPreference = 'Stop'
$stateDir  = Join-Path $env:LOCALAPPDATA 'MitsumoriServerSync'
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
if (-not (Test-Path -LiteralPath $sourceRoot)) { Write-Log "会社サーバーに接続できません（スキップ）: $sourceRoot"; exit 0 }
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
  exit 1
}

# ── 対象ファイルを集める ──
$now = Get-Date
$targets = @()
Get-ChildItem -LiteralPath $sourceRoot -Directory | Where-Object { $_.Name -like $cfg.ClientFolderPattern } | ForEach-Object {
  $clientDir = $_
  $client = $clientDir.Name -replace $cfg.ClientNameStrip, ''
  foreach ($cat in $cfg.Categories) {
    $src = Join-Path $clientDir.FullName $cat.Source
    if (-not (Test-Path -LiteralPath $src)) { continue }
    Get-ChildItem -LiteralPath $src -Recurse -File -Filter '*.pdf' -ErrorAction SilentlyContinue | ForEach-Object {
      $targets += [pscustomobject]@{ File = $_; Client = $client; Dest = (Join-Path $destRoot $cat.Dest) }
    }
  }
}

$copied = 0; $baselined = 0; $pending = 0; $failed = 0
foreach ($t in $targets) {
  $f = $t.File
  $key = $f.FullName
  $sig = '{0}|{1}' -f $f.Length, $f.LastWriteTimeUtc.Ticks
  if ($state.ContainsKey($key) -and $state[$key] -eq $sig) { continue }

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
  if ($WhatIf) { Write-Host ('[確認] {0}  →  {1}' -f $key, $destPath); continue }
  try {
    New-Item -ItemType Directory -Force -Path $t.Dest | Out-Null
    Copy-Item -LiteralPath $key -Destination $destPath -ErrorAction Stop
    $state[$key] = $sig; $copied++
    Write-Log ('コピー: {0}  →  {1}' -f $key, $destPath)
  } catch {
    $failed++; Write-Log ('失敗: {0}  ({1})' -f $key, $_.Exception.Message)
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
