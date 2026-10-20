<#
  drive_to_server_export.ps1
  見積・注文管理システムが Googleドライブ「マイドライブ\見積管理_サーバー保存」へ書き出した見積書PDFを、
  会社サーバーの見積書フォルダ（export_config.json の ServerRoot）へ、件名・取引先・基板名で振り分けて保存する。
  システム側（44_server_export.gs）は金曜15時に書き出し、本スクリプトは金曜17時に実行（install_task.ps1 で登録）。

  振り分け（上から順に判定）
    取引先が藤商事
      ・件名に「価格改定」                  → 価格改定案内
      ・番号がD / 件名に「回路設計」「保守」 → 見積書（回路設計）\<年度>年度（4月始まり）
      ・番号がP（開発・試作）               → 見積書（F社開発）\<基板のフォルダ>
      ・件名に「注残」                      → 見積書（F社量産）\03_注残処理
      ・件名に【A__】等（機種共通基板）     → 見積書（F社量産）\組立_共通見積書
      ・件名に「仕掛基板」「PCB」          → 見積書（F社量産）\02_仕掛基板
      ・それ以外（M・N…量産）              → 見積書（F社量産）\<機種・基板のフォルダ>（送料・部品などは 99_機種以外）
    取引先がコナミ → 見積書（KAM）（件名に「試作」→ 試作見積、「ケース」→ ケース見積）
    取引先がエキサイト → 見積書（EXCITE）
    その他 → その他(Ｆ社以外)\<取引先のフォルダ>
  <基板のフォルダ>は、既存のサブフォルダ名（先頭の「26-003 」等の番号を除いた部分）が件名・基板名・機種コードに
  含まれるものを選ぶ（複数あれば名前の長いもの→新しいもの）。見つからない時は各フォルダ内の「_自動保存(振分け未定)」へ。

  使い方：
    確認だけ        : .\drive_to_server_export.ps1 -WhatIf
    既存で振分け検証: .\drive_to_server_export.ps1 -TestExisting   … サーバーにある見積書PDFで、振分け結果と実際の場所を比べる
#>
param(
  [switch]$WhatIf,
  [switch]$TestExisting,
  [string]$ConfigPath = (Join-Path $PSScriptRoot 'export_config.json'),
  [string]$StateDir = (Join-Path $env:LOCALAPPDATA 'MitsumoriServerSync')
)

$ErrorActionPreference = 'Stop'
$logPath = Join-Path $StateDir 'export.log'
New-Item -ItemType Directory -Force -Path $StateDir | Out-Null
function Write-Log([string]$msg) {
  $line = '{0}  {1}' -f (Get-Date -Format 'yyyy/MM/dd HH:mm:ss'), $msg
  Write-Host $line
  if (-not $WhatIf -and -not $TestExisting) { Add-Content -Path $logPath -Value $line -Encoding UTF8 }
}

$cfg = Get-Content -Path $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
$root = $cfg.ServerRoot
$src  = $cfg.DriveFolder
$F = $cfg.Folders
$unsorted = '_自動保存(振分け未定)'

function Norm([string]$s) {
  if (-not $s) { return '' }
  $t = $s.Normalize([Text.NormalizationForm]::FormKC).ToUpperInvariant()
  return ($t -replace '[\s　]+', '')
}

# サブフォルダ一覧のキャッシュ（キー = 番号を除いたフォルダ名を正規化したもの）
$folderCache = @{}
function Get-SubKeys([string]$dir) {
  if ($folderCache.ContainsKey($dir)) { return $folderCache[$dir] }
  $list = @()
  if (Test-Path -LiteralPath $dir) {
    Get-ChildItem -LiteralPath $dir -Directory -ErrorAction SilentlyContinue | ForEach-Object {
      $k = $_.Name -replace '^\s*(\d{2}-[\d,\-]+|N\d{4}-\d{3}|\d{2}_)\s*', ''
      $k = Norm $k
      if ($k.Length -ge 2 -and $k -notmatch '^[_☆]') { $list += [pscustomobject]@{ Key = $k; Name = $_.Name; Time = $_.LastWriteTime } }
    }
  }
  $folderCache[$dir] = $list
  return $list
}

# 件名などに含まれる既存サブフォルダを探す（英数字の途中では一致させない）
function Find-Sub([string]$dir, [string[]]$texts) {
  $hay = Norm (($texts | Where-Object { $_ }) -join ' / ')
  if (-not $hay) { return $null }
  $hits = Get-SubKeys $dir | Where-Object {
    $pat = '(?<![A-Z0-9])' + [regex]::Escape($_.Key) + '(?![0-9])'
    if ($_.Key -match '^[A-Z0-9+\-]+$') { $hay -match $pat } else { $hay.Contains($_.Key) }
  }
  if (-not $hits) { return $null }
  return ($hits | Sort-Object @{ e = { $_.Key.Length }; Descending = $true }, @{ e = { $_.Time }; Descending = $true } | Select-Object -First 1).Name
}

function Get-FiscalYear([string]$yyyymmdd) {
  if ($yyyymmdd -notmatch '^(\d{4})(\d{2})') { $d = Get-Date } else { $d = Get-Date -Year ([int]$Matches[1]) -Month ([int]$Matches[2]) -Day 1 }
  if ($d.Month -ge 4) { return $d.Year } else { return $d.Year - 1 }
}

# 振り分け先（ServerRoot からの相対パス）を返す
function Get-Destination($e) {
  $c = Norm $e.client; $s = $e.subject; $no = (Norm $e.quoteNo) -replace '^【見積書】', ''
  $texts = @($e.subject, $e.board, $e.model)
  # 宛先が自社名で読まれていても、FJ+ の基板なら藤商事案件として扱う
  if ($c -match 'サン電子' -and ((Norm "$($e.subject) $($e.board)") -match 'FJ\+')) { $c = '藤商事' }
  if ($c -match '藤商事') {
    if ($s -match '価格改定') { return $F.PriceRevision }
    if ($no -match '^D' -or $s -match '回路設計|保守') { return (Join-Path $F.Circuit ('{0}年度' -f (Get-FiscalYear $e.date))) }
    if ($no -match '^P') {
      $sub = Find-Sub (Join-Path $root $F.FDev) $texts
      return (Join-Path $F.FDev $(if ($sub) { $sub } else { $unsorted }))
    }
    if ($s -match '注残') { return (Join-Path $F.FMass '03_注残処理') }
    # 【A__】【E__】… は頭文字の機種に共通する見積（機種共通基板）→ 組立_共通見積書
    if ((Norm $s) -match '(?<![A-Z0-9])[A-Z]{1,2}__') { return (Join-Path $F.FMass '組立_共通見積書') }
    if ($s -match '仕掛基板|PCB') { return (Join-Path $F.FMass '02_仕掛基板') }
    $sub = Find-Sub (Join-Path $root $F.FMass) $texts
    if ($sub) { return (Join-Path $F.FMass $sub) }
    if ((Norm $s) -match '送料|輸送|運賃|金型|部品|イヤホン|イヤフォン|ROM書|キー|AXELL|AG6R|治具|リユース|作業費') { return (Join-Path $F.FMass '99_機種以外') }
    return (Join-Path $F.FMass $unsorted)
  }
  if ($c -match 'コナミ|KONAMI') {
    if ($s -match '試作') { return (Join-Path $F.KAM '試作見積') }
    if ($s -match 'ケース') { return (Join-Path $F.KAM 'ケース見積') }
    return $F.KAM
  }
  if ($c -match 'エキサイト|EXCITE') { return $F.Excite }
  $sub = Find-Sub (Join-Path $root $F.Other) @($e.client)
  return (Join-Path $F.Other $(if ($sub) { $sub } else { $unsorted }))
}

# ── 検証モード：サーバーにある見積書PDFで、振分け結果と実際の場所を比べる ──
if ($TestExisting) {
  $ok = 0; $ng = 0
  Get-ChildItem -LiteralPath $root -Recurse -File -Filter *.pdf -ErrorAction SilentlyContinue |
    Where-Object { $_.LastWriteTime -gt (Get-Date).AddMonths(-10) -and $_.Name -match '^(【見積書】)?([A-Z]\d{6,8}(-\d+)?)_(.+)_(\d{8})_(.+?)(_+\d)?( \(\d+\))?\.pdf$' } |
    ForEach-Object {
      $m = [regex]::Match($_.Name, '^(【見積書】)?([A-Z]\d{6,8}(-\d+)?)_(.+)_(\d{8})_(.+?)(_+\d)?( \(\d+\))?\.pdf$')
      $e = [pscustomobject]@{ quoteNo = $m.Groups[2].Value; subject = $m.Groups[4].Value; date = $m.Groups[5].Value; client = $m.Groups[6].Value; board = ''; model = '' }
      $actual = (Split-Path $_.FullName -Parent).Substring($root.Length + 1)
      $pred = Get-Destination $e
      if ($actual -eq $pred -or $actual.StartsWith($pred + '\')) { $ok++ } else { $ng++; Write-Host ('不一致: {0}' -f $_.Name); Write-Host ('    実際 = {0}' -f $actual); Write-Host ('    判定 = {0}' -f $pred) }
    }
  Write-Host ('一致 {0}件／不一致 {1}件' -f $ok, $ng)
  exit 0
}

# ── 通常：書き出されたPDFを会社サーバーへ ──
if (-not (Test-Path -LiteralPath $root)) { Write-Log "会社サーバーに接続できません（スキップ）: $root"; exit 0 }
if (-not (Test-Path -LiteralPath $src)) { Write-Log "書き出しフォルダがまだありません: $src"; exit 0 }
$sentDir = Join-Path $src '送信済み'
$saved = 0; $waiting = 0; $failed = 0
Get-ChildItem -LiteralPath $src -File -Filter '_manifest_*.json' | Sort-Object Name | ForEach-Object {
  $mf = $_
  $data = Get-Content -LiteralPath $mf.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
  $remaining = 0
  foreach ($e in $data.entries) {
    $pdf = Join-Path $src $e.file
    if (-not (Test-Path -LiteralPath $pdf)) {
      if (Test-Path -LiteralPath (Join-Path $sentDir $e.file)) { continue }   # 送信済み
      $remaining++; $waiting++; continue                                       # まだ同期されていない
    }
    $rel = Get-Destination $e
    $destDir = Join-Path $root $rel
    $dest = Join-Path $destDir $e.file
    if ($WhatIf) { Write-Host ('[確認] {0}  →  {1}' -f $e.file, $destDir); $remaining++; continue }
    try {
      New-Item -ItemType Directory -Force -Path $destDir | Out-Null
      if (Test-Path -LiteralPath $dest) { Write-Log ('同名ファイルがあるため上書きせずスキップ: {0}' -f $dest) }
      else { Copy-Item -LiteralPath $pdf -Destination $dest -ErrorAction Stop; $saved++; Write-Log ('保存: {0}  →  {1}' -f $e.file, $destDir) }
      New-Item -ItemType Directory -Force -Path $sentDir | Out-Null
      Move-Item -LiteralPath $pdf -Destination (Join-Path $sentDir $e.file) -Force
    } catch {
      $remaining++; $failed++; Write-Log ('失敗: {0}  ({1})' -f $e.file, $_.Exception.Message)
    }
  }
  if (-not $WhatIf -and $remaining -eq 0) {
    New-Item -ItemType Directory -Force -Path $sentDir | Out-Null
    Move-Item -LiteralPath $mf.FullName -Destination (Join-Path $sentDir $mf.Name) -Force
  }
}
$msg = '見積書 保存 {0}件／同期待ち {1}件／失敗 {2}件' -f $saved, $waiting, $failed
if ($saved -or $failed -or $waiting) { Write-Log $msg } else { Write-Host $msg }
