<#
  watch_server_folders.ps1
  sync_config.json の Sources（会社サーバーの保存先フォルダ）を見張り、PDFが保存されたら
  約1分後に server_to_drive_sync.ps1 を実行して Googleドライブの取込フォルダへ送る（常駐）。
  ・保存中のファイルを送らないよう、最後の変化から70秒待ってから実行
  ・ネットワーク切断などで見張りが止まったら、10分ごとに張り直す
  ・多重起動しない
  タスクスケジューラ（install_task.ps1）からログオン時に起動される。
#>
param(
  [string]$ConfigPath = (Join-Path $PSScriptRoot 'sync_config.json'),
  [string]$StateDir = (Join-Path $env:LOCALAPPDATA 'MitsumoriServerSync')
)

$mutex = New-Object System.Threading.Mutex($false, 'Local\MitsumoriServerWatch')
if (-not $mutex.WaitOne(0)) { exit 0 } # 既に起動している

$syncScript = Join-Path $PSScriptRoot 'server_to_drive_sync.ps1'
$logPath = Join-Path $StateDir 'sync.log'
New-Item -ItemType Directory -Force -Path (Split-Path $logPath) | Out-Null
function Write-Log([string]$m) { Add-Content -Path $logPath -Value ('{0}  [監視] {1}' -f (Get-Date -Format 'yyyy/MM/dd HH:mm:ss'), $m) -Encoding UTF8 }

$cfg = Get-Content -Path $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
$watchers = @()
function Start-Watchers {
  foreach ($w in $script:watchers) { try { $w.EnableRaisingEvents = $false; $w.Dispose() } catch {} }
  Get-EventSubscriber -ErrorAction SilentlyContinue | Where-Object { $_.SourceIdentifier -like 'MSW_*' } | Unregister-Event
  $script:watchers = @()
  $i = 0
  foreach ($s in $cfg.Sources) {
    if (-not (Test-Path -LiteralPath $s.Path)) { Write-Log "フォルダに接続できません（10分後に再試行）: $($s.Path)"; continue }
    $w = New-Object System.IO.FileSystemWatcher($s.Path, '*.pdf')
    $w.IncludeSubdirectories = $true
    $w.NotifyFilter = [System.IO.NotifyFilters]'FileName, LastWrite, Size'
    $w.InternalBufferSize = 65536
    foreach ($ev in 'Created', 'Changed', 'Renamed', 'Error') {
      Register-ObjectEvent -InputObject $w -EventName $ev -SourceIdentifier ('MSW_{0}_{1}' -f $i, $ev) | Out-Null
    }
    $w.EnableRaisingEvents = $true
    $script:watchers += $w; $i++
  }
  Write-Log ('見張りを開始しました（{0}フォルダ）' -f $script:watchers.Count)
}

function Invoke-Sync([string]$why) {
  Write-Log "送信を実行（$why）"
  & powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File $syncScript -ConfigPath $ConfigPath -StateDir $StateDir | Out-Null
}

Start-Watchers
Invoke-Sync '起動時'
$lastEvent = $null
$lastRebuild = Get-Date
while ($true) {
  $e = Wait-Event -Timeout 15
  if ($e) {
    if ($e.SourceIdentifier -like '*_Error') { Write-Log '見張りでエラーが発生したため張り直します'; Start-Watchers }
    else { $lastEvent = Get-Date }
    Remove-Event -EventIdentifier $e.EventIdentifier -ErrorAction SilentlyContinue
    continue
  }
  if ($lastEvent -and ((Get-Date) - $lastEvent).TotalSeconds -ge 70) { $lastEvent = $null; Invoke-Sync 'フォルダにPDFが保存されました' }
  if (((Get-Date) - $lastRebuild).TotalMinutes -ge 10) {
    $lastRebuild = Get-Date
    if (($cfg.Sources | Where-Object { Test-Path -LiteralPath $_.Path }).Count -ne $watchers.Count) { Start-Watchers }
  }
}
