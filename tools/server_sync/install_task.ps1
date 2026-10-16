<#
  install_task.ps1
  タスクスケジューラに2つ登録する（現在のユーザー・ログオン中に実行）
    ① 見積管理_サーバー監視（常駐）      … ログオン時に watch_server_folders.ps1 を起動。PDFが保存されたら約1分後に送る
    ② 見積管理_サーバー定期送信（平日30分） … 月〜金 7:00〜21:00 に30分おきに server_to_drive_sync.ps1（見張りの取りこぼし対策）
  削除する時：
    Unregister-ScheduledTask -TaskName '見積管理_サーバー監視（常駐）' -Confirm:$false
    Unregister-ScheduledTask -TaskName '見積管理_サーバー定期送信（平日30分）' -Confirm:$false
#>
$user = "$env:USERDOMAIN\$env:USERNAME"
$principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Limited
# 黒い画面を出さないため、wscript.exe（画面なし）経由で PowerShell を起動する（run_hidden.vbs）
$wsh = 'wscript.exe'
$vbs = Join-Path $PSScriptRoot 'run_hidden.vbs'

# ① 常駐の見張り
$watchName = '見積管理_サーバー監視（常駐）'
$watchAct  = New-ScheduledTaskAction -Execute $wsh -Argument ('"{0}" "{1}"' -f $vbs, (Join-Path $PSScriptRoot 'watch_server_folders.ps1'))
$watchTrg  = New-ScheduledTaskTrigger -AtLogOn -User $user
$watchSet  = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
             -MultipleInstances IgnoreNew -ExecutionTimeLimit ([TimeSpan]::Zero) -RestartCount 999 -RestartInterval (New-TimeSpan -Minutes 1)
Register-ScheduledTask -TaskName $watchName -Action $watchAct -Trigger $watchTrg -Settings $watchSet -Principal $principal `
  -Description '会社サーバーの注文書保存フォルダを見張り、PDFが保存されたら Googleドライブ「見積管理_自動取込」へ送る' -Force | Out-Null

# ② 平日30分おき（7:00〜21:00）
$syncName = '見積管理_サーバー定期送信（平日30分）'
$syncAct  = New-ScheduledTaskAction -Execute $wsh -Argument ('"{0}" "{1}"' -f $vbs, (Join-Path $PSScriptRoot 'server_to_drive_sync.ps1'))
$syncTrg  = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday, Tuesday, Wednesday, Thursday, Friday -At '07:00'
$syncTrg.Repetition = (New-ScheduledTaskTrigger -Once -At '07:00' -RepetitionInterval (New-TimeSpan -Minutes 30) -RepetitionDuration (New-TimeSpan -Hours 14)).Repetition
$syncSet  = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
            -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 20)
Register-ScheduledTask -TaskName $syncName -Action $syncAct -Trigger $syncTrg -Settings $syncSet -Principal $principal `
  -Description '会社サーバーの注文書PDFを Googleドライブ「見積管理_自動取込」へ送る（平日30分おき・見張りの取りこぼし対策）' -Force | Out-Null

# 旧版のタスク（10分おき）があれば削除
Unregister-ScheduledTask -TaskName '見積管理_サーバーからDriveへ自動コピー' -Confirm:$false -ErrorAction SilentlyContinue

# 初回の記録が無ければ作る（今あるPDFは送らず、これから保存される分だけ送る）
if (-not (Test-Path (Join-Path $env:LOCALAPPDATA 'MitsumoriServerSync\state.json'))) {
  & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'server_to_drive_sync.ps1') -InitialDays 0
}

Start-ScheduledTask -TaskName $watchName
Write-Host "登録しました: $watchName／$syncName（見張りを開始しました）"
