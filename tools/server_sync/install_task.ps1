<#
  install_task.ps1
  server_to_drive_sync.ps1 を Windows のタスクスケジューラに登録する（ログオン中、10分ごとに実行）。
  削除する時： Unregister-ScheduledTask -TaskName '見積管理_サーバーからDriveへ自動コピー' -Confirm:$false
#>
param([int]$IntervalMinutes = 10)
$taskName = '見積管理_サーバーからDriveへ自動コピー'
$script   = Join-Path $PSScriptRoot 'server_to_drive_sync.ps1'
$action   = New-ScheduledTaskAction -Execute 'powershell.exe' `
            -Argument ('-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "{0}"' -f $script)
$trigger  = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) `
            -RepetitionInterval (New-TimeSpan -Minutes $IntervalMinutes) -RepetitionDuration (New-TimeSpan -Days 3650)
$logon    = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -DontStopIfGoingOnBatteries -AllowStartIfOnBatteries `
            -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 30)
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited
Register-ScheduledTask -TaskName $taskName -Action $action -Trigger @($trigger, $logon) -Settings $settings `
  -Principal $principal -Description '会社サーバーの注文書PDFを Googleドライブ「見積管理_自動取込」へコピー（見積・注文管理システム）' -Force | Out-Null
Write-Host "登録しました: $taskName（$IntervalMinutes 分ごと）"
