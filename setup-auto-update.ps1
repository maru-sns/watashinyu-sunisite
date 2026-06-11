# setup-auto-update.ps1
# ニュース自動更新をWindowsタスクスケジューラーに登録する
# 初回だけ実行してください（管理者権限不要）

$root       = $PSScriptRoot
if (-not $root) { $root = (Get-Location).Path }
$scriptPath = Join-Path $root 'update-news.ps1'
$vbs        = Join-Path $root 'run-hidden.vbs'   # 黒いウィンドウを出さない隠し起動ランチャー

# ---- タスク1: ニュース更新（10分ごと、ログオン中） ----
# wscript.exe 経由で run-hidden.vbs を呼び、PowerShell を完全非表示で実行（コンソール非表示）
$taskName   = 'KuroDo_NewsUpdate'
$action     = New-ScheduledTaskAction `
    -Execute 'wscript.exe' `
    -Argument "`"$vbs`" `"$scriptPath`""
$trigger    = New-ScheduledTaskTrigger -RepetitionInterval (New-TimeSpan -Minutes 10) -Once -At (Get-Date)
$settings   = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 2) `
    -MultipleInstances IgnoreNew `
    -StartWhenAvailable

# 既存タスクを削除してから再登録
Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
Register-ScheduledTask `
    -TaskName $taskName `
    -Action   $action `
    -Trigger  $trigger `
    -Settings $settings `
    -RunLevel Limited | Out-Null

Write-Host "タスク登録完了: $taskName" -ForegroundColor Green
Write-Host "  → 10 分ごとに update-news.ps1 が自動実行されます"

# ---- タスク2: 毎朝 7:00 に全データ更新（農業産出額 + ニュース） ----
$taskNameFull = 'KuroDo_FullUpdate'
$actionFull   = New-ScheduledTaskAction `
    -Execute 'wscript.exe' `
    -Argument "`"$vbs`" `"$(Join-Path $root 'update-data.ps1')`""
$triggerFull  = New-ScheduledTaskTrigger -Daily -At '07:00'
$settingsFull = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 10) `
    -StartWhenAvailable

Unregister-ScheduledTask -TaskName $taskNameFull -Confirm:$false -ErrorAction SilentlyContinue
Register-ScheduledTask `
    -TaskName $taskNameFull `
    -Action   $actionFull `
    -Trigger  $triggerFull `
    -Settings $settingsFull `
    -RunLevel Limited | Out-Null

Write-Host "タスク登録完了: $taskNameFull" -ForegroundColor Green
Write-Host "  → 毎朝 7:00 に全データ（農業産出額＋ニュース）を更新します"

# ---- すぐに1回実行 ----
Write-Host ""
Write-Host "初回実行中..." -ForegroundColor Cyan
& $scriptPath
Write-Host "完了！ニュースデータを取得しました。" -ForegroundColor Green
Write-Host ""
Write-Host "登録済みタスク一覧:" -ForegroundColor Cyan
Get-ScheduledTask | Where-Object { $_.TaskName -like 'KuroDo_*' } |
  Select-Object TaskName, State |
  Format-Table -AutoSize
