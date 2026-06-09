# 鹿行5市ダッシュボード｜データ更新のしくみ

`index.html` を常に最新に保つための運用メモです。

## 仕組み
- 全データは `index.html` 内の `data` 配列と `VINTAGE` / `LAST_UPDATED` に集約。
- ページ上部に「データ最終更新日」と各データの年次バッジを表示。45日以内＝緑（最新）、120日以内＝黄、超過＝赤（要更新）。
- `update-data.ps1` が公式ソースを再取得し、農業産出額（e-Stat）を自動で書き換え、最終更新日を当日に更新します。

## 手動で更新する
```powershell
cd C:\Users\81907\Desktop\kuro-do
& .\update-data.ps1
```
- `sources\` に最新のPDF・Excelがダウンロードされます。
- 農業産出額（令和○年）と最終更新日は自動更新されます。

## 自動で更新する（定期実行）
Windowsタスクスケジューラに登録済みなら、毎週月曜に自動実行されます。
- 確認: `Get-ScheduledTask -TaskName RokkouDashboardUpdate`
- 解除: `Unregister-ScheduledTask -TaskName RokkouDashboardUpdate -Confirm:$false`
- 手動登録する場合:
```powershell
$action  = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument '-NoProfile -File "C:\Users\81907\Desktop\kuro-do\update-data.ps1"'
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At 9:00am
Register-ScheduledTask -TaskName RokkouDashboardUpdate -Action $action -Trigger $trigger -Description '鹿行5市ダッシュボードのデータ更新'
```

## 自動化できない部分（年1回・手動／Claude依頼）
人口・財政力指数・決算（歳入歳出・基金・健全化指標）・製造品出荷額は**PDF形式**のため、スクリプトは再ダウンロードのみ行います。新年度版が公表されたら：
1. `update-data.ps1` を実行して `sources\` を最新化
2. Claude に「`sources` のPDFから最新値で `index.html` を更新して」と依頼
   （または `update-data.ps1` 冒頭の `$sources` のURLを新年度版に差し替え）

## 出典
- 人口: 茨城県／国土地理協会 住民基本台帳
- 財政力指数・決算・基金・健全化指標・製造品出荷額・市内総生産・住民所得: 茨城県「行財政概要」
- 農業産出額: 農林水産省「市町村別農業産出額（推計）」/ e-Stat
