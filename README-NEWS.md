# ニュース・SNS統合フィード

鹿嶋市、神栖市、アントラーズ関連のニュースとSNS発信を自動取得・統一表示するフィード機能です。

## 機能

### リアルタイム表示
- 5分ごとに自動更新（バックグラウンド）
- 新しいニュースは常に最新状態
- 更新ボタンで手動リロードも可能

### 複数ソースの統一表示
取得対象：
- **朝日新聞 RSS**（茨城関連版）
- **アントラーズ公式 RSS**（鹿島アントラーズニュース）
- **その他地元ニュースサイト**（拡張予定）

### 重複排除
- 同じニュースが複数ソースにある場合、自動で統一
- 「同報件数」バッジで表示
- タイトル正規化で精度向上

### カテゴリ分類
自動的に3カテゴリに分類：
- 🔴 **antlers**（アントラーズ）
- 🟢 **kamisu**（神栖市）
- 🔴 **kashima**（鹿嶋市）

## ファイル構成

| ファイル | 説明 |
|---------|------|
| `index.html` | メインダッシュボード＋ニュースセクション |
| `news-data.json` | ニュースデータ格納（JSON） |
| `update-data.ps1` | RSS自動取得＆更新スクリプト |

## 使い方

### 1. ブラウザで確認
```
http://localhost:8765
```
"ニュース・SNS" セクションを見てください。

### 2. ニュース更新

**自動更新（デフォルト）**
- 5分ごとに自動で最新化
- バックグラウンドで動作

**手動更新**
- ページ上の「🔄 ニュースを更新」ボタンをクリック
- すぐに最新ニュースを取得

**PowerShell で更新**
```powershell
cd C:\Users\81907\Desktop\kuro-do
.\update-data.ps1
```

### 3. 更新スケジュール設定（Windows タスク）

毎日 8:00 に自動更新を実行するには：

```powershell
# PowerShell を管理者として実行
$action = New-ScheduledTaskAction -Execute 'powershell.exe' `
  -Argument '-NoProfile -ExecutionPolicy Bypass -File "C:\Users\81907\Desktop\kuro-do\update-data.ps1"'
$trigger = New-ScheduledTaskTrigger -Daily -At 08:00
Register-ScheduledTask -TaskName 'kuro-do-news-update' -Action $action -Trigger $trigger
```

## JSON フォーマット

`news-data.json` の形式：

```json
[
  {
    "title": "ニュースタイトル",
    "description": "説明文",
    "published": "2026-06-10T15:30:00Z",  // ISO 8601 形式
    "url": "https://example.com",
    "source": "朝日新聞",
    "category": "kashima",  // kashima | kamisu | antlers
    "duplicates": [
      {"source": "茨城新聞", "url": "https://..."}
    ]
  }
]
```

## PowerShell スクリプト動作

`update-data.ps1` を実行すると：

1. **農業産出額 Excel を再取得**
2. **ニュース RSS を自動取得**
   - 朝日新聞（茨城版）
   - アントラーズ公式
3. **同じニュース を統一**（タイトルで正規化）
4. **JSON に出力**（news-data.json）
5. **ブラウザが自動リロード**

## トラブルシューティング

### ニュースが表示されない
- ブラウザを再読み込み（F5）
- コンソール（F12 → Console）でエラーを確認

### ニュースが古い
- 「🔄 ニュースを更新」ボタンをクリック
- またはPowerShellで `.\update-data.ps1` を実行

### PowerShell スクリプトが失敗
```powershell
# 実行ポリシーを確認
Get-ExecutionPolicy

# 必要に応じて設定
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## 今後の拡張

- [ ] 茨城新聞 RSS 自動取得
- [ ] Twitter/X API 統合（認証設定後）
- [ ] 鹿嶋市・神栖市 公式サイト スクレイピング
- [ ] キーワード検索機能
- [ ] ニュースアーカイブ（期間検索）
- [ ] Slack/Teams通知連携
