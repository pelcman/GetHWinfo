# セットアップガイド

Googleスプレッドシートへの同期が動作しない場合の対処方法

## 問題: "ページが見つかりません" エラー

Web App URLが無効、または削除されています。新しくデプロイする必要があります。

## 解決手順

### 1. Googleスプレッドシートを開く

スプレッドシートを開くか、新規作成してください。

URL: https://docs.google.com/spreadsheets/d/1J8wYgDj4mOb-mw3r4AcsHccpNnyweQJLlcpZteF7IyA/edit

### 2. Apps Scriptを開く

1. スプレッドシートで **拡張機能** > **Apps Script** をクリック
2. 既存のコードを全て削除

### 3. AppsScript.gsの内容をコピペ

1. `GetHWinfo\AppsScript.gs` ファイルを開く
2. 全ての内容をコピー
3. Apps Scriptエディタに貼り付け
4. **保存**ボタンをクリック（Ctrl+S）

### 4. 新しいデプロイを作成

1. **デプロイ** > **新しいデプロイ** をクリック
2. 歯車アイコン > **ウェブアプリ** を選択
3. 設定:
   - **説明**: `Hardware Info Sync` (任意)
   - **次のユーザーとして実行**: **自分**
   - **アクセスできるユーザー**: **全員** を選択
4. **デプロイ** をクリック
5. 権限の確認が表示されたら:
   - アカウントを選択
   - "このアプリは確認されていません" → **詳細** → **安全ではないページに移動**
   - **許可** をクリック
6. **ウェブアプリ URL** をコピー

### 5. config.jsonを更新

1. `GetHWinfo\config.json` を開く
2. WebAppUrlを新しいURLに更新:

```json
{
  "WebAppUrl": "ここに新しいURLを貼り付け"
}
```

3. 保存

### 6. テスト実行

PowerShellで以下を実行:

```powershell
cd c:\Users\yoshida_k\Desktop\GetHWinfo
.\scripts\test_connection.ps1
```

成功すれば "All tests passed!" と表示されます。

### 7. 実際のデータで同期

```powershell
.\scripts\sync_to_spreadsheet_curl.ps1
```

または `run.bat` を実行。

## トラブルシューティング

### タイムアウトエラーが出る場合

`sync_to_spreadsheet_curl.ps1` を使用してください（curlベース）。

### 権限エラーが出る場合

Apps Scriptエディタで:
1. **実行** > **testUpdateSpreadsheet** を選択
2. **実行** をクリック
3. 権限を許可

### データが同期されない場合

1. スプレッドシートを開く
2. シート名が "Sheet1" であることを確認
3. または `AppsScript.gs` の17行目を変更:
   ```javascript
   const SHEET_NAME = '実際のシート名';
   ```

## 補足: sync_to_spreadsheet_curl.ps1 を標準に設定

`run.bat` を編集して、curl版を使用するようにします:

```batch
PowerShell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\sync_to_spreadsheet_curl.ps1"
```
