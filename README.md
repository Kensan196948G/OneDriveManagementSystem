#OneDrive for Business KFM運用管理ツール

## 概要
このツールは、OneDrive for Businessのデプロイメント、監視、およびトラブルシューティングを自動化するPowerShellスクリプト群です。Known Folder Move (KFM)の一括展開や運用管理を効率的に実施できます。

## 重要な前提条件
- グローバル管理者権限が必須です
- Microsoft GraphモジュールとMSOnlineモジュールが必要です（自動インストールされます）
- インターネット接続が必要です

## 主な機能

### 1. KFM展開機能 (Deploy-OneDriveKFM.ps1)
- OneDriveクライアントの自動アップデート
- KFM（Known Folder Move）の設定展開
- レジストリを使用した各種設定の自動化
  - サイレントサインイン
  - 外部共有の制御
  - ファイルオンデマンド機能の有効化
- 既存データを保持したままの移行対応
- OneDriveプロセスの自動再起動

### 2. 監視機能 (Monitor-OneDriveStatus.ps1)
- OneDriveの同期エラー検出
- ユーザーごとの使用状況レポート生成
  - 氏名
  - ログオンアカウント名
  - メールアドレス
  - アカウント状態
  - ストレージ使用量
  - 残容量
  - 使用率
  - 最終更新日時
- OneDriveクライアントバージョン管理
- ポリシーコンプライアンス監視
- 詳細なログ記録
- CSVフォーマットでのレポート出力

### 3. トラブルシューティング機能 (Troubleshoot-OneDriveKFM.ps1)
- OneDrive接続状態の診断
- KFM設定の検証
- 既知のフォルダー移行状況の確認
  - デスクトップ
  - ドキュメント
  - ピクチャ
- 自動修復機能
- 診断情報の自動収集

### 4. 通知機能 (Send-OneDriveNotification.ps1)
- アラート機能
  - ストレージ使用率の監視（閾値超過通知）
  - エラー・警告数の監視
  - カスタマイズ可能な閾値設定
- 通知方式
  - メール通知（SMTP）
  - Microsoft Teams通知（Webhook）
- 監査ログ
  - 操作ログの自動記録
  - ログの自動ローテーション
  - カスタム保持期間設定

### 5. 自動化機能 (Register-OneDriveScheduler.ps1)
- タスクスケジューラーによる自動実行
  - 監視タスク（1時間ごと）
  - 通知タスク（エラー発生時）
  - 診断タスク（毎日深夜）
  - クリーンアップタスク（毎日早朝）
- エラー処理
  - 自動リトライ機能
  - エラー通知
  - 詳細なログ記録
- メンテナンス機能
  - 古いレポートの自動クリーンアップ
  - ログローテーション
  - 設定のバックアップ

### レポート機能の拡張
- インタラクティブなデータ表示
  - 使用率グラフ
  - エラー統計グラフ
  - データテーブルのソート機能
  - 検索/フィルタリング機能
- 複数のエクスポート形式
  - CSV形式
  - PDF形式
  - 印刷用最適化レイアウト
- 自動更新機能
  - カスタム更新間隔の設定
  - バックグラウンド更新
  - リアルタイムデータ表示

## 動作要件
- Windows 10/11
- PowerShell 5.0以上
- グローバル管理者権限（必須）
- インターネット接続（Microsoft Graph APIアクセス用）
- 必須モジュール（自動インストール）:
  - Microsoft.Graph.Authentication
  - Microsoft.Graph.Users
  - Microsoft.Graph.Reports
  - MSOnline
- OneDrive for Business

## セットアップ手順

### 1. 事前準備
1. グローバル管理者アカウントでサインイン
2. PowerShellの実行ポリシーを確認
```powershell
Get-ExecutionPolicy
```
3. 必要に応じて実行ポリシーを変更
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```
4. MSOnlineモジュールのインストール（監視機能を使用する場合）
```powershell
Install-Module -Name MSOnline -Force -AllowClobber
```

### 2. 初期設定
1. スクリプトをダウンロードし、任意のフォルダに配置
2. `Deploy-OneDriveKFM.ps1`の以下の変数を環境に合わせて設定
   - `$TenantID`: あなたの組織のOneDriveテナントID
   - `$KFMSettings`: KFMの展開設定（必要に応じてカスタマイズ）

### 3. 実行手順

#### KFMの展開
1. PowerShellを管理者として起動
2. 以下のコマンドを実行して展開設定を確認
```powershell
.\Deploy-OneDriveKFM.ps1 -TenantID "<テナントID>" -WhatIf
```

3. 実際の展開を実行（KFMなし）
```powershell
.\Deploy-OneDriveKFM.ps1 -TenantID "<テナントID>"
```

4. KFMを有効にする場合（オプション）
```powershell
# 通知付きでKFMを有効化
.\Deploy-OneDriveKFM.ps1 -TenantID "<テナントID>" -EnableKFM

# 通知なしでKFMを有効化
.\Deploy-OneDriveKFM.ps1 -TenantID "<テナントID>" -EnableKFM -EnableNotification:$false
```

利用可能なパラメータ:
- `-TenantID`: （必須）組織のOneDriveテナントID
- `-WhatIf`: 実際の変更を行わずに実行内容を確認
- `-EnableKFM`: KFM機能を有効化（任意）
- `-EnableNotification`: KFM有効化時のユーザー通知（既定：有効）

#### 監視の開始
1. PowerShellを管理者として起動
2. 以下のコマンドを実行
```powershell
.\Monitor-OneDriveStatus.ps1
```
- オプション: 出力先フォルダの指定
```powershell
.\Monitor-OneDriveStatus.ps1 -OutputPath "C:\Reports" -LogPath "C:\Logs"
```

#### トラブルシューティング
1. PowerShellを管理者として起動
2. 以下のコマンドを実行
```powershell
.\Troubleshoot-OneDriveKFM.ps1
```

### 4. 通知設定
1. config/notification-config.jsonを編集
```json
{
    "Alerts": {
        "StorageThreshold": 90,    // ストレージ使用率閾値（%）
        "ErrorThreshold": 5,       // エラー数閾値
        "WarningThreshold": 10     // 警告数閾値
    },
    "Notification": {
        "Email": {
            "Enabled": true,
            "SmtpServer": "smtp.office365.com",
            "FromAddress": "sender@yourdomain.com",
            "ToAddress": ["recipient@yourdomain.com"]
        },
        "Teams": {
            "Enabled": true,
            "WebhookUrl": "https://outlook.office.com/webhook/..."
        }
    },
    "AuditLog": {
        "Enabled": true,
        "Path": "D:\\OneDriveForBusiness運用ツール\\Logs\\audit",
        "RetentionDays": 90        // ログ保持期間（日数）
    }
}
```

2. 通知機能の実行
```powershell
.\Send-OneDriveNotification.ps1
```

オプションパラメータ:
- `-ConfigPath`: 設定ファイルのパス
- `-TeamsWebhookUrl`: Teams Webhook URL
- `-SmtpServer`: SMTPサーバーアドレス
- `-FromAddress`: 送信元メールアドレス
- `-ToAddress`: 送信先メールアドレス（カンマ区切り）

### 5. レポート機能の利用

#### HTML形式レポートの操作
1. テーブルの検索
   - 検索ボックスに検索語を入力
   - リアルタイムフィルタリング

2. データの並べ替え
   - 列ヘッダーをクリックしてソート
   - 昇順/降順の切り替え

3. グラフ表示
   - 使用率グラフ：ユーザーごとのストレージ使用状況
   - エラー統計：エラー/警告/情報の分布

4. エクスポート
   - CSVダウンロード：データ分析用
   - PDFダウンロード：レポート共有用
   - 印刷：最適化されたレイアウトで印刷

5. 自動更新
   - 更新間隔の選択（1分～1時間）
   - バックグラウンド自動更新
   - 手動更新も可能

### 6. 自動化の設定
1. スケジューラー設定の編集
```json
{
    "Tasks": [
        {
            "Name": "OneDriveMonitoring",
            "Schedule": {
                "Frequency": "Daily",
                "Time": "09:00",
                "Interval": 60
            }
        }
    ]
}
```

2. タスクの登録
```powershell
.\Register-OneDriveScheduler.ps1 -Register
```

3. タスクの削除
```powershell
.\Register-OneDriveScheduler.ps1 -Unregister
```

### 自動化されるタスク
1. 監視タスク（Monitor-OneDriveStatus.ps1）
   - 1時間ごとの状態チェック
   - レポート生成
   - エラー検知

2. 通知タスク（Send-OneDriveNotification.ps1）
   - エラー発生時の即時通知
   - 閾値超過の警告
   - 定期レポートの送信

3. 診断タスク（Troubleshoot-OneDriveKFM.ps1）
   - 毎日のシステム診断
   - 詳細レポートの生成
   - 潜在的な問題の検出

4. クリーンアップタスク（Clean-OneDriveReports.ps1）
   - 古いレポートの削除
   - ログの圧縮・アーカイブ
   - ディスク容量の最適化

### メンテナンス設定
- レポート保持期間: 30日
- ログローテーション: 毎日
- バックアップ数: 5世代
- モジュール更新: 週次チェック

### エラー処理
- 最大リトライ回数: 3回
- リトライ間隔: 5分
- エラー通知: 即時
- エラーログ: 詳細記録

### 監視項目の自動チェック
1. ストレージ使用率
   - 警告閾値: 90%
   - 危険閾値: 95%
   - チェック間隔: 1時間

2. 同期状態
   - エラー閾値: 5件以上
   - 警告閾値: 10件以上
   - チェック間隔: 1時間

3. クライアント状態
   - バージョンチェック
   - プロセス監視
   - 設定検証

4. コンプライアンス
   - ポリシー準拠確認
   - 外部共有制御
   - アクセス権確認

## 出力ファイル

### 監視レポート
- `OneDriveSyncErrors_[日付].csv`: 同期エラーレポート（CSV形式）
- `OneDriveSyncErrors_[日付].html`: 同期エラーレポート（HTML形式）
- `OneDriveUsage_[日付].csv`: 使用状況レポート（CSV形式）
- `OneDriveUsage_[日付].html`: 使用状況レポート（HTML形式）
- `OneDriveCompliance_[日付].csv`: ポリシー準拠状況レポート（CSV形式）
- `OneDriveCompliance_[日付].html`: ポリシー準拠状況レポート（HTML形式）

### 診断情報
- `OneDriveDiagnostics_[日付_時刻].txt`: 詳細な診断情報（テキスト形式）
- `OneDriveDiagnostics_[日付_時刻].html`: 詳細な診断情報（HTML形式）
  - OneDriveバージョン情報
  - プロセス状態
  - レジストリ設定
  - イベントログ（過去24時間）
  - 既知のフォルダー状態
  - グラフィカルなステータス表示
  - エラー・警告の色分け表示

### レポート形式の特徴
- CSV形式: データの二次利用やスプレッドシートでの分析に最適
- HTML形式: 
  - 見やすいレイアウト
  - 色分けされたステータス表示
  - インタラクティブな表形式
  - ブラウザでの閲覧が可能
  - 印刷に適したフォーマット

## 注意事項
- すべてのスクリプトはグローバル管理者権限が必要です
- 初回実行時はMicrosoft Graph認証が必要です
- 初回実行時はテナントIDの設定が必須です
- 大規模展開時はテスト環境での検証を推奨します
- ネットワーク環境によってはOneDriveクライアントの更新に時間がかかる場合があります
- ユーザーデータのバックアップを推奨します

## トラブルシューティング

### よくある問題と解決策
1. スクリプトが実行できない
   - 実行ポリシーの確認と変更
   - 管理者権限での実行確認
2. OneDriveの同期エラー
   - `Troubleshoot-OneDriveKFM.ps1`での診断実行
   - ネットワーク接続の確認
3. KFMが適用されない
   - テナントIDの確認
   - レジストリ設定の確認
   - OneDriveクライアントの再起動

### 通知関連の問題
1. メール通知が届かない
   - SMTPサーバー設定の確認
   - 認証情報の確認
   - ファイアウォール設定の確認

2. Teams通知が届かない
   - Webhook URLの有効性確認
   - チャネルの権限設定確認
   - プロキシ設定の確認

3. 監査ログが生成されない
   - フォルダーのアクセス権確認
   - ディスク容量の確認
   - Logging Enabled設定の確認

## ライセンス
MIT License

## 作者
[みらい建設工業株式会社管理本部総務部システムグループ：有藤健太郎]

## 更新履歴
- v1.0.0 (2025-02-19)
  - 初回リリース
  - 基本機能の実装
  - 監視機能の追加
  - トラブルシューティング機能の実装
- v1.1.0 (2024-02-XX)
  - 通知システムの追加
  - グラフィカルレポート機能の実装
  - PDF出力機能の追加
  - 自動更新機能の実装
  - 監査ログ機能の追加