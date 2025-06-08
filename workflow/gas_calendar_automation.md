# GASを利用した研修カレンダー自動化システム

## 1. システム概要

### 1.1 目的
- Google Apps Script (GAS)を活用した研修カレンダー招待の自動化
- 人事担当者の作業負荷軽減
- 処理状況の可視化と管理効率の向上

### 1.2 基本機能
1. スプレッドシートでの入社者情報管理
2. 定期的な研修カレンダー招待の自動送信
3. 処理状況の自動更新と管理
4. エラー発生時の通知と記録

## 2. 前提条件と環境設定

### 2.1 Google Workspace設定
1. Google Workspaceアカウントの準備
2. 必要なサービスの有効化
   - Google Calendar
   - Google Spreadsheet
   - Google Apps Script

### 2.2 スプレッドシート構成
1. 入社者リストシート
   - 入社日
   - 氏名
   - メールアドレス
   - 処理状況（未処理/招待済み/エラー）

2. 研修マスターシート
   - 研修名
   - 所要時間
   - 実施タイミング（入社日からの日数）
   - 必須/任意区分

## 3. システム構成

### 3.1 アーキテクチャ構成
```mermaid
graph TB
    subgraph "データ管理"
        A[入社者リスト<br>スプレッドシート]
        B[研修マスター<br>スプレッドシート]
        C[招待ログ<br>スプレッドシート]
    end

    subgraph "処理エンジン"
        D[Google Apps Script]
        E[Time-Driven Trigger]
    end

    subgraph "外部サービス"
        F[Google Calendar]
        G[Gmail]
    end

    E -->|定期実行| D
    D -->|読取/書込| A
    D -->|読取| B
    D -->|書込| C
    D -->|イベント作成| F
    D -->|メール送信| G
```

### 3.2 データフロー

#### 3.2.1 システム処理フロー
```mermaid
sequenceDiagram
    participant System as "システム"
    participant List as "入社者リスト"
    participant Training as "研修マスター"
    participant Calendar as "カレンダー"
    participant Log as "招待ログ"

    Note over System: 毎月1日 or 15日の指定時刻
    System->>List: 入社者リストをチェック
    List-->>System: 本日入社の対象者情報
    System->>Training: 実施すべき全研修リストを取得
    Training-->>System: 研修リスト

    loop 対象者ごとの処理
        System->>System: 対象者を選択
        loop 研修ごとの処理
            System->>Log: 招待ログを確認
            Log-->>System: 招待状況
            alt 未招待の場合
                System->>Calendar: カレンダーにイベント作成
                System->>System: 招待メール送信
                System->>Log: 招待ログに記録
            end
        end
        System->>List: 処理状況を「招待済み」に更新
    end

    Note over System: エラー発生時は入社者リストの処理状況にエラー内容を記録
```

#### 3.2.2 人事担当者の運用フロー
```mermaid
sequenceDiagram
    participant HR as "人事担当者"
    participant Sheet as "スプレッドシート"
    participant System as "システム"
    participant List as "入社者リスト"

    Note over HR: 新しい入社者が決定
    HR->>Sheet: スプレッドシートを開く
    HR->>Sheet: 入社者情報を追記<br>(入社日、氏名、メールアドレス)
    HR->>Sheet: スプレッドシートを閉じる

    Note over System: 毎月1日と15日の朝9時
    System->>System: スクリプト自動起動
    System->>List: カレンダー招待を一括送信

    Note right of HR: 任意の確認作業
    HR->>List: 処理状況が「招待済み」になっているか確認
```

## 4. 実装仕様

### 4.1 GASスクリプト構成
1. メイン処理
   ```javascript
   function processNewEmployees() {
     // 新入社員の研修カレンダー招待処理
   }
   ```

2. トリガー設定
   ```javascript
   function setTrigger() {
     // 毎月1日と15日の朝9時に実行
   }
   ```

3. ユーティリティ関数
   ```javascript
   function createCalendarEvent() {
     // カレンダーイベント作成
   }
   
   function updateStatus() {
     // 処理状況の更新
   }
   
   function sendErrorNotification() {
     // エラー通知の送信
   }
   ```

### 4.2 エラーハンドリング
- スプレッドシートアクセスエラー
- カレンダーAPI制限エラー
- メール送信エラー
- データ不整合エラー

## 5. セキュリティ設定

### 5.1 アクセス制御
1. スプレッドシートの共有設定
2. スクリプトの実行権限
3. カレンダーのアクセス権限

### 5.2 データ保護
1. 個人情報の取り扱い
2. エラーログの管理
3. 実行ログの保管

## 6. 運用管理

### 6.1 モニタリング
1. 実行ログの確認方法
2. エラー通知の設定
3. 処理状況の確認手順

### 6.2 メンテナンス
1. スクリプトの定期的な見直し
2. マスターデータの更新手順
3. バックアップ方法

## 7. 今後の拡張性
1. 機能拡張
   - 研修の出欠管理機能
   - リマインダー機能
   - 研修資料の自動配布

2. インテグレーション
   - Slack連携
   - 社内システムとの連携
   - レポート機能の強化 