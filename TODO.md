# 研修カレンダー自動化システム実装 TODOリスト

## 1. 環境設定 (Environment Setup)
- [ ] Google Workspace APIの有効化
    - [ ] Google Calendar API
    - [ ] Google Sheets API
    - [ ] Gmail API (通知用)
- [ ] スプレッドシートの準備
    - [ ] **入社者連携シート**の作成
        - `入社者氏名`
        - `入社日`
        - `メールアドレス`
        - `役職レベル`
        - `処理状況`
    - [ ] **ONB管理表 (研修マスターシート)** の作成
        - `研修名称`
        - `研修対象者（役職）`
        - `実施タイミング`
        - `担当者・講師`
        - `講師メールアドレス`
        - `会議室名・リソースID`
    - [ ] **設定・ログシート**の作成
        - `実行日時`
        - `処理結果`
        - `詳細メッセージ`

## 2. GASスクリプト実装 (GAS Script Implementation)
- [ ] **メイン処理 (`executeONBAutomation`)**
    - [ ] `executeONBAutomation` 関数の基本構造を実装
    - [ ] 入社者連携シートから未処理のデータを読み込むロジックを実装
    - [ ] 各入社者に対して処理を行うループを実装
- [ ] **ユーティリティ関数 (Utility Functions)**
    - [ ] `getRequiredTrainings(role)`: ONB管理表から役職に応じた研修リスト（講師情報含む）を取得する関数を実装
    - [ ] `createCalendarInvitation(employee, training)`: カレンダーイベントを作成し、入社者、講師を招待し、会議室を予約する関数を実装
    - [ ] `updateStatus(row, status)`: 入社者連携シートのステータスを更新する関数を実装
    - [ ] `logExecution(status, message)`: ログシートに実行結果を記録する関数を実装
- [ ] **エラーハンドリング (Error Handling)**
    - [ ] メイン処理に `try-catch` ブロックを実装し、各入社者の処理を個別に行う
    - [ ] `handleError(error, employee)`: エラー内容をログに記録し、該当者のステータスを「エラー」に更新する関数を実装
    - [ ] エラー発生時にGmailで担当者に通知する機能（オプション）
- [ ] **UI連携 (UI Integration)**
    - [ ] スプレッドシートに「実行」ボタンを配置
    - [ ] ボタンに `executeONBAutomation` 関数を割り当てる

## 3. セキュリティと権限設定
- [ ] スプレッドシートの共有設定を行い、アクセスを関係者に限定
- [ ] GASプロジェクトの `appsscript.json` に必要なOAuthスコープを定義
    - `https://www.googleapis.com/auth/calendar`
    - `https://www.googleapis.com/auth/spreadsheets`
    - `https://www.googleapis.com/auth/gmail.send` (通知機能実装時)

## 4. 運用・保守
- [ ] 運用マニュアルの作成
    - [ ] 処理の実行手順
    - [ ] ログの確認方法
    - [ ] エラー発生時の対応手順
- [ ] マスターデータ（ONB管理表）のメンテナンス手順を定義
- [ ] システム停止時の代替手動手順を文書化

## 5. テスト
- [ ] **単体テスト**
    - [ ] 各ユーティリティ関数が正しく動作することを確認
- [ ] **結合テスト**
    - [ ] **正常系**: 複数役職の入社者データで、カレンダー招待とステータス更新が正しく行われることを確認
    - [ ] **異常系**: 不正なデータ（存在しない役職、不正なメールアドレス等）でエラーが記録され、他の処理が続行されることを確認

## 6. 今後の拡張（オプション）
- [ ] 人事システムとのAPI連携の調査・検討
- [ ] 研修リマインダーメールの自動送信機能の実装
- [ ] Slackなどチャットツールへの通知機能の実装 