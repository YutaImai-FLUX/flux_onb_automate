// =========================================
// 定数設定
// =========================================

// 各スプレッドシートのID
var SPREADSHEET_IDS = {
  EXECUTION:   '1o3JlRORxDgE6Hv2NyohB2uMRRlr1iSsOA8czmylUGxE', // DXS本部_ONB研修自動化_実行ファイル
  NEW_HIRES:   '1BOfOUnIHotCLatJYKxDNNfndtptlC8NjoU7NoO2niAQ', // DXS本部_入社者連携シート
  TRAINING_MASTER: '1YNmiNiqSe7ctkHkW3qMOK7K13iodMixKs5WuDrMXsO8', // DXS本部_ONB研修マスタ
  ROOM_MASTER:   '16xML66Ywi8Q5oFf8e5NEHdJrQPwkjUcJdswwXPxtnCQ', // FLUX_六本木オフィス会議室マスタ
};

// シート名
var SHEET_NAMES = {
  EXECUTION:   '実行シート',
  NEW_HIRES:   '入社者リスト',
  TRAINING_MASTER: '研修管理表マスタ',
  ROOM_MASTER:   '会議室マスタ',
};

// 通知メールの送信先
var NOTIFICATION_EMAIL = Session.getActiveUser().getEmail(); 