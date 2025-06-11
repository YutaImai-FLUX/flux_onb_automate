// =========================================
// UI（メニュー）作成
// =========================================

/**
 * スプレッドシートを開いた時にカスタムメニューを追加する
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('研修自動化')
    .addItem('カレンダー招待を実行', 'executeONBAutomation')
    .addToUi();
}


// =========================================
// メイン処理
// =========================================

/**
 * 実行ボタンから呼び出されるメイン関数
 */
function executeONBAutomation() {
  var ui = SpreadsheetApp.getUi();
  var executionSheet = SpreadsheetApp.openById(SPREADSHEET_IDS.EXECUTION).getSheetByName(SHEET_NAMES.EXECUTION);
  var executionParams = {
    sheet: executionSheet,
    user: Session.getActiveUser().getEmail(),
    startDate: executionSheet.getRange('C2').getValue(), // C列2行目から取得
    endDate: executionSheet.getRange('D2').getValue(),   // D列2行目から取得
  };

  try {
    writeLog('INFO', '処理開始: 研修カレンダー自動化');
    writeLog('INFO', '実行ユーザー: ' + executionParams.user);
    writeLog('INFO', 'カレンダー招待期間: ' + Utilities.formatDate(executionParams.startDate, 'Asia/Tokyo', 'yyyy年MM月dd日') + 
             ' から ' + Utilities.formatDate(executionParams.endDate, 'Asia/Tokyo', 'yyyy年MM月dd日') + ' まで');
    
    // 1. 未処理の入社者リストを取得
    var newHiresData = getNewHires();
    if (newHiresData.length === 0) {
      writeLog('INFO', '処理対象の入社者が見つかりませんでした');
      ui.alert('処理対象の入社者がいません。');
      logExecution(executionParams, '完了', '処理対象の入社者が見つかりませんでした。');
      return;
    }
    
    // 対象者情報をログに出力
    logTargetUsers(newHiresData);
    
    // 2. 入社者データの必須項目を検証
    writeLog('INFO', '入社者データの検証を開始');
    validateNewHires(newHiresData);
    writeLog('INFO', '入社者データの検証が完了');

    // 3. 研修情報をグループ化
    writeLog('INFO', '研修情報のグループ化を開始');
    var trainingGroups = groupTrainingsForHires(newHiresData);
    writeLog('INFO', '研修グループ数: ' + trainingGroups.length);

    // 3.5. マッピング結果をシートに表示
    if (trainingGroups.length > 0) {
      createMappingSheet(trainingGroups, newHiresData, executionParams.startDate, executionParams.endDate);
    }

    // 4. カレンダーイベントを作成
    writeLog('INFO', 'カレンダーイベント作成を開始');
    writeLog('INFO', '指定期間: ' + executionParams.startDate + ' から ' + executionParams.endDate);
    for (var i = 0; i < trainingGroups.length; i++) {
      var group = trainingGroups[i];
      writeLog('INFO', '研修: ' + group.name + ' (参加者数: ' + group.attendees.length + '名, 講師: ' + (group.lecturer || 'なし') + ')');
      
      var roomName = null;
      if (group.needsRoom) {
        roomName = findAvailableRoomName(group.attendees.length);
        writeLog('INFO', '会議室確保: ' + roomName);
      }
      createCalendarEvent(group, roomName, executionParams.startDate, executionParams.endDate);
      writeLog('INFO', 'カレンダーイベント作成完了: ' + group.name);
    }

    // 5. 処理ステータスを更新
    updateStatuses(newHiresData);

    // 6. 成功ログ記録とメール通知
    var successMessage = '処理が正常に完了しました。対象者: ' + newHiresData.length + '名';
    writeLog('INFO', '処理完了: ' + successMessage);
    logExecution(executionParams, '成功', successMessage);
    sendNotificationEmail('【成功】研修カレンダー招待処理', successMessage);
    ui.alert(successMessage);

  } catch (e) {
    // 7. エラー発生時のログ記録とメール通知
    logError(e, 'executeONBAutomation', { 
      user: executionParams.user, 
      targetCount: newHiresData ? newHiresData.length : 0 
    });
    
    var errorMessage = 'エラーが発生しました: ' + e.message + '\nスタックトレース: ' + e.stack;
    logExecution(executionParams, 'エラー', errorMessage);
    sendNotificationEmail('【エラー】研修カレンダー招待処理', errorMessage);
    ui.alert(errorMessage);
  }
} 