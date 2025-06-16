// =========================================
// UI（メニュー）作成
// =========================================

/**
 * スプレッドシートを開いた時にカスタムメニューを追加する
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('研修自動化')
    .addItem('📋 実行ログを表示', 'showExecutionLogs')
    .addSeparator()
    .addItem('カレンダー招待を実行（従来版）', 'executeONBAutomation')
    .addSeparator()
    .addItem('全カレンダーイベントを削除', 'deleteAllCalendarEvents')
    .addItem('特定研修のイベントを削除', 'deleteSpecificEvent')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('テスト機能')
      .addItem('📊 マッピングシート作成デバッグ', 'debugMappingSheet')
      .addSeparator()
      .addItem('複数講師対応テスト', 'テスト_複数講師対応')
      .addItem('実施順処理テスト', 'テスト_実施順処理')
      .addItem('入社者データ取得テスト', 'テスト_入社者データ取得')
      .addItem('会議室予約管理テスト', 'テスト_会議室予約管理')
      .addItem('時間枠計算テスト', 'テスト_時間枠計算')
      .addItem('拡張検証テスト', 'テスト_拡張検証')
      .addSeparator()
      .addItem('🔍 カレンダー重複問題検証', 'テスト_カレンダー重複問題検証'))
    .addToUi();
}

// =========================================
// メイン処理
// =========================================

/**
 * 入社日取得のテスト関数
 */
function testHireDateRetrieval() {
  var executionSheet = SpreadsheetApp.openById(SPREADSHEET_IDS.EXECUTION).getSheetByName(SHEET_NAMES.EXECUTION);
  var hireDateValue = executionSheet.getRange('B3').getValue(); // B列3行目から取得（入社日）
  
  console.log('B3セルの値:', hireDateValue);
  console.log('B3セルの型:', typeof hireDateValue);
  console.log('B3セルがDate?:', hireDateValue instanceof Date);
  
  // 直接アラートでも確認
  SpreadsheetApp.getUi().alert('B3（入社日）: ' + hireDateValue + ' (型: ' + typeof hireDateValue + ')');
}

/**
 * 実行ボタンから呼び出されるメイン関数
 */
function executeONBAutomation() {
  var ui = SpreadsheetApp.getUi();
  var executionSheet = SpreadsheetApp.openById(SPREADSHEET_IDS.EXECUTION).getSheetByName(SHEET_NAMES.EXECUTION);
  var hireDateValue = executionSheet.getRange('B3').getValue(); // B列3行目から取得（入社日）
  
  // 日付の型を確認し、適切にDateオブジェクトに変換
  function convertToDate(value) {
    writeLog('DEBUG', '日付変換: 入力値=' + value + ', 型=' + typeof value);
    
    if (value instanceof Date) {
      writeLog('DEBUG', '既にDateオブジェクト: ' + value);
      return value;
    } else if (typeof value === 'number') {
      // Googleスプレッドシートの日付シリアル値の場合
      writeLog('DEBUG', '数値から日付変換: ' + value);
      var date = new Date(value);
      writeLog('DEBUG', '変換結果: ' + date);
      return date;
    } else if (typeof value === 'string') {
      writeLog('DEBUG', '文字列から日付変換: ' + value);
      
      // 「エラー」などの明らかに日付でない文字列をチェック
      if (value === 'エラー' || value === 'ERROR' || value === '#VALUE!' || value === '#REF!' || value === '#NAME?') {
        writeLog('WARN', 'セルにエラー値または無効な文字列が入力されています: ' + value);
        return null;
      }
      
      // 様々な日付形式を試行
      var formats = [
        value,
        value.replace(/\//g, '-'),
        value.replace(/-/g, '/'),
      ];
      
      for (var i = 0; i < formats.length; i++) {
        var parsed = new Date(formats[i]);
        if (!isNaN(parsed.getTime())) {
          writeLog('DEBUG', '変換成功: ' + formats[i] + ' -> ' + parsed);
          return parsed;
        }
      }
      writeLog('WARN', '文字列の日付変換に失敗: ' + value + ' (試行フォーマット: ' + formats.join(', ') + ')');
      return null;
    } else {
      writeLog('DEBUG', '不明な型: ' + typeof value);
      return null;
    }
  }
  
  var hireDate = convertToDate(hireDateValue);
  
  var executionParams = {
    sheet: executionSheet,
    user: Session.getActiveUser().getEmail(),
    hireDate: hireDate,
  };

  try {
    writeLog('INFO', '処理開始: 研修カレンダー自動化');
    writeLog('INFO', '実行ユーザー: ' + executionParams.user);
    writeLog('DEBUG', '入社日取得値: ' + hireDateValue + ' (型: ' + typeof hireDateValue + ')');
    writeLog('DEBUG', '変換後入社日: ' + executionParams.hireDate + ' (型: ' + typeof executionParams.hireDate + ')');
    
    if (executionParams.hireDate && 
        executionParams.hireDate instanceof Date &&
        !isNaN(executionParams.hireDate.getTime())) {
      writeLog('INFO', '入社日基準: ' + Utilities.formatDate(executionParams.hireDate, 'Asia/Tokyo', 'yyyy年MM月dd日'));
    } else {
      writeLog('ERROR', '入社日の変換に失敗しました');
      writeLog('ERROR', 'B3セル（入社日）: ' + hireDateValue + ' → 変換結果: ' + (hireDate === null ? 'null' : hireDate));
      
      var errorDetails = [];
      if (!hireDate || !(hireDate instanceof Date) || isNaN(hireDate.getTime())) {
        errorDetails.push('B3セル（入社日）に正しい日付を入力してください。現在の値: ' + hireDateValue);
      }
      
      throw new Error('入社日エラー:\n' + errorDetails.join('\n'));
    }
    
    // 1. 未処理の入社者リストを取得（入社日でフィルタリング）
    var newHiresData = getNewHires(executionParams.hireDate);
    if (newHiresData.length === 0) {
      var infoMsg = '処理対象の入社者が見つかりませんでした。処理をスキップして正常終了します。';
      writeLog('INFO', infoMsg);
      ui.alert('情報', infoMsg, ui.ButtonSet.OK);
      logExecution(executionParams, '成功', infoMsg);
      return; // 正常終了
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

    // 3.1. 研修期間を計算（入社日から最大実施日まで）
    var periodStart = new Date(executionParams.hireDate);
    var periodEnd = new Date(executionParams.hireDate);
    var maxImplementationDay = 1; // デフォルト値
    
    if (trainingGroups.length > 0) {
      // 研修グループから最大実施日を取得
      for (var i = 0; i < trainingGroups.length; i++) {
        var group = trainingGroups[i];
        if (group.implementationDay && group.implementationDay > maxImplementationDay) {
          maxImplementationDay = group.implementationDay;
        }
      }
      
      // 最大実施日に基づいて期間終了日を計算（営業日ベース）
      var daysAdded = 0;
      var tempDate = new Date(periodStart);
      while (daysAdded < maxImplementationDay - 1) {
        tempDate.setDate(tempDate.getDate() + 1);
        if (tempDate.getDay() !== 0 && tempDate.getDay() !== 6) { // 平日のみ
          daysAdded++;
        }
      }
      periodEnd = new Date(tempDate);
      
      writeLog('INFO', '研修期間: ' + Utilities.formatDate(periodStart, 'Asia/Tokyo', 'yyyy/MM/dd') + 
               ' - ' + Utilities.formatDate(periodEnd, 'Asia/Tokyo', 'yyyy/MM/dd') + 
               ' (最大実施日: ' + maxImplementationDay + '営業日目)');
    }

    // 3.5. マッピング結果をシートに表示（プログレス付きインクリメンタル処理）
    if (trainingGroups.length > 0) {
      var mappingSheet = createIncrementalMappingSheet(trainingGroups, newHiresData, periodStart, periodEnd);
      
      // 4. カレンダーイベントを作成（インクリメンタル処理、リファクタリング版）
      writeLog('INFO', 'インクリメンタルカレンダーイベント作成を開始');
      writeLog('INFO', '入社日基準: ' + executionParams.hireDate);
      processTrainingGroupsIncrementally(trainingGroups, newHiresData, executionParams.hireDate, mappingSheet);
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

// =========================================
// カレンダー削除関連のUI操作
// =========================================

/**
 * 全カレンダーイベントを削除するメニュー関数
 */
function deleteAllCalendarEvents() {
  var ui = SpreadsheetApp.getUi();
  
  // 確認ダイアログを表示
  var response = ui.alert(
    '確認', 
    '最新のマッピングシートに記録されているすべてのカレンダーイベントを削除します。\nこの操作は取り消せません。実行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    ui.alert('操作がキャンセルされました。');
    return;
  }
  
  try {
    writeLog('INFO', '全カレンダーイベント削除操作開始: 実行者=' + Session.getActiveUser().getEmail());
    
    var result = deleteCalendarEventsFromMappingSheet();
    
    var message = '削除完了:\n' +
                  '・対象シート: ' + result.sheetName + '\n' +
                  '・成功: ' + result.success + '件\n' +
                  '・失敗: ' + result.failed + '件\n' +
                  '・総数: ' + result.total + '件';
    
    if (result.errors.length > 0) {
      message += '\n\nエラー詳細:\n' + result.errors.join('\n');
    }
    
    writeLog('INFO', '全カレンダーイベント削除完了: ' + JSON.stringify(result));
    ui.alert('削除完了', message, ui.ButtonSet.OK);
    
  } catch (e) {
    writeLog('ERROR', '全カレンダーイベント削除でエラー: ' + e.message);
    ui.alert('エラー', 'カレンダーイベントの削除中にエラーが発生しました:\n' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * 特定研修のイベントを削除するメニュー関数
 */
function deleteSpecificEvent() {
  var ui = SpreadsheetApp.getUi();
  
  // 研修名を入力
  var response = ui.prompt(
    '特定研修の削除',
    '削除する研修名を入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    ui.alert('操作がキャンセルされました。');
    return;
  }
  
  var trainingName = response.getResponseText().trim();
  if (!trainingName) {
    ui.alert('エラー', '研修名が入力されていません。', ui.ButtonSet.OK);
    return;
  }
  
  // 確認ダイアログを表示
  var confirmResponse = ui.alert(
    '確認', 
    '研修「' + trainingName + '」のカレンダーイベントを削除します。\nこの操作は取り消せません。実行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (confirmResponse !== ui.Button.YES) {
    ui.alert('操作がキャンセルされました。');
    return;
  }
  
  try {
    writeLog('INFO', '特定研修カレンダーイベント削除操作開始: 研修名=' + trainingName + ', 実行者=' + Session.getActiveUser().getEmail());
    
    var success = deleteSpecificTrainingEvent(trainingName);
    
    if (success) {
      writeLog('INFO', '特定研修カレンダーイベント削除成功: ' + trainingName);
      ui.alert('削除完了', '研修「' + trainingName + '」のカレンダーイベントを削除しました。', ui.ButtonSet.OK);
    } else {
      writeLog('WARN', '特定研修カレンダーイベント削除失敗: ' + trainingName);
      ui.alert('削除失敗', '研修「' + trainingName + '」のカレンダーイベントの削除に失敗しました。\n研修名が正確か、またはカレンダーIDが設定されているかを確認してください。', ui.ButtonSet.OK);
    }
    
  } catch (e) {
    writeLog('ERROR', '特定研修カレンダーイベント削除でエラー: ' + e.message + ' (研修名: ' + trainingName + ')');
    ui.alert('エラー', 'カレンダーイベントの削除中にエラーが発生しました:\n' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * マッピングシート作成のデバッグ関数
 */
function debugMappingSheet() {
  try {
    writeLog('INFO', '=== マッピングシート作成デバッグ開始 ===');
    
    // デバッグ用：入社日フィルタなし
    var newHires = getNewHires();
    if (newHires.length === 0) {
      SpreadsheetApp.getUi().alert('警告', '入社者データがありません。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    var trainingGroups = groupTrainingsForHires(newHires);
    
    var message = 'マッピングシート作成デバッグ結果:\n\n';
    message += '入社者数: ' + newHires.length + '名\n';
    message += '研修グループ数: ' + trainingGroups.length + '件\n\n';
    
    message += '入社者詳細:\n';
    for (var i = 0; i < Math.min(3, newHires.length); i++) {
      var hire = newHires[i];
      message += (i + 1) + '. ' + hire.name + ' (' + hire.rank + '/' + hire.experience + '/' + hire.department + ')\n';
    }
    if (newHires.length > 3) {
      message += '... 他' + (newHires.length - 3) + '名\n';
    }
    
    message += '\n研修グループ詳細:\n';
    for (var i = 0; i < Math.min(3, trainingGroups.length); i++) {
      var group = trainingGroups[i];
      message += (i + 1) + '. ' + group.name + ' (参加者: ' + (group.attendees ? group.attendees.length : 0) + '名)\n';
    }
    if (trainingGroups.length > 3) {
      message += '... 他' + (trainingGroups.length - 3) + '件\n';
    }
    
    // 今日の日付を使用してマッピングシート作成
    var hireDate = new Date();
    createMappingSheet(trainingGroups, newHires, hireDate, hireDate);
    
    message += '\nマッピングシートを作成しました。\n';
    message += '詳細なログは実行ログシートをご確認ください。';
    
    SpreadsheetApp.getUi().alert('デバッグ結果', message, SpreadsheetApp.getUi().ButtonSet.OK);
    writeLog('INFO', '=== マッピングシート作成デバッグ完了 ===');
    
  } catch (e) {
    writeLog('ERROR', 'マッピングシート作成デバッグでエラー: ' + e.message);
    SpreadsheetApp.getUi().alert('エラー', 'マッピングシート作成デバッグでエラーが発生しました:\n' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}