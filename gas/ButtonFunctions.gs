// =========================================
// ボタン関数とテスト関数
// =========================================

/**
 * 複数講師対応のテスト関数
 */
function テスト_複数講師対応() {
  try {
    writeLog('INFO', '=== 複数講師対応テスト開始 ===');
    
    // 入社者データ取得（テスト用：入社日フィルタなし）
    var newHires = getNewHires();
    if (newHires.length === 0) {
      SpreadsheetApp.getUi().alert('警告', '入社者データがありません。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // 研修グループ作成
    var trainingGroups = groupTrainingsForHires(newHires);
    
    var message = '複数講師対応テスト結果:\n\n';
    var multiLecturerCount = 0;
    
    for (var i = 0; i < Math.min(10, trainingGroups.length); i++) {
      var group = trainingGroups[i];
      var lecturerInfo = '';
      
      if (group.lecturerEmails && group.lecturerEmails.length > 1) {
        multiLecturerCount++;
        lecturerInfo = '複数講師: ' + group.lecturerNames.join(', ') + ' (' + group.lecturerEmails.length + '名)';
      } else if (group.lecturerNames && group.lecturerNames.length > 0) {
        lecturerInfo = '単一講師: ' + group.lecturerNames[0];
      } else {
        lecturerInfo = '講師未設定';
      }
      
      message += (i + 1) + '. ' + group.name + '\n';
      message += '   ' + lecturerInfo + '\n\n';
    }
    
    if (trainingGroups.length > 10) {
      message += '... 他' + (trainingGroups.length - 10) + '件\n\n';
    }
    
    message += '複数講師研修数: ' + multiLecturerCount + '/' + trainingGroups.length + '\n\n';
    message += '詳細なログは実行ログシートをご確認ください。';
    
    SpreadsheetApp.getUi().alert('テスト結果', message, SpreadsheetApp.getUi().ButtonSet.OK);
    writeLog('INFO', '=== 複数講師対応テスト完了 ===');
    
  } catch (e) {
    writeLog('ERROR', '複数講師対応テストでエラー: ' + e.message);
    SpreadsheetApp.getUi().alert('エラー', '複数講師対応テストでエラーが発生しました:\n' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 実施順処理のテスト関数
 */
function テスト_実施順処理() {
  try {
    testSequenceProcessing();
  } catch (e) {
    writeLog('ERROR', '実施順処理テストでエラー: ' + e.message);
    SpreadsheetApp.getUi().alert('エラー', '実施順処理テストでエラーが発生しました:\n' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 入社者データ取得のテスト関数
 */
function テスト_入社者データ取得() {
  try {
    writeLog('INFO', '=== 入社者データ取得テスト開始 ===');
    
    // テスト用：入社日フィルタなし
    var newHires = getNewHires();
    
    var message = '入社者データ取得テスト結果:\n\n';
    message += '対象入社者数: ' + newHires.length + '名\n\n';
    
    for (var i = 0; i < Math.min(5, newHires.length); i++) {
      var hire = newHires[i];
      message += (i + 1) + '. ' + hire.name + '\n';
      message += '   職位: ' + hire.rank + '\n';
      message += '   経験: ' + hire.experience + '\n';
      message += '   メール: ' + hire.email + '\n';
      message += '   行番号: ' + hire.rowNum + '\n\n';
    }
    
    if (newHires.length > 5) {
      message += '... 他' + (newHires.length - 5) + '名\n\n';
    }
    
    message += '詳細なログは実行ログシートをご確認ください。';
    
    SpreadsheetApp.getUi().alert('テスト結果', message, SpreadsheetApp.getUi().ButtonSet.OK);
    writeLog('INFO', '=== 入社者データ取得テスト完了 ===');
    
  } catch (e) {
    writeLog('ERROR', '入社者データ取得テストでエラー: ' + e.message);
    SpreadsheetApp.getUi().alert('エラー', '入社者データ取得テストでエラーが発生しました:\n' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 会議室予約管理システムのテスト関数
 */
function テスト_会議室予約管理() {
  try {
    writeLog('INFO', '=== 会議室予約管理システムテスト開始 ===');
    
    var roomManager = RoomReservationManager.getInstance();
    roomManager.reset();
    
    var testDate = new Date();
    testDate.setHours(9, 0, 0, 0); // 9:00
    var testEndDate = new Date(testDate.getTime() + 60 * 60 * 1000); // 10:00
    
    // テスト1: 同じ部屋の異なる時間帯
    var success1 = roomManager.reserveRoom('TEST_ROOM', 'test@example.com', testDate, testEndDate, '研修A');
    var testDate2 = new Date(testDate.getTime() + 2 * 60 * 60 * 1000); // 11:00
    var testEndDate2 = new Date(testDate2.getTime() + 60 * 60 * 1000); // 12:00
    var success2 = roomManager.reserveRoom('TEST_ROOM', 'test@example.com', testDate2, testEndDate2, '研修B');
    
    // テスト2: 同じ部屋の重複時間帯（失敗するはず）
    var testDate3 = new Date(testDate.getTime() + 30 * 60 * 1000); // 9:30
    var testEndDate3 = new Date(testDate3.getTime() + 60 * 60 * 1000); // 10:30
    var success3 = roomManager.reserveRoom('TEST_ROOM', 'test@example.com', testDate3, testEndDate3, '研修C');
    
    roomManager.logCurrentReservations();
    
    var message = '会議室予約管理システムテスト結果:\n\n';
    message += '予約1 (9:00-10:00): ' + (success1 ? '成功' : '失敗') + '\n';
    message += '予約2 (11:00-12:00): ' + (success2 ? '成功' : '失敗') + '\n';
    message += '予約3 (9:30-10:30, 重複): ' + (success3 ? '成功(異常!)' : '失敗(正常)') + '\n\n';
    message += '現在の予約数: ' + roomManager.getReservations().length + '\n\n';
    message += '詳細なログは実行ログシートをご確認ください。';
    
    SpreadsheetApp.getUi().alert('テスト結果', message, SpreadsheetApp.getUi().ButtonSet.OK);
    writeLog('INFO', '=== 会議室予約管理システムテスト完了 ===');
    
  } catch (e) {
    writeLog('ERROR', '会議室予約管理システムテストでエラー: ' + e.message);
    SpreadsheetApp.getUi().alert('エラー', '会議室予約管理システムテストでエラーが発生しました:\n' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 時間枠計算テスト関数
 */
function テスト_時間枠計算() {
  try {
    writeLog('INFO', '=== 時間枠計算テスト開始 ===');
    
    // テスト用の入社日（今日）
    var hireDate = new Date();
    hireDate.setHours(9, 0, 0, 0);
    
    // テスト用の研修グループを作成
    var testGroups = [
      {
        name: '【DX オンボ】 ONBキックオフ（C/SC）',
        implementationDay: 1,
        sequence: 1,
        time: '60分',
        needsRoom: true,
        attendees: ['test1@example.com', 'test2@example.com']
      },
      {
        name: '【DX オンボ】 ONBキックオフ（Mup）',
        implementationDay: 1,
        sequence: 2,
        time: '60分',
        needsRoom: true,
        attendees: ['test3@example.com', 'test4@example.com']
      },
      {
        name: '【DX オンボ】 ツール・社内ルール',
        implementationDay: 1,
        sequence: 3,
        time: '30分',
        needsRoom: true,
        attendees: ['test1@example.com', 'test2@example.com', 'test3@example.com']
      },
      {
        name: '【DX オンボ】 【各自視聴】週間ゴール設計',
        implementationDay: 1,
        sequence: 4,
        time: '30分',
        needsRoom: false,
        attendees: ['test1@example.com', 'test2@example.com']
      }
    ];
    
    var message = '時間枠計算テスト結果:\n\n';
    message += '入社日: ' + Utilities.formatDate(hireDate, 'Asia/Tokyo', 'yyyy/MM/dd') + '\n\n';
    
    // 会議室予約管理をリセット
    var roomManager = RoomReservationManager.getInstance();
    roomManager.reset();
    
    var scheduledEvents = [];
    
    for (var i = 0; i < testGroups.length; i++) {
      var group = testGroups[i];
      writeLog('INFO', '時間枠計算テスト: ' + group.name);
      
      var timeSlot = findAvailableTimeSlot(group, hireDate);
      
      if (timeSlot) {
        scheduledEvents.push({
          name: group.name,
          startTime: timeSlot.start,
          endTime: timeSlot.end,
          attendees: group.attendees
        });
        
        message += (i + 1) + '. ' + group.name + '\n';
        message += '   予定時間: ' + Utilities.formatDate(timeSlot.start, 'Asia/Tokyo', 'MM/dd HH:mm') + 
                   '-' + Utilities.formatDate(timeSlot.end, 'Asia/Tokyo', 'HH:mm') + '\n';
        message += '   実施日: ' + group.implementationDay + '営業日目, 実施順: ' + group.sequence + '\n\n';
        
        writeLog('INFO', '時間枠確保成功: ' + group.name + ' → ' + 
                 Utilities.formatDate(timeSlot.start, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + 
                 Utilities.formatDate(timeSlot.end, 'Asia/Tokyo', 'HH:mm'));
      } else {
        message += (i + 1) + '. ' + group.name + '\n';
        message += '   予定時間: 確保失敗\n\n';
        writeLog('ERROR', '時間枠確保失敗: ' + group.name);
      }
    }
    
    // 重複チェック
    var conflicts = [];
    for (var i = 0; i < scheduledEvents.length; i++) {
      for (var j = i + 1; j < scheduledEvents.length; j++) {
        var event1 = scheduledEvents[i];
        var event2 = scheduledEvents[j];
        
        // 時間重複チェック
        if (!(event1.endTime <= event2.startTime || event1.startTime >= event2.endTime)) {
          conflicts.push(event1.name + ' と ' + event2.name);
        }
      }
    }
    
    if (conflicts.length > 0) {
      message += '⚠️ 時間重複検出:\n';
      for (var i = 0; i < conflicts.length; i++) {
        message += '  ' + conflicts[i] + '\n';
      }
    } else {
      message += '✅ 時間重複なし - 正常に配置されました\n';
    }
    
    message += '\n詳細なログは実行ログシートをご確認ください。';
    
    SpreadsheetApp.getUi().alert('テスト結果', message, SpreadsheetApp.getUi().ButtonSet.OK);
    writeLog('INFO', '=== 時間枠計算テスト完了 ===');
    
  } catch (e) {
    writeLog('ERROR', '時間枠計算テストでエラー: ' + e.message);
    SpreadsheetApp.getUi().alert('エラー', '時間枠計算テストでエラーが発生しました:\n' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 拡張検証機能のテスト関数
 */
function テスト_拡張検証() {
  try {
    writeLog('INFO', '=== 拡張検証テスト開始 ===');
    
    // テスト用の入社者データ
    var testNewHires = [
      {
        name: 'テスト太郎',
        rank: 'C',
        experience: '未経験者',
        email: 'test1@example.com'
      },
      {
        name: 'テスト花子',
        rank: 'M',
        experience: '経験者',
        email: 'test2@example.com'
      },
      {
        name: 'テスト一郎',
        rank: 'SC',
        experience: '経験者',
        email: 'test3@example.com'
      }
    ];
    
    // テスト用の研修グループデータ
    var testTrainingGroups = [
      {
        name: 'テスト研修A（未経験者向け）',
        attendees: ['test1@example.com', 'instructor1@example.com'],
        calendarEventId: null
      },
      {
        name: 'テスト研修B（M向け）',
        attendees: ['test2@example.com', 'instructor2@example.com'],
        calendarEventId: null
      },
      {
        name: 'テスト研修C（経験者向け）',
        attendees: ['test3@example.com', 'instructor3@example.com'],
        calendarEventId: null
      }
    ];
    
    var message = '拡張検証テスト結果:\n\n';
    
    // 1. 職位別研修数検証のテスト
    writeLog('INFO', '職位別研修数検証テスト実行中...');
    var validationResults = validateTrainingCountByPosition(testNewHires, testTrainingGroups);
    
    message += '【職位別研修数検証】\n';
    message += '全体判定: ' + (validationResults.overall ? '✅ 正常' : '❌ 異常') + '\n';
    message += '期待研修数: A=' + validationResults.expectedCounts['A'] + 
               ', B=' + validationResults.expectedCounts['B'] + 
               ', C=' + validationResults.expectedCounts['C'] + 
               ', D=' + validationResults.expectedCounts['D'] + '\n';
    
    if (validationResults.details && validationResults.details.length > 0) {
      message += '詳細結果:\n';
      for (var i = 0; i < validationResults.details.length; i++) {
        var detail = validationResults.details[i];
        var icon = detail.isValid ? '✅' : '❌';
        message += '  ' + detail.name + ': 期待' + detail.expected + '→実際' + detail.actual + ' ' + icon + '\n';
      }
    }
    message += '\n';
    
    // 2. カレンダー重複チェックのテスト
    writeLog('INFO', 'カレンダー重複チェックテスト実行中...');
    var conflictResults = validateCalendarTimeSlots(testTrainingGroups);
    
    message += '【カレンダー重複チェック】\n';
    message += '時間重複: ' + (conflictResults.hasConflicts ? '❌ あり' : '✅ なし') + '\n';
    message += 'チェック対象: ' + conflictResults.checkedEvents + '/' + conflictResults.totalEvents + 'イベント\n';
    
    if (conflictResults.conflicts && conflictResults.conflicts.length > 0) {
      message += '重複詳細:\n';
      for (var j = 0; j < Math.min(3, conflictResults.conflicts.length); j++) {
        var conflict = conflictResults.conflicts[j];
        message += '  ' + conflict.event1 + ' vs ' + conflict.event2 + '\n';
      }
    }
    message += '\n';
    
    message += '注意: 実際のカレンダーイベントがない場合、重複チェックは空の結果となります。\n';
    message += '詳細なログは実行ログシートをご確認ください。';
    
    SpreadsheetApp.getUi().alert('テスト結果', message, SpreadsheetApp.getUi().ButtonSet.OK);
    writeLog('INFO', '=== 拡張検証テスト完了 ===');
    
  } catch (e) {
    writeLog('ERROR', '拡張検証テストでエラー: ' + e.message);
    SpreadsheetApp.getUi().alert('エラー', '拡張検証テストでエラーが発生しました:\n' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * カレンダー重複問題の詳細検証テスト関数
 */
function テスト_カレンダー重複問題検証() {
  try {
    writeLog('INFO', '=== カレンダー重複問題検証テスト開始 ===');
    
    // 入社者データ取得（テスト用：入社日フィルタなし）
    var newHires = getNewHires();
    if (newHires.length === 0) {
      SpreadsheetApp.getUi().alert('警告', '入社者データがありません。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // 研修グループ作成
    var trainingGroups = groupTrainingsForHires(newHires);
    
    // テスト用の入社日（今日）
    var hireDate = new Date();
    hireDate.setHours(9, 0, 0, 0);
    
    writeLog('INFO', '重複検証テスト: 入社者' + newHires.length + '名, 研修' + trainingGroups.length + '件');
    
    // 一時的にscheduledEventsをリセット
    var originalScheduledEvents = scheduledEvents.slice();
    scheduledEvents = [];
    
    var roomManager = RoomReservationManager.getInstance();
    roomManager.reset();
    
    // 各研修の時間枠を計算
    var timeSlots = [];
    var conflicts = [];
    
    for (var i = 0; i < Math.min(16, trainingGroups.length); i++) {
      var group = trainingGroups[i];
      
      writeLog('INFO', '研修時間枠計算: ' + (i + 1) + '/' + Math.min(16, trainingGroups.length) + ' - ' + group.name);
      
      var timeSlot = findAvailableTimeSlot(group, hireDate);
      
      if (timeSlot) {
        // 既存の時間枠との重複をチェック
        for (var j = 0; j < timeSlots.length; j++) {
          var existing = timeSlots[j];
          
          // 時間重複チェック
          if (!(timeSlot.end <= existing.start || timeSlot.start >= existing.end)) {
            // 参加者重複チェック
            var hasCommonAttendee = false;
            if (group.attendees && existing.group.attendees) {
              for (var k = 0; k < group.attendees.length; k++) {
                if (existing.group.attendees.indexOf(group.attendees[k]) !== -1) {
                  hasCommonAttendee = true;
                  break;
                }
              }
            }
            
            conflicts.push({
              event1: group.name,
              event2: existing.group.name,
              time1: Utilities.formatDate(timeSlot.start, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + 
                     Utilities.formatDate(timeSlot.end, 'Asia/Tokyo', 'HH:mm'),
              time2: Utilities.formatDate(existing.start, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + 
                     Utilities.formatDate(existing.end, 'Asia/Tokyo', 'HH:mm'),
              hasCommonAttendee: hasCommonAttendee
            });
          }
        }
        
        timeSlots.push({
          group: group,
          start: timeSlot.start,
          end: timeSlot.end
        });
        
        // scheduledEventsに追加（後続の重複チェックのため）
        scheduledEvents.push({
          name: group.name,
          startTime: timeSlot.start,
          endTime: timeSlot.end,
          attendees: group.attendees || [],
          uniqueKey: generateEventUniqueKey({
            name: group.name,
            startTime: timeSlot.start,
            attendees: group.attendees || []
          })
        });
      } else {
        writeLog('WARN', '時間枠取得失敗: ' + group.name);
      }
    }
    
    // scheduledEventsを元に戻す
    scheduledEvents = originalScheduledEvents;
    
    var message = 'カレンダー重複問題検証結果（詰め配置版）:\n\n';
    message += '【スケジューリング結果】\n';
    message += '対象研修数: ' + Math.min(16, trainingGroups.length) + '件\n';
    message += '成功: ' + timeSlots.length + '件\n';
    message += '失敗: ' + (Math.min(16, trainingGroups.length) - timeSlots.length) + '件\n\n';
    
    // 詰め配置の効果を確認
    if (timeSlots.length > 1) {
      message += '【詰め配置効果確認】\n';
      for (var i = 0; i < Math.min(5, timeSlots.length); i++) {
        var slot = timeSlots[i];
        message += (i + 1) + '. ' + slot.group.name.substring(0, 30) + '...\n';
        message += '   時間: ' + Utilities.formatDate(slot.start, 'Asia/Tokyo', 'MM/dd HH:mm') + 
                   '-' + Utilities.formatDate(slot.end, 'Asia/Tokyo', 'HH:mm') + '\n';
        
        if (i > 0) {
          var prevSlot = timeSlots[i - 1];
          var gap = (slot.start.getTime() - prevSlot.end.getTime()) / (1000 * 60);
          message += '   前研修との間隔: ' + gap + '分\n';
        }
        message += '\n';
      }
    }
    
    message += '【重複チェック結果】\n';
    if (conflicts.length === 0) {
      message += '✅ 時間重複なし - 正常にスケジュールされました\n';
    } else {
      message += '❌ 時間重複あり: ' + conflicts.length + '件\n\n';
      message += '重複詳細:\n';
      for (var i = 0; i < Math.min(3, conflicts.length); i++) {
        var conflict = conflicts[i];
        var icon = conflict.hasCommonAttendee ? '⚠️' : 'ℹ️';
        message += icon + ' ' + conflict.event1 + ' vs ' + conflict.event2 + '\n';
        message += '   ' + conflict.time1 + ' / ' + conflict.time2 + '\n';
      }
      
      if (conflicts.length > 3) {
        message += '... 他' + (conflicts.length - 3) + '件\n';
      }
    }
    
    message += '\n詳細なログは実行ログシートをご確認ください。';
    
    SpreadsheetApp.getUi().alert('検証結果', message, SpreadsheetApp.getUi().ButtonSet.OK);
    writeLog('INFO', '=== カレンダー重複問題検証テスト完了 ===');
    
  } catch (e) {
    writeLog('ERROR', 'カレンダー重複問題検証テストでエラー: ' + e.message);
    SpreadsheetApp.getUi().alert('エラー', 'カレンダー重複問題検証テストでエラーが発生しました:\n' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

