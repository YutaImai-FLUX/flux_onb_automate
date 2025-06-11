// =========================================
// カレンダー関連ユーティリティ
// =========================================

/**
 * 参加人数に合う会議室名を検索する
 * @param {number} numberOfAttendees - 参加者数
 * @returns {string} 会議室名
 */
function findAvailableRoomName(numberOfAttendees) {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_IDS.ROOM_MASTER).getSheetByName(SHEET_NAMES.ROOM_MASTER);
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    
    for (var i = 0; i < data.length; i++) {
        var row = data[i];
        var roomName = row[0]; // A列: 会議室名
        var capacity = row[1]; // B列: 定員
        if (capacity >= numberOfAttendees) {
            return roomName;
        }
    }
    throw new Error('定員 ' + numberOfAttendees + ' 名以上の会議室が見つかりませんでした。');
}

/**
 * カレンダーイベントを作成する
 * @param {Object} trainingDetails - 研修情報
 * @param {string|null} roomName - 会議室名
 * @param {Date} periodStart - 研修期間開始日
 * @param {Date} periodEnd - 研修期間終了日
 */
function createCalendarEvent(trainingDetails, roomName, periodStart, periodEnd) {
    var title = trainingDetails.name;
    var sequence = trainingDetails.sequence || 1;
    var timeInfo = trainingDetails.time; // 例: "60分"
    
    writeLog('DEBUG', 'カレンダーイベント作成: ' + title + ', 順番: ' + sequence + ', 時間: ' + timeInfo);
    writeLog('DEBUG', '期間: ' + periodStart + ' から ' + periodEnd);
    
    // 期間の開始日から順番に基づいて実施日を計算
    var eventDate = new Date(periodStart);
    
    // 研修の順番に基づいて日付を調整（sequenceは1から始まる想定）
    // 平日のみを考慮して日付を進める
    var businessDaysToAdd = sequence - 1;
    var daysAdded = 0;
    while (daysAdded < businessDaysToAdd) {
        eventDate.setDate(eventDate.getDate() + 1);
        // 土曜日(6)と日曜日(0)をスキップ
        if (eventDate.getDay() !== 0 && eventDate.getDay() !== 6) {
            daysAdded++;
        }
    }
    
    // 期間終了日を超えないようにチェック
    if (eventDate > periodEnd) {
        writeLog('WARN', '研修 "' + title + '" の実施予定日が期間終了日を超えています。期間終了日に設定します。');
        eventDate = new Date(periodEnd);
    }
    
    // 時間情報を解析（例: "60分"）
    var durationMinutes = 60; // デフォルト60分
    if (timeInfo && timeInfo.indexOf('分') !== -1) {
        var timeMatch = timeInfo.match(/(\d+)分/);
        if (timeMatch) {
            durationMinutes = parseInt(timeMatch[1]);
        }
    }
    
    // 研修の開始時間を設定（9:00開始を基本とする）
    var startTime = new Date(eventDate.getFullYear(), eventDate.getMonth(), eventDate.getDate(), 9, 0);
    var endTime = new Date(startTime.getTime() + (durationMinutes * 60 * 1000));
    
    writeLog('INFO', '研修 "' + title + '" を ' + Utilities.formatDate(startTime, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + 
             ' から ' + Utilities.formatDate(endTime, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + ' で作成');
    
    var options = {
        description: trainingDetails.memo,
        guests: trainingDetails.attendees.join(','),
        location: roomName || '',
    };
    
    try {
        CalendarApp.createEvent(title, startTime, endTime, options);
        writeLog('INFO', 'カレンダーイベント作成成功: ' + title);
    } catch (e) {
        writeLog('ERROR', 'カレンダーイベント作成失敗: ' + title + ' - ' + e.message);
        throw e;
    }
} 