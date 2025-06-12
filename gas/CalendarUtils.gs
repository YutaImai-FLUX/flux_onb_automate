// =========================================
// カレンダー関連ユーティリティ
// =========================================

/**
 * 参加人数に合う会議室名を検索する
 * @param {number} numberOfAttendees - 参加者数
 * @returns {string} 会議室名
 */
function findAvailableRoomName(numberOfAttendees) {
    writeLog('DEBUG', '会議室検索関数開始: 必要人数=' + numberOfAttendees);
    
    try {
        var sheet = SpreadsheetApp.openById(SPREADSHEET_IDS.ROOM_MASTER).getSheetByName(SHEET_NAMES.ROOM_MASTER);
        var lastRow = sheet.getLastRow();
        writeLog('DEBUG', '会議室マスタ最終行: ' + lastRow);
        
        if (lastRow <= 1) {
            throw new Error('会議室マスタにデータがありません');
        }
        
        var data = sheet.getRange(2, 1, lastRow - 1, 4).getValues(); // D列まで取得
        writeLog('DEBUG', '会議室データ取得: ' + data.length + '件');
        
        // 条件を満たす会議室を収集
        var suitableRooms = [];
        for (var i = 0; i < data.length; i++) {
            var row = data[i];
            var roomName = row[0]; // A列: 会議室名
            var calendarId = row[1]; // B列: カレンダーID
            var capacity = row[2]; // C列: 定員
            var onbAvailable = row[3]; // D列: ONB研修対象
            
            writeLog('DEBUG', '会議室チェック: ' + roomName + ' (定員: ' + capacity + ', カレンダーID: ' + calendarId + ', ONB対象: ' + onbAvailable + ')');
            
            // ONB研修対象チェック：「利用可能」でない場合はスキップ
            if (onbAvailable !== '利用可能') {
                writeLog('DEBUG', 'ONB研修対象外のためスキップ: ' + roomName);
                continue;
            }
            
            if (capacity >= numberOfAttendees) {
                suitableRooms.push({
                    name: roomName,
                    calendarId: calendarId,
                    capacity: capacity
                });
                writeLog('DEBUG', '条件適合会議室: ' + roomName + ' (定員: ' + capacity + ', カレンダーID: ' + calendarId + ')');
            }
        }
        
        if (suitableRooms.length === 0) {
            throw new Error('定員 ' + numberOfAttendees + ' 名以上の会議室が見つかりませんでした。');
        }
        
        // 定員でソートして最小の適合会議室を選択
        suitableRooms.sort(function(a, b) {
            return a.capacity - b.capacity;
        });
        
        var selectedRoom = suitableRooms[0];
        writeLog('INFO', '最適会議室選択: ' + selectedRoom.name + ' (定員: ' + selectedRoom.capacity + 
                 ', 必要人数: ' + numberOfAttendees + ', 無駄席数: ' + (selectedRoom.capacity - numberOfAttendees) + ')');
        
        return selectedRoom.name;
        
    } catch (e) {
        writeLog('ERROR', '会議室検索でエラー: ' + e.message);
        throw e;
    }
}

/**
 * 会議室を時間枠と併せて確保する
 * @param {number} numberOfAttendees - 参加者数
 * @param {Date} startTime - 開始時間
 * @param {Date} endTime - 終了時間
 * @returns {string} 会議室名
 */
function findAndReserveRoom(numberOfAttendees, startTime, endTime) {
    writeLog('DEBUG', '会議室確保開始: 必要人数=' + numberOfAttendees + ', 時間=' + 
             Utilities.formatDate(startTime, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + 
             Utilities.formatDate(endTime, 'Asia/Tokyo', 'HH:mm'));
    
    // 参加者数が無効な場合はエラー
    if (!numberOfAttendees || numberOfAttendees <= 0) {
        throw new Error('参加者数が0以下です: ' + numberOfAttendees);
    }
    
    try {
        var sheet = SpreadsheetApp.openById(SPREADSHEET_IDS.ROOM_MASTER).getSheetByName(SHEET_NAMES.ROOM_MASTER);
        var lastRow = sheet.getLastRow();
        
        if (lastRow <= 1) {
            throw new Error('会議室マスタにデータがありません');
        }
        
        var data = sheet.getRange(2, 1, lastRow - 1, 4).getValues(); // D列まで取得
        
        // 条件を満たす会議室を収集
        var suitableRooms = [];
        for (var i = 0; i < data.length; i++) {
            var row = data[i];
            var roomName = row[0]; // A列: 会議室名
            var calendarId = row[1]; // B列: カレンダーID
            var capacity = row[2]; // C列: 定員
            var onbAvailable = row[3]; // D列: ONB研修対象
            
            // ONB研修対象チェック：「利用可能」でない場合はスキップ
            if (onbAvailable !== '利用可能') {
                writeLog('DEBUG', 'ONB研修対象外のためスキップ: ' + roomName + ' (ONB研修対象: ' + onbAvailable + ')');
                continue;
            }
            
            if (capacity >= numberOfAttendees) {
                // カレンダーIDを使って直接会議室の可用性をチェック
                if (calendarId && isGoogleCalendarRoomAvailable(calendarId, startTime, endTime)) {
                    suitableRooms.push({
                        name: roomName,
                        capacity: capacity,
                        calendarId: calendarId,
                        resourceEmail: calendarId,
                        resourceName: roomName
                    });
                    writeLog('DEBUG', '利用可能会議室: ' + roomName + ' (定員: ' + capacity + ', カレンダーID: ' + calendarId + ')');
                } else if (calendarId) {
                    writeLog('DEBUG', 'Googleカレンダー上で時間重複により利用不可: ' + roomName + ' (カレンダーID: ' + calendarId + ')');
                } else {
                    writeLog('WARN', 'カレンダーIDが設定されていません: ' + roomName);
                    // フォールバック: ローカル管理での可用性チェック
                    if (isRoomAvailable(roomName, startTime, endTime)) {
                        suitableRooms.push({
                            name: roomName,
                            capacity: capacity,
                            calendarId: null,
                            resourceEmail: null,
                            resourceName: roomName
                        });
                        writeLog('DEBUG', 'フォールバック利用可能会議室: ' + roomName + ' (定員: ' + capacity + ')');
                    }
                }
            }
        }
        
        if (suitableRooms.length === 0) {
            throw new Error('指定時間に利用可能で定員 ' + numberOfAttendees + ' 名以上の会議室が見つかりませんでした。');
        }
        
        // 定員でソートして最小の適合会議室を選択
        suitableRooms.sort(function(a, b) {
            return a.capacity - b.capacity;
        });
        
        var selectedRoom = suitableRooms[0];
        
        // 会議室を予約
        scheduledRooms.push({
            roomName: selectedRoom.name,
            calendarId: selectedRoom.calendarId,
            resourceEmail: selectedRoom.resourceEmail,
            startTime: startTime,
            endTime: endTime
        });
        
        writeLog('INFO', '会議室確保成功: ' + selectedRoom.name + ' (定員: ' + selectedRoom.capacity + 
                 ', 必要人数: ' + numberOfAttendees + ', 無駄席数: ' + (selectedRoom.capacity - numberOfAttendees) + 
                 ', カレンダーID: ' + (selectedRoom.calendarId || '未設定') + ')');
        
        return selectedRoom.resourceName || selectedRoom.name;
        
    } catch (e) {
        writeLog('ERROR', '会議室確保でエラー: ' + e.message);
        throw e;
    }
}

/**
 * 会議室マスタの名前を含むGoogleカレンダーリソースを検索する
 * @param {string} masterRoomName - 会議室マスタのA列の名前
 * @returns {Object|null} {email: string, name: string} または null
 */
function findGoogleCalendarRoom(masterRoomName) {
    writeLog('DEBUG', 'Googleカレンダーリソース検索開始: ' + masterRoomName);
    
    try {
        // Admin SDK Directory APIを使用してリソースを検索
        var resources = AdminDirectory.Resources.Calendars.list(Session.getActiveUser().getEmail().split('@')[1]);
        
        if (resources && resources.items) {
            writeLog('DEBUG', '利用可能なリソース数: ' + resources.items.length);
            
            for (var i = 0; i < resources.items.length; i++) {
                var resource = resources.items[i];
                var resourceName = resource.resourceName || '';
                var resourceEmail = resource.resourceEmail || '';
                
                // マスタの名前が含まれるリソースを検索
                if (resourceName.indexOf(masterRoomName) !== -1 || resourceEmail.indexOf(masterRoomName) !== -1) {
                    writeLog('INFO', 'マッチングリソース発見: ' + resourceName + ' (' + resourceEmail + ') - マスタ名: ' + masterRoomName);
                    return {
                        email: resourceEmail,
                        name: resourceName
                    };
                }
            }
            
            writeLog('WARN', 'マスタ名に一致するリソースが見つかりません: ' + masterRoomName);
        } else {
            writeLog('WARN', 'リソース一覧の取得に失敗しました');
        }
    } catch (e) {
        writeLog('WARN', 'Googleカレンダーリソース検索でエラー: ' + e.message + ' (マスタ名: ' + masterRoomName + ')');
        // Admin SDK APIが利用できない場合は、CalendarApp APIでフォールバック検索を試行
        return findGoogleCalendarRoomFallback(masterRoomName);
    }
    
    return null;
}

/**
 * CalendarApp APIを使用したフォールバック会議室検索
 * @param {string} masterRoomName - 会議室マスタのA列の名前
 * @returns {Object|null} {email: string, name: string} または null
 */
function findGoogleCalendarRoomFallback(masterRoomName) {
    writeLog('DEBUG', 'フォールバック会議室検索開始: ' + masterRoomName);
    
    try {
        // 一般的な会議室名のパターンでメールアドレスを構築して試行
        var domain = Session.getActiveUser().getEmail().split('@')[1];
        var possibleEmails = [
            masterRoomName + '@' + domain,
            masterRoomName.toLowerCase() + '@' + domain,
            masterRoomName.replace(/\s+/g, '') + '@' + domain,
            masterRoomName.replace(/\s+/g, '').toLowerCase() + '@' + domain,
            'room-' + masterRoomName.toLowerCase() + '@' + domain,
            'conference-' + masterRoomName.toLowerCase() + '@' + domain
        ];
        
        for (var i = 0; i < possibleEmails.length; i++) {
            var email = possibleEmails[i];
            try {
                var calendar = CalendarApp.getCalendarById(email);
                if (calendar) {
                    writeLog('INFO', 'フォールバック検索成功: ' + email + ' (マスタ名: ' + masterRoomName + ')');
                    return {
                        email: email,
                        name: calendar.getName() || masterRoomName
                    };
                }
            } catch (e) {
                // この email では見つからない場合は次を試行
                writeLog('DEBUG', 'フォールバック検索失敗: ' + email + ' - ' + e.message);
            }
        }
        
        writeLog('WARN', 'フォールバック検索でも会議室が見つかりません: ' + masterRoomName);
    } catch (e) {
        writeLog('WARN', 'フォールバック会議室検索でエラー: ' + e.message);
    }
    
    return null;
}

/**
 * Googleカレンダーリソースの可用性をチェック
 * @param {string} resourceEmail - リソースのメールアドレス
 * @param {Date} startTime - 開始時間
 * @param {Date} endTime - 終了時間
 * @returns {boolean} 利用可能かどうか
 */
function isGoogleCalendarRoomAvailable(resourceEmail, startTime, endTime) {
    writeLog('DEBUG', 'Googleカレンダーリソース可用性チェック: ' + resourceEmail + ' (' + 
             Utilities.formatDate(startTime, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + 
             Utilities.formatDate(endTime, 'Asia/Tokyo', 'HH:mm') + ')');
    
    try {
        var resourceCalendar = CalendarApp.getCalendarById(resourceEmail);
        if (!resourceCalendar) {
            writeLog('WARN', 'リソースカレンダーにアクセスできません: ' + resourceEmail);
            return false;
        }
        
        var existingEvents = resourceCalendar.getEvents(startTime, endTime);
        if (existingEvents.length > 0) {
            writeLog('DEBUG', 'リソースに既存予定あり: ' + existingEvents.length + '件');
            
            // 実際に時間重複している予定のみをチェック
            for (var i = 0; i < existingEvents.length; i++) {
                var event = existingEvents[i];
                var eventStart = event.getStartTime();
                var eventEnd = event.getEndTime();
                
                // 終日イベントをスキップ
                if (event.isAllDayEvent()) {
                    writeLog('DEBUG', '終日イベントスキップ: ' + event.getTitle());
                    continue;
                }
                
                // 実際の時間重複チェック
                if (!(endTime <= eventStart || startTime >= eventEnd)) {
                    writeLog('DEBUG', 'リソース時間重複: ' + event.getTitle() + ' (' + 
                             Utilities.formatDate(eventStart, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + 
                             Utilities.formatDate(eventEnd, 'Asia/Tokyo', 'HH:mm') + ')');
                    return false;
                }
            }
        }
        
        writeLog('DEBUG', 'Googleカレンダーリソース利用可能: ' + resourceEmail);
        return true;
        
    } catch (e) {
        writeLog('WARN', 'Googleカレンダーリソース可用性チェックでエラー: ' + e.message + ' (リソース: ' + resourceEmail + ')');
        return false;
    }
}

/**
 * 指定した時間枠で会議室が利用可能かチェック（従来のローカル管理）
 * @param {string} roomName - 会議室名
 * @param {Date} startTime - 開始時間
 * @param {Date} endTime - 終了時間
 * @returns {boolean} 利用可能かどうか
 */
function isRoomAvailable(roomName, startTime, endTime) {
    for (var i = 0; i < scheduledRooms.length; i++) {
        var reservation = scheduledRooms[i];
        if (reservation.roomName === roomName) {
            // 時間重複チェック
            if (!(endTime <= reservation.startTime || startTime >= reservation.endTime)) {
                return false;
            }
        }
    }
    return true;
}

// グローバル変数：スケジュール管理用
var scheduledEvents = [];
var scheduledRooms = []; // 会議室の時間枠管理

/**
 * 全ての研修のカレンダーイベントを作成する（時間重複を回避）
 * @param {Array<Object>} trainingGroups - 研修グループの配列
 * @param {Date} hireDate - 入社日
 */
function createAllCalendarEvents(trainingGroups, hireDate) {
    writeLog('INFO', '全研修のカレンダーイベント作成開始');
    
    // スケジュール管理用配列をリセット
    scheduledEvents = [];
    scheduledRooms = [];
    
    // 研修グループは既にLogic.gsでソート済みなので、順序をそのまま使用
    var sortedGroups = trainingGroups.slice();
    
    var scheduleResults = [];
    
    for (var i = 0; i < sortedGroups.length; i++) {
        var group = sortedGroups[i];
        
        writeLog('INFO', '研修処理開始: ' + group.name + ' (ユニークキー: ' + (group.uniqueKey || 'なし') + ')');
        
        // ユニークキーベースで重複チェック
        var isDuplicate = false;
        if (group.uniqueKey) {
            for (var j = 0; j < scheduledEvents.length; j++) {
                var existingEvent = scheduledEvents[j];
                if (existingEvent.uniqueKey === group.uniqueKey) {
                    writeLog('WARN', '重複イベントを検出、スキップ: ' + group.name + ' (ユニークキー: ' + group.uniqueKey + ')');
                    isDuplicate = true;
                    break;
                }
            }
        } else {
            // ユニークキーがない場合は従来の方法でチェック
            for (var j = 0; j < scheduledEvents.length; j++) {
                var existingEvent = scheduledEvents[j];
                if (existingEvent.name === group.name && 
                    existingEvent.lecturer === group.lecturer &&
                    JSON.stringify(existingEvent.attendees.sort()) === JSON.stringify(group.attendees.sort())) {
                    writeLog('WARN', '重複イベントを検出、スキップ: ' + group.name + ' (講師: ' + group.lecturer + ')');
                    isDuplicate = true;
                    break;
                }
            }
        }
        
        if (isDuplicate) {
            continue;
        }
        
        // 参加者数チェック（講師以外の参加者がいるかチェック）
        var participantCount = 0;
        if (group.attendees && group.attendees.length > 0) {
            // 講師を除いた参加者数を計算
            for (var k = 0; k < group.attendees.length; k++) {
                if (group.attendees[k] !== group.lecturer) {
                    participantCount++;
                }
            }
        }
        
        writeLog('DEBUG', '参加者数チェック: ' + group.name + ' - 講師除く参加者数: ' + participantCount);
        
        if (participantCount === 0) {
            writeLog('INFO', '参加者が0人のため研修をスキップ: ' + group.name);
            scheduleResults.push({
                training: group,
                scheduled: false,
                roomName: '実施不要',
                eventTime: null,
                error: '参加者0人のため実施不要'
            });
            continue;
        }
        
        // 時間枠を先に確保
        var eventTime = findAvailableTimeSlot(group, hireDate);
        
        if (!eventTime) {
            writeLog('ERROR', '研修 "' + group.name + '" の適切な時間枠が見つかりませんでした');
            scheduleResults.push({
                training: group,
                scheduled: false,
                roomName: group.needsRoom ? '会議室未確保' : 'オンライン',
                eventTime: null,
                error: '時間枠未確保'
            });
            continue;
        }
        
        // 会議室が必要な場合は確保（時間枠も考慮）
        var roomName = null;
        var roomError = null;
        if (group.needsRoom) {
            try {
                // 会議室定員は講師＋参加者の総数で計算（講師も席が必要）
                var lecturerCount = group.lecturerEmails ? group.lecturerEmails.length : 1; // 講師数を取得（デフォルト1）
                var totalAttendeeCount = participantCount + lecturerCount; // 講師分を追加
                writeLog('DEBUG', '会議室確保試行: 研修=' + group.name + ', 総参加者数=' + totalAttendeeCount + ' (講師' + lecturerCount + '名+参加者' + participantCount + '名), 時間=' + 
                         Utilities.formatDate(eventTime.start, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + 
                         Utilities.formatDate(eventTime.end, 'Asia/Tokyo', 'HH:mm'));
                
                roomName = findAndReserveRoom(totalAttendeeCount, eventTime.start, eventTime.end);
                writeLog('INFO', '会議室確保成功: ' + roomName + ' (研修: ' + group.name + ', 総参加者: ' + totalAttendeeCount + '名, 講師' + lecturerCount + '名+参加者' + participantCount + '名)');
            } catch (e) {
                writeLog('ERROR', '会議室確保失敗: ' + e.message + ' (研修: ' + group.name + ')');
                roomName = '会議室未確保';
                roomError = '会議室確保失敗: ' + e.message;
                // 会議室確保に失敗してもカレンダーイベントは作成する
            }
        } else {
            roomName = 'オンライン';
        }
        
        // カレンダーイベント作成
        try {
            createSingleCalendarEvent(group, roomName, eventTime.start, eventTime.end);
            
            // カレンダーイベントIDが正しく設定されているかチェック
            if (!group.calendarEventId) {
                writeLog('ERROR', 'カレンダーイベント作成後にIDが設定されていません: ' + group.name);
                throw new Error('カレンダーイベントIDが設定されませんでした');
            }
            
            // スケジュール管理配列に追加
            scheduledEvents.push({
                name: group.name,
                lecturer: group.lecturer,
                startTime: eventTime.start,
                endTime: eventTime.end,
                attendees: group.attendees.slice(),
                uniqueKey: group.uniqueKey,
                roomName: roomName,
                calendarEventId: group.calendarEventId
            });
            
            scheduleResults.push({
                training: group,
                scheduled: true,
                roomName: roomName,
                eventTime: eventTime,
                error: roomError,  // 会議室確保失敗のエラーも記録
                calendarEventId: group.calendarEventId  // カレンダーイベントIDを追加
            });
            
            writeLog('INFO', 'カレンダーイベント作成完了: ' + group.name + ' (ID: ' + group.calendarEventId + ')');
        } catch (e) {
            writeLog('ERROR', 'カレンダーイベント作成失敗: ' + group.name + ' - ' + e.message);
            scheduleResults.push({
                training: group,
                scheduled: false,
                roomName: roomName,
                eventTime: eventTime,
                error: 'カレンダー作成失敗',
                calendarEventId: null
            });
        }
    }
    
    writeLog('INFO', '全研修のカレンダーイベント作成完了: ' + scheduledEvents.length + '件');
    return scheduleResults;
}

/**
 * 指定した時間枠で利用可能な会議室があるかチェックする（予約はしない）
 * @param {number} numberOfAttendees - 参加者数
 * @param {Date} startTime - 開始時間
 * @param {Date} endTime - 終了時間
 * @returns {boolean} 利用可能な会議室があるかどうか
 */
function isAnyRoomAvailable(numberOfAttendees, startTime, endTime) {
    writeLog('DEBUG', '利用可能な会議室の存在チェック: 必要人数=' + numberOfAttendees + ', 時間=' + 
             Utilities.formatDate(startTime, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + 
             Utilities.formatDate(endTime, 'Asia/Tokyo', 'HH:mm'));
    try {
        var sheet = SpreadsheetApp.openById(SPREADSHEET_IDS.ROOM_MASTER).getSheetByName(SHEET_NAMES.ROOM_MASTER);
        var lastRow = sheet.getLastRow();
        if (lastRow <= 1) return false;
        
        var data = sheet.getRange(2, 1, lastRow - 1, 4).getValues(); // D列まで取得
        
        for (var i = 0; i < data.length; i++) {
            var row = data[i];
            var roomName = row[0];
            var calendarId = row[1];
            var capacity = row[2];
            var onbAvailable = row[3]; // D列: ONB研修対象
            
            // ONB研修対象チェック：「利用可能」でない場合はスキップ
            if (onbAvailable !== '利用可能') {
                writeLog('DEBUG', 'ONB研修対象外のためスキップ: ' + roomName);
                continue;
            }
            
            if (capacity >= numberOfAttendees) {
                // この実行で確保済みの会議室リストと重複していないかチェック
                if (!isRoomAvailable(roomName, startTime, endTime)) {
                    writeLog('DEBUG', '内部スケジュールで予約済みのためスキップ: ' + roomName);
                    continue; 
                }

                // Googleカレンダー上で空いているかチェック
                if (calendarId && isGoogleCalendarRoomAvailable(calendarId, startTime, endTime)) {
                    writeLog('DEBUG', '利用可能な会議室を発見: ' + roomName);
                    return true;
                }
            }
        }
        
        writeLog('DEBUG', '利用可能な会議室が見つかりませんでした');
        return false;
    } catch (e) {
        writeLog('WARN', '会議室の空き状況チェックでエラー: ' + e.message);
        return false; // エラー時は利用不可とみなす
    }
}

/**
 * 講師の空き時間を考慮して適切な時間枠を見つける（実施日を考慮）
 * @param {Object} trainingGroup - 研修グループ
 * @param {Date} hireDate - 入社日
 * @returns {Object|null} {start: Date, end: Date} または null
 */
function findAvailableTimeSlot(trainingGroup, hireDate) {
    var durationMinutes = 60; // デフォルト60分
    if (trainingGroup.time && trainingGroup.time.indexOf('分') !== -1) {
        var timeMatch = trainingGroup.time.match(/(\d+)分/);
        if (timeMatch) {
            durationMinutes = parseInt(timeMatch[1]);
        }
    }
    
    writeLog('DEBUG', '時間枠検索開始: ' + trainingGroup.name + ' (時間: ' + durationMinutes + '分, 実施日: ' + (trainingGroup.implementationDay || '未指定') + '営業日目)');
    
    // 実施日が指定されている場合は、入社日を基準にその日を計算
    var targetDate = null;
    if (trainingGroup.implementationDay && trainingGroup.implementationDay !== 999) {
        targetDate = calculateImplementationDate(hireDate, trainingGroup.implementationDay);
        writeLog('DEBUG', '指定実施日: ' + trainingGroup.implementationDay + '営業日目 = ' + Utilities.formatDate(targetDate, 'Asia/Tokyo', 'yyyy/MM/dd(E)'));
    } else {
        // 実施日が指定されていない場合は入社日当日から検索
        targetDate = new Date(hireDate);
        writeLog('DEBUG', '実施日未指定のため、入社日から検索: ' + Utilities.formatDate(targetDate, 'Asia/Tokyo', 'yyyy/MM/dd(E)'));
    }
    
    // 指定日またはその近辺で時間枠を検索（最大30日先まで）
    var searchStartDate = new Date(targetDate);
    var searchEndDate = new Date(targetDate);
    searchEndDate.setDate(searchEndDate.getDate() + 30); // 30日先まで検索
    
    var currentDate = new Date(searchStartDate);
    
    while (currentDate <= searchEndDate) {
        // 平日のみ処理
        if (currentDate.getDay() !== 0 && currentDate.getDay() !== 6) {
            
            // 実施日に基づいて開始時間を決定
            var startHour = 9; // デフォルトは9時
            if (trainingGroup.implementationDay === 1) {
                startHour = 15; // 1営業日目は15時以降
                writeLog('DEBUG', '実施日1日目のため、検索開始時間を15時に設定');
            } else if (trainingGroup.implementationDay === 2) {
                startHour = 16; // 2営業日目は16時以降
                writeLog('DEBUG', '実施日2日目のため、検索開始時間を16時に設定');
            }
            
            // 1日の時間枠をチェック（30分刻み）
            // 1日目のみ19時まで、それ以外は18時まで
            var maxHour = (trainingGroup.implementationDay === 1) ? 19 : 18;
            var maxEndTime = (trainingGroup.implementationDay === 1) ? 20 : 19;
            
            for (var hour = startHour; hour <= maxHour; hour++) {
                // 00分と30分の両方をチェック
                var minutes = [0, 30];
                for (var m = 0; m < minutes.length; m++) {
                    var minute = minutes[m];
                    var proposedStart = new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate(), hour, minute);
                    var proposedEnd = new Date(proposedStart.getTime() + (durationMinutes * 60 * 1000));
                    
                    // 1日目は20:00、それ以外は19:00を超える場合はスキップ
                    if (proposedEnd.getHours() > maxEndTime || (proposedEnd.getHours() === maxEndTime && proposedEnd.getMinutes() > 0)) {
                        continue; // breakではなくcontinueで次の時間枠をチェック
                    }
                    
                    // 昼休み時間帯（12:00-13:00）との重複チェック
                    var lunchStart = new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate(), 12, 0);
                    var lunchEnd = new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate(), 13, 0);
                    
                    // 提案時間と昼休み時間の重複チェック（開始時間 < 昼休み終了時間 AND 終了時間 > 昼休み開始時間）
                    if (proposedStart < lunchEnd && proposedEnd > lunchStart) {
                        writeLog('DEBUG', '昼休み時間帯（12:00-13:00）と重複するためスキップ: ' + 
                                 Utilities.formatDate(proposedStart, 'Asia/Tokyo', 'HH:mm') + '-' + 
                                 Utilities.formatDate(proposedEnd, 'Asia/Tokyo', 'HH:mm'));
                        continue; // 昼休み時間帯と重複するためスキップ
                    }
                    
                    // 時間重複チェック（参加者）
                    if (isTimeSlotAvailable(proposedStart, proposedEnd, trainingGroup)) {
                        // 参加者が利用可能な場合、会議室の要否をチェック
                        if (trainingGroup.needsRoom) {
                            // 会議室が必要な場合、空きがあるかチェック（参加者全員の人数で）
                            if (!isAnyRoomAvailable(trainingGroup.attendees.length, proposedStart, proposedEnd)) {
                                writeLog('DEBUG', '時間枠 ' + Utilities.formatDate(proposedStart, 'Asia/Tokyo', 'HH:mm') + '-' + Utilities.formatDate(proposedEnd, 'Asia/Tokyo', 'HH:mm') + ' は参加者全員が空いているが、利用可能な会議室がないためスキップ');
                                continue; // 利用可能な会議室がないため、次の時間枠へ
                            }
                        }

                        // この時間枠は利用可能
                        writeLog('INFO', '適切な時間枠発見: ' + Utilities.formatDate(proposedStart, 'Asia/Tokyo', 'MM/dd HH:mm') + 
                                 '-' + Utilities.formatDate(proposedEnd, 'Asia/Tokyo', 'HH:mm') + 
                                 ' (実施日' + (trainingGroup.implementationDay || '未指定') + '営業日目)');
                        return {
                            start: proposedStart,
                            end: proposedEnd
                        };
                    }
                }
            }
        }
        
        // 次の日へ
        currentDate.setDate(currentDate.getDate() + 1);
    }
    
    writeLog('WARN', '適切な時間枠が見つかりませんでした: ' + trainingGroup.name + ' (実施日: ' + (trainingGroup.implementationDay || '未指定') + '営業日目)');
    return null;
}

/**
 * 入社日から指定営業日数後の日付を計算する
 * @param {Date} hireDate - 入社日
 * @param {number} businessDays - 営業日数
 * @returns {Date} 計算された日付
 */
function calculateImplementationDate(hireDate, businessDays) {
    var currentDate = new Date(hireDate);
    var daysAdded = 0;
    
    // 1営業日目は入社日当日とする
    businessDays = businessDays - 1;
    
    while (daysAdded < businessDays) {
        currentDate.setDate(currentDate.getDate() + 1);
        // 平日のみカウント
        if (currentDate.getDay() !== 0 && currentDate.getDay() !== 6) {
            daysAdded++;
        }
    }
    
    return currentDate;
}

/**
 * 指定した時間枠が利用可能かチェック
 * @param {Date} proposedStart - 提案開始時間
 * @param {Date} proposedEnd - 提案終了時間
 * @param {Object} trainingGroup - チェック対象の研修グループ
 * @returns {boolean} 利用可能かどうか
 */
function isTimeSlotAvailable(proposedStart, proposedEnd, trainingGroup) {
    var proposedAttendees = trainingGroup.attendees || [];
    var lecturerEmail = trainingGroup.lecturer;

    // 1. 既にこの実行でスケジュールされたイベントとの重複チェック
    for (var i = 0; i < scheduledEvents.length; i++) {
        var scheduledEvent = scheduledEvents[i];
        var scheduledAttendees = scheduledEvent.attendees || [];

        // 参加者に重複があるかチェック（講師も参加者に含まれる）
        var hasOverlappingAttendee = proposedAttendees.some(function(attendee) {
            return scheduledAttendees.indexOf(attendee) !== -1;
        });
        
        if (hasOverlappingAttendee) {
            // 時間重複チェック
            if (!(proposedEnd <= scheduledEvent.startTime || proposedStart >= scheduledEvent.endTime)) {
                writeLog('DEBUG', '時間重複検出（内部スケジュール）: ' + 
                         trainingGroup.name + ' と ' + scheduledEvent.name + ' で参加者重複');
                return false;
            }
        }
    }
    
    // 2. 講師のGoogleカレンダーとの重複チェック (元のロジックを維持)
    try {
        if (lecturerEmail && lecturerEmail.trim() !== '') {
            writeLog('DEBUG', '講師カレンダーチェック開始: ' + lecturerEmail);
            
            // 講師のカレンダーイベントを取得
            var lecturerCalendar = CalendarApp.getCalendarById(lecturerEmail);
            if (lecturerCalendar) {
                var existingEvents = lecturerCalendar.getEvents(proposedStart, proposedEnd);
                if (existingEvents.length > 0) {
                    writeLog('DEBUG', '講師の既存予定候補: ' + lecturerEmail + ' (' + existingEvents.length + '件の予定)');
                    
                    // 実際に時間重複している予定のみをチェック
                    var conflictingEvents = [];
                    for (var i = 0; i < existingEvents.length; i++) {
                        var event = existingEvents[i];
                        var eventStart = event.getStartTime();
                        var eventEnd = event.getEndTime();
                        
                        // 終日イベント（00:00-00:00や日付のみ）をスキップ
                        var eventDuration = eventEnd.getTime() - eventStart.getTime();
                        var isAllDayEvent = event.isAllDayEvent() || 
                                          (eventStart.getHours() === 0 && eventStart.getMinutes() === 0 &&
                                           eventEnd.getHours() === 0 && eventEnd.getMinutes() === 0) ||
                                          eventDuration >= 24 * 60 * 60 * 1000; // 24時間以上
                        
                        if (isAllDayEvent) {
                            writeLog('DEBUG', '終日イベントのためスキップ: ' + event.getTitle() + ' (' + 
                                     Utilities.formatDate(eventStart, 'Asia/Tokyo', 'MM/dd HH:mm') + 
                                     '-' + Utilities.formatDate(eventEnd, 'Asia/Tokyo', 'HH:mm') + ')');
                            continue;
                        }
                        
                        // 実際の時間重複チェック
                        if (!(proposedEnd <= eventStart || proposedStart >= eventEnd)) {
                            conflictingEvents.push(event);
                            writeLog('DEBUG', '時間重複予定: ' + event.getTitle() + ' (' + 
                                     Utilities.formatDate(eventStart, 'Asia/Tokyo', 'MM/dd HH:mm') + 
                                     '-' + Utilities.formatDate(eventEnd, 'Asia/Tokyo', 'HH:mm') + ')');
                        } else {
                            writeLog('DEBUG', '時間重複なし: ' + event.getTitle() + ' (' + 
                                     Utilities.formatDate(eventStart, 'Asia/Tokyo', 'MM/dd HH:mm') + 
                                     '-' + Utilities.formatDate(eventEnd, 'Asia/Tokyo', 'HH:mm') + ')');
                        }
                    }
                    
                    if (conflictingEvents.length > 0) {
                        writeLog('DEBUG', '講師の実際の重複予定: ' + conflictingEvents.length + '件');
                        return false;
                    } else {
                        writeLog('DEBUG', '講師カレンダー重複なし（終日イベント等は除外）');
                    }
                }
            } else {
                writeLog('WARN', '講師のカレンダーにアクセスできません: ' + lecturerEmail);
            }
        }
    } catch (e) {
        writeLog('WARN', '講師カレンダーチェックでエラー: ' + e.message + ' (講師: ' + lecturerEmail + ')');
        // エラーの場合は重複していないとみなす（厳格すぎるより寛容に）
    }
    
    return true;
}

/**
 * 単一のカレンダーイベントを作成する
 * @param {Object} trainingDetails - 研修情報
 * @param {string|null} roomName - 会議室名
 * @param {Date} startTime - 開始時間
 * @param {Date} endTime - 終了時間
 */
function createSingleCalendarEvent(trainingDetails, roomName, startTime, endTime) {
    var title = trainingDetails.name;
    
    writeLog('INFO', '研修 "' + title + '" を ' + Utilities.formatDate(startTime, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + 
             ' から ' + Utilities.formatDate(endTime, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + ' で作成');
    
    // 有効なメールアドレスのみをフィルタリング
    var validEmails = [];
    for (var i = 0; i < trainingDetails.attendees.length; i++) {
        var email = trainingDetails.attendees[i];
        if (email && email.trim() !== '' && email.indexOf('@') > 0 && email.indexOf('@') < email.length - 1) {
            validEmails.push(email.trim());
        } else {
            writeLog('WARN', '無効なメールアドレスをスキップ: ' + email);
        }
    }
    
    // 会議室リソースがある場合は参加者に追加
    var roomResourceEmail = null;
    if (roomName && roomName !== 'オンライン' && roomName !== '会議室未確保' && roomName !== '実施不要') {
        // スケジュールされた会議室情報から対応するカレンダーIDを取得
        for (var i = 0; i < scheduledRooms.length; i++) {
            var scheduledRoom = scheduledRooms[i];
            if (scheduledRoom.roomName && scheduledRoom.roomName === roomName) {
                if (scheduledRoom.calendarId) {
                    roomResourceEmail = scheduledRoom.calendarId;
                    validEmails.push(roomResourceEmail);
                    writeLog('INFO', '会議室カレンダーIDを参加者に追加: ' + roomResourceEmail);
                    break;
                } else if (scheduledRoom.resourceEmail) {
                    roomResourceEmail = scheduledRoom.resourceEmail;
                    validEmails.push(roomResourceEmail);
                    writeLog('INFO', '会議室リソースを参加者に追加: ' + roomResourceEmail);
                    break;
                }
            }
        }
        
        // スケジュールされた会議室情報にない場合は、会議室マスタから検索
        if (!roomResourceEmail) {
            try {
                var sheet = SpreadsheetApp.openById(SPREADSHEET_IDS.ROOM_MASTER).getSheetByName(SHEET_NAMES.ROOM_MASTER);
                var lastRow = sheet.getLastRow();
                if (lastRow > 1) {
                    var data = sheet.getRange(2, 1, lastRow - 1, 4).getValues(); // D列まで取得
                    for (var j = 0; j < data.length; j++) {
                        var row = data[j];
                        var onbAvailable = row[3]; // D列: ONB研修対象
                        // ONB研修対象が「利用可能」の会議室のみ使用
                        if (row[0] === roomName && row[1] && onbAvailable === '利用可能') { // A列: 会議室名, B列: カレンダーID, D列: ONB研修対象
                            roomResourceEmail = row[1];
                            validEmails.push(roomResourceEmail);
                            writeLog('INFO', '会議室カレンダーIDを参加者に追加（マスタ検索）: ' + roomResourceEmail);
                            break;
                        }
                    }
                }
            } catch (e) {
                writeLog('WARN', '会議室マスタ検索でエラー: ' + e.message);
            }
        }
    }
    
    writeLog('DEBUG', '有効なメールアドレス (' + validEmails.length + '件): ' + validEmails.join(', '));
    
    // location文字列を決定
    var locationString = roomName || '';
    // 会議室リソースがカレンダーに追加されている場合、locationはリソースで表現されるため、テキストのlocationは空にする
    if (roomResourceEmail) {
        locationString = '';
    }
    
    var options = {
        description: trainingDetails.memo,
        guests: validEmails.join(','),
        location: locationString,
    };
    
    try {
        var calendarEvent = CalendarApp.createEvent(title, startTime, endTime, options);
        var eventId = calendarEvent.getId();
        
        // trainingDetailsにイベントIDを追加して返す
        trainingDetails.calendarEventId = eventId;
        
        writeLog('INFO', 'カレンダーイベント作成成功: ' + title + ' (ID: ' + eventId + ', 参加者: ' + validEmails.length + '名' + 
                 (roomResourceEmail ? ', 会議室カレンダーID: ' + roomResourceEmail : '') + ')');
    } catch (e) {
        writeLog('ERROR', 'カレンダーイベント作成失敗: ' + title + ' - ' + e.message);
        writeLog('ERROR', '参加者リスト: ' + validEmails.join(', '));
        throw e;
    }
}

/**
 * 従来のcreateCalendarEvent関数（互換性のため）
 * @param {Object} trainingDetails - 研修情報
 * @param {string|null} roomName - 会議室名
 * @param {Date} hireDate - 入社日
 */
function createCalendarEvent(trainingDetails, roomName, hireDate) {
    // 新しいスケジューリングシステムを使用
    createAllCalendarEvents([trainingDetails], hireDate);
}

// =========================================
// カレンダー削除関連
// =========================================

/**
 * 単一のカレンダーイベントを削除する
 * @param {string} eventId - 削除するイベントのID
 * @returns {boolean} 削除成功かどうか
 */
function deleteSingleCalendarEvent(eventId) {
    writeLog('INFO', 'カレンダーイベント削除開始: ' + eventId);
    
    try {
        var event = CalendarApp.getEventById(eventId);
        if (!event) {
            writeLog('WARN', 'イベントが見つかりません: ' + eventId);
            return false;
        }
        
        var eventTitle = event.getTitle();
        event.deleteEvent();
        
        writeLog('INFO', 'カレンダーイベント削除成功: ' + eventTitle + ' (ID: ' + eventId + ')');
        return true;
    } catch (e) {
        writeLog('ERROR', 'カレンダーイベント削除失敗: ' + eventId + ' - ' + e.message);
        return false;
    }
}

/**
 * マッピングシートからカレンダーIDを取得してイベントを削除する
 * @param {string} sheetName - マッピングシート名（省略時は最新のシートを使用）
 * @returns {Object} 削除結果のサマリー
 */
function deleteCalendarEventsFromMappingSheet(sheetName) {
    writeLog('INFO', 'マッピングシートからのカレンダーイベント削除開始');
    
    var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_IDS.EXECUTION);
    var mappingSheet = null;
    
    if (sheetName) {
        try {
            mappingSheet = spreadsheet.getSheetByName(sheetName);
        } catch (e) {
            writeLog('ERROR', '指定されたシートが見つかりません: ' + sheetName);
            throw new Error('指定されたシートが見つかりません: ' + sheetName);
        }
    } else {
        // 最新のマッピングシートを探す
        var existingSheets = spreadsheet.getSheets();
        for (var i = 0; i < existingSheets.length; i++) {
            if (existingSheets[i].getName().indexOf('マッピング結果_') === 0) {
                mappingSheet = existingSheets[i];
                break;
            }
        }
    }
    
    if (!mappingSheet) {
        writeLog('ERROR', 'マッピングシートが見つかりません');
        throw new Error('マッピングシートが見つかりません');
    }
    
    writeLog('INFO', '対象シート: ' + mappingSheet.getName());
    
    // ヘッダー行を取得してカレンダーID列の位置を特定
    var lastCol = mappingSheet.getLastColumn();
    var lastRow = mappingSheet.getLastRow();
    
    if (lastRow <= 1) {
        writeLog('WARN', 'データが存在しません');
        return { success: 0, failed: 0, total: 0, errors: [] };
    }
    
    var headerRow = mappingSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var calendarIdCol = -1;
    var trainingNameCol = 1; // A列: 研修名
    
    for (var col = 0; col < headerRow.length; col++) {
        if (headerRow[col] === 'カレンダーID') {
            calendarIdCol = col + 1; // 1ベースのインデックス
            break;
        }
    }
    
    if (calendarIdCol === -1) {
        writeLog('ERROR', 'カレンダーID列が見つかりません');
        throw new Error('カレンダーID列が見つかりません。シート構造を確認してください。');
    }
    
    writeLog('DEBUG', 'カレンダーID列: ' + calendarIdCol + '列目');
    
    // データ行を処理
    var successCount = 0;
    var failedCount = 0;
    var totalCount = 0;
    var errors = [];
    
    for (var row = 2; row <= lastRow; row++) {
        var trainingName = mappingSheet.getRange(row, trainingNameCol).getValue();
        var eventId = mappingSheet.getRange(row, calendarIdCol).getValue();
        
        if (!eventId || eventId.toString().trim() === '') {
            writeLog('DEBUG', '行' + row + ': カレンダーIDが空のためスキップ - ' + trainingName);
            continue;
        }
        
        totalCount++;
        writeLog('DEBUG', '行' + row + ': イベント削除試行 - ' + trainingName + ' (ID: ' + eventId + ')');
        
        if (deleteSingleCalendarEvent(eventId)) {
            successCount++;
            // 削除成功時はカレンダーID列をクリア
            mappingSheet.getRange(row, calendarIdCol).setValue('削除済み');
        } else {
            failedCount++;
            var errorMsg = '削除失敗: ' + trainingName + ' (ID: ' + eventId + ')';
            errors.push(errorMsg);
            // 削除失敗時は「削除失敗」と記録
            mappingSheet.getRange(row, calendarIdCol).setValue('削除失敗');
        }
    }
    
    var result = {
        success: successCount,
        failed: failedCount,
        total: totalCount,
        errors: errors,
        sheetName: mappingSheet.getName()
    };
    
    writeLog('INFO', 'カレンダーイベント削除完了: 成功=' + successCount + ', 失敗=' + failedCount + ', 総数=' + totalCount);
    
    return result;
}

/**
 * 特定の研修名のカレンダーイベントを削除する
 * @param {string} trainingName - 削除対象の研修名
 * @param {string} sheetName - マッピングシート名（省略時は最新のシートを使用）
 * @returns {boolean} 削除成功かどうか
 */
function deleteSpecificTrainingEvent(trainingName, sheetName) {
    writeLog('INFO', '特定研修のカレンダーイベント削除: ' + trainingName);
    
    var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_IDS.EXECUTION);
    var mappingSheet = null;
    
    if (sheetName) {
        try {
            mappingSheet = spreadsheet.getSheetByName(sheetName);
        } catch (e) {
            writeLog('ERROR', '指定されたシートが見つかりません: ' + sheetName);
            return false;
        }
    } else {
        // 最新のマッピングシートを探す
        var existingSheets = spreadsheet.getSheets();
        for (var i = 0; i < existingSheets.length; i++) {
            if (existingSheets[i].getName().indexOf('マッピング結果_') === 0) {
                mappingSheet = existingSheets[i];
                break;
            }
        }
    }
    
    if (!mappingSheet) {
        writeLog('ERROR', 'マッピングシートが見つかりません');
        return false;
    }
    
    // ヘッダー行を取得してカレンダーID列の位置を特定
    var lastCol = mappingSheet.getLastColumn();
    var lastRow = mappingSheet.getLastRow();
    
    if (lastRow <= 1) {
        writeLog('WARN', 'データが存在しません');
        return false;
    }
    
    var headerRow = mappingSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var calendarIdCol = -1;
    var trainingNameCol = 1; // A列: 研修名
    
    for (var col = 0; col < headerRow.length; col++) {
        if (headerRow[col] === 'カレンダーID') {
            calendarIdCol = col + 1; // 1ベースのインデックス
            break;
        }
    }
    
    if (calendarIdCol === -1) {
        writeLog('ERROR', 'カレンダーID列が見つかりません');
        return false;
    }
    
    // 指定された研修名のイベントを探して削除
    for (var row = 2; row <= lastRow; row++) {
        var currentTrainingName = mappingSheet.getRange(row, trainingNameCol).getValue();
        
        if (currentTrainingName === trainingName) {
            var eventId = mappingSheet.getRange(row, calendarIdCol).getValue();
            
            if (!eventId || eventId.toString().trim() === '') {
                writeLog('WARN', '研修 "' + trainingName + '" のカレンダーIDが空です');
                return false;
            }
            
            if (deleteSingleCalendarEvent(eventId)) {
                mappingSheet.getRange(row, calendarIdCol).setValue('削除済み');
                writeLog('INFO', '研修 "' + trainingName + '" のカレンダーイベント削除成功');
                return true;
            } else {
                mappingSheet.getRange(row, calendarIdCol).setValue('削除失敗');
                writeLog('ERROR', '研修 "' + trainingName + '" のカレンダーイベント削除失敗');
                return false;
            }
        }
    }
    
    writeLog('WARN', '研修 "' + trainingName + '" が見つかりませんでした');
    return false;
} 