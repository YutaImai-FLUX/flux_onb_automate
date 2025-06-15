// =========================================
// カレンダー関連ユーティリティ
// =========================================

/**
 * 参加人数に合う会議室名を検索する（ハイブリッド対応）
 * @param {number} numberOfAttendees - 参加者数
 * @returns {Object} {roomName: string, isHybrid: boolean, maxCapacity: number}
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
        var allAvailableRooms = []; // ハイブリッド用（定員不足でも利用可能な会議室）
        
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
            
            // 利用可能な会議室として記録（定員に関係なく）
            allAvailableRooms.push({
                name: roomName,
                calendarId: calendarId,
                capacity: capacity
            });
            
            // 完全収容可能な会議室として記録
            if (capacity >= numberOfAttendees) {
                suitableRooms.push({
                    name: roomName,
                    calendarId: calendarId,
                    capacity: capacity
                });
                writeLog('DEBUG', '条件適合会議室: ' + roomName + ' (定員: ' + capacity + ', カレンダーID: ' + calendarId + ')');
            }
        }
        
        // 完全収容可能な会議室がある場合
        if (suitableRooms.length > 0) {
            // 定員でソートして最小の適合会議室を選択
            suitableRooms.sort(function(a, b) {
                return a.capacity - b.capacity;
            });
            
            var selectedRoom = suitableRooms[0];
            writeLog('INFO', '最適会議室選択: ' + selectedRoom.name + ' (定員: ' + selectedRoom.capacity + 
                     ', 必要人数: ' + numberOfAttendees + ', 無駄席数: ' + (selectedRoom.capacity - numberOfAttendees) + ')');
            
            return {
                roomName: selectedRoom.name,
                isHybrid: false,
                maxCapacity: selectedRoom.capacity
            };
        }
        
        // 完全収容不可能な場合、最大定員の会議室でハイブリッド開催
        if (allAvailableRooms.length > 0) {
            // 定員でソート（降順：最大定員から）
            allAvailableRooms.sort(function(a, b) {
                return b.capacity - a.capacity;
            });
            
            var largestRoom = allAvailableRooms[0];
            var onlineParticipants = numberOfAttendees - largestRoom.capacity;
            
            writeLog('INFO', 'ハイブリッド開催選択: ' + largestRoom.name + ' (定員: ' + largestRoom.capacity + 
                     ', 必要人数: ' + numberOfAttendees + ', 会議室参加: ' + largestRoom.capacity + 
                     '名, オンライン参加: ' + onlineParticipants + '名)');
            
            return {
                roomName: largestRoom.name + ' + オンライン',
                isHybrid: true,
                maxCapacity: largestRoom.capacity,
                onlineCount: onlineParticipants,
                actualRoomName: largestRoom.name
            };
        }
        
        // 会議室が全く利用できない場合
        throw new Error('利用可能な会議室が見つかりませんでした。');
        
    } catch (e) {
        writeLog('ERROR', '会議室検索でエラー: ' + e.message);
        throw e;
    }
}

/**
 * 会議室を時間枠と併せて確保する（新会議室管理システム対応）
 * @param {number} numberOfAttendees - 参加者数
 * @param {Date} startTime - 開始時間
 * @param {Date} endTime - 終了時間
 * @param {string} trainingName - 研修名
 * @returns {Object} {roomName: string, isHybrid: boolean, capacity: number, onlineCount: number|null}
 */
function findAndReserveRoom(numberOfAttendees, startTime, endTime, trainingName) {
    writeLog('DEBUG', '会議室確保開始: 必要人数=' + numberOfAttendees + ', 時間=' + 
             Utilities.formatDate(startTime, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + 
             Utilities.formatDate(endTime, 'Asia/Tokyo', 'HH:mm') + ', 研修=' + (trainingName || '未指定'));
    
    // 参加者数が無効な場合はエラー
    if (!numberOfAttendees || numberOfAttendees <= 0) {
        throw new Error('参加者数が0以下です: ' + numberOfAttendees);
    }
    
    var roomManager = RoomReservationManager.getInstance();
    roomManager.logCurrentReservations(); // デバッグ用
    
    try {
        var sheet = SpreadsheetApp.openById(SPREADSHEET_IDS.ROOM_MASTER).getSheetByName(SHEET_NAMES.ROOM_MASTER);
        var lastRow = sheet.getLastRow();
        
        if (lastRow <= 1) {
            throw new Error('会議室マスタにデータがありません');
        }
        
        var data = sheet.getRange(2, 1, lastRow - 1, 4).getValues(); // D列まで取得
        
        var availableRooms = [];
        for (var i = 0; i < data.length; i++) {
            var row = data[i];
            var roomName = row[0];
            var calendarId = row[1];
            var capacity = row[2];
            var onbAvailable = row[3];
            
            if (onbAvailable !== '利用可能') continue;
            
            // 新予約管理システムで内部重複チェック
            if (!roomManager.isRoomAvailable(roomName, startTime, endTime)) {
                writeLog('DEBUG', '内部予約システムで使用中のためスキップ: ' + roomName);
                continue;
            }
            
            // Googleカレンダー上での重複チェック
            if (calendarId && isGoogleCalendarRoomAvailable(calendarId, startTime, endTime)) {
                availableRooms.push({
                    name: roomName,
                    capacity: capacity,
                    calendarId: calendarId,
                    resourceEmail: calendarId,
                    resourceName: roomName
                });
                writeLog('DEBUG', '利用可能会議室: ' + roomName + ' (定員: ' + capacity + ')');
            } else {
                writeLog('DEBUG', 'Googleカレンダーで使用中のためスキップ: ' + roomName);
            }
        }
        
        if (availableRooms.length === 0) {
            throw new Error('指定時間に利用可能な会議室が見つかりませんでした。');
        }

        // 参加者全員を収容できる会議室を検索
        var suitableRooms = availableRooms.filter(function(room) {
            return room.capacity >= numberOfAttendees;
        });
        
        var selectedRoom;
        var isHybrid = false;
        var onlineCount = 0;

        if (suitableRooms.length > 0) {
            // 8名超えの場合は、8名以上収容可能な会議室を優先選択
            if (numberOfAttendees > 8) {
                var eightPlusRooms = suitableRooms.filter(function(room) {
                    return room.capacity >= 8;
                });
                
                if (eightPlusRooms.length > 0) {
                    // 8名以上収容可能な会議室の中から最適なものを選択
                    eightPlusRooms.sort(function(a, b) { return a.capacity - b.capacity; });
                    selectedRoom = eightPlusRooms[0];
                    writeLog('INFO', '8名超え対応: 8名以上収容会議室選択: ' + selectedRoom.name + ' (定員: ' + selectedRoom.capacity + ', 必要人数: ' + numberOfAttendees + ')');
                } else {
                    // 8名以上収容可能な会議室がない場合は通常ロジック
                    suitableRooms.sort(function(a, b) { return a.capacity - b.capacity; });
                    selectedRoom = suitableRooms[0];
                    writeLog('WARN', '8名超えだが8名以上の会議室なし: ' + selectedRoom.name + ' (定員: ' + selectedRoom.capacity + ', 必要人数: ' + numberOfAttendees + ')');
                }
            } else {
                // 8名以下の場合は通常の最適会議室選択
                suitableRooms.sort(function(a, b) { return a.capacity - b.capacity; });
                selectedRoom = suitableRooms[0];
                writeLog('INFO', '最適会議室選択: ' + selectedRoom.name + ' (定員: ' + selectedRoom.capacity + ', 必要人数: ' + numberOfAttendees + ')');
            }
        } else {
            // 全員を収容できない場合、最大の会議室を選択してハイブリッド開催
            availableRooms.sort(function(a, b) { return b.capacity - a.capacity; });
            selectedRoom = availableRooms[0];
            isHybrid = true;
            onlineCount = numberOfAttendees - selectedRoom.capacity;
            writeLog('INFO', 'ハイブリッド開催で最大会議室選択: ' + selectedRoom.name + ' (定員: ' + selectedRoom.capacity + ', オンライン参加: ' + onlineCount + '名)');
        }
        
        // 新予約管理システムで会議室を予約
        if (!roomManager.reserveRoom(selectedRoom.name, selectedRoom.calendarId, startTime, endTime, trainingName || '未指定研修')) {
            throw new Error('会議室の予約に失敗しました: ' + selectedRoom.name);
        }
        
        return {
            roomName: selectedRoom.resourceName || selectedRoom.name,
            isHybrid: isHybrid,
            capacity: selectedRoom.capacity,
            onlineCount: onlineCount,
            calendarId: selectedRoom.calendarId
        };
        
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
 * 指定した時間枠で会議室が利用可能かチェック（RoomReservationManager使用）
 * @param {string} roomName - 会議室名
 * @param {Date} startTime - 開始時間
 * @param {Date} endTime - 終了時間
 * @returns {boolean} 利用可能かどうか
 */
function isRoomAvailable(roomName, startTime, endTime) {
    var roomManager = RoomReservationManager.getInstance();
    return roomManager.isRoomAvailable(roomName, startTime, endTime);
}

// グローバル変数：スケジュール管理用
var scheduledEvents = [];
// 注意: scheduledRooms配列は廃止し、RoomReservationManagerに統一

// =========================================
// 会議室予約管理システム
// =========================================

/**
 * 会議室予約管理クラス（シングルトンパターン）
 */
var RoomReservationManager = (function() {
    var instance = null;
    
    function createInstance() {
        var reservations = []; // 予約情報を格納
        
        return {
            /**
             * 予約情報をリセット
             */
            reset: function() {
                reservations = [];
                writeLog('INFO', '会議室予約管理をリセットしました');
            },
            
            /**
             * 会議室が指定時間に利用可能かチェック
             * @param {string} roomName - 会議室名
             * @param {Date} startTime - 開始時間
             * @param {Date} endTime - 終了時間
             * @returns {boolean} 利用可能かどうか
             */
            isRoomAvailable: function(roomName, startTime, endTime) {
                for (var i = 0; i < reservations.length; i++) {
                    var reservation = reservations[i];
                    if (reservation.roomName === roomName) {
                        // 時間重複チェック（終了時間 <= 開始時間 または 開始時間 >= 終了時間でなければ重複）
                        if (!(endTime <= reservation.startTime || startTime >= reservation.endTime)) {
                            writeLog('DEBUG', '会議室重複検出: ' + roomName + ' (' + 
                                     Utilities.formatDate(startTime, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + 
                                     Utilities.formatDate(endTime, 'Asia/Tokyo', 'HH:mm') + ') vs 既存予約(' +
                                     Utilities.formatDate(reservation.startTime, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + 
                                     Utilities.formatDate(reservation.endTime, 'Asia/Tokyo', 'HH:mm') + ')');
                            return false;
                        }
                    }
                }
                return true;
            },
            
            /**
             * 会議室を予約
             * @param {string} roomName - 会議室名
             * @param {string} calendarId - カレンダーID
             * @param {Date} startTime - 開始時間
             * @param {Date} endTime - 終了時間
             * @param {string} trainingName - 研修名
             * @returns {boolean} 予約成功かどうか
             */
            reserveRoom: function(roomName, calendarId, startTime, endTime, trainingName) {
                if (!this.isRoomAvailable(roomName, startTime, endTime)) {
                    writeLog('WARN', '会議室予約失敗（時間重複）: ' + roomName + ' for ' + trainingName);
                    return false;
                }
                
                var reservation = {
                    roomName: roomName,
                    calendarId: calendarId,
                    startTime: new Date(startTime),
                    endTime: new Date(endTime),
                    trainingName: trainingName,
                    reservedAt: new Date()
                };
                
                reservations.push(reservation);
                writeLog('INFO', '会議室予約成功: ' + roomName + ' (' + 
                         Utilities.formatDate(startTime, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + 
                         Utilities.formatDate(endTime, 'Asia/Tokyo', 'HH:mm') + ') for ' + trainingName);
                return true;
            },
            
            /**
             * 現在の予約状況を取得
             * @returns {Array} 予約情報の配列
             */
            getReservations: function() {
                return reservations.slice(); // コピーを返す
            },
            
            /**
             * 特定の研修の予約を削除
             * @param {string} trainingName - 研修名
             * @returns {boolean} 削除成功かどうか
             */
            cancelReservation: function(trainingName) {
                var originalLength = reservations.length;
                reservations = reservations.filter(function(reservation) {
                    return reservation.trainingName !== trainingName;
                });
                var deleted = originalLength - reservations.length;
                if (deleted > 0) {
                    writeLog('INFO', '会議室予約削除: ' + trainingName + ' (' + deleted + '件)');
                    return true;
                }
                return false;
            },
            
            /**
             * デバッグ用：現在の予約状況をログ出力
             */
            logCurrentReservations: function() {
                writeLog('DEBUG', '=== 現在の会議室予約状況 ===');
                if (reservations.length === 0) {
                    writeLog('DEBUG', '予約なし');
                } else {
                    for (var i = 0; i < reservations.length; i++) {
                        var res = reservations[i];
                        writeLog('DEBUG', (i + 1) + '. ' + res.roomName + ' (' + 
                                 Utilities.formatDate(res.startTime, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + 
                                 Utilities.formatDate(res.endTime, 'Asia/Tokyo', 'HH:mm') + ') - ' + res.trainingName);
                    }
                }
                writeLog('DEBUG', '=== 予約状況終了 ===');
            }
        };
    }
    
    return {
        getInstance: function() {
            if (!instance) {
                instance = createInstance();
            }
            return instance;
        }
    };
})();

/**
 * 全ての研修のカレンダーイベントを作成する（時間重複を回避）
 * @param {Array<Object>} trainingGroups - 研修グループの配列
 * @param {Date} hireDate - 入社日
 */
function createAllCalendarEvents(trainingGroups, hireDate) {
    writeLog('INFO', '全研修のカレンダーイベント作成開始');

    // スケジュール管理用配列をリセット
    scheduledEvents = [];

    // 会議室予約管理システムをリセット
    var roomManager = RoomReservationManager.getInstance();
    roomManager.reset();

    // 研修グループを実施日、実施順で明示的にソートする
    trainingGroups.sort(function(a, b) {
        var dayA = a.implementationDay || 999;
        var dayB = b.implementationDay || 999;
        if (dayA !== dayB) {
            return dayA - dayB;
        }
        var seqA = a.sequence || 999;
        var seqB = b.sequence || 999;
        return seqA - seqB;
    });
    writeLog('INFO', '研修グループを実施日・実施順でソートしました。');

    var scheduleResults = [];

    for (var i = 0; i < trainingGroups.length; i++) {
        var group = trainingGroups[i];

        writeLog('INFO', '研修処理開始: ' + group.name + ' (ユニークキー: ' + (group.uniqueKey || 'なし') + ')');

        // 改善されたユニークキーによる重複チェック
        var eventUniqueKey = generateEventUniqueKey(group);
        var isDuplicate = false;
        
        for (var j = 0; j < scheduledEvents.length; j++) {
            var existingEvent = scheduledEvents[j];
            var existingUniqueKey = existingEvent.uniqueKey || generateEventUniqueKey(existingEvent);
            
            if (eventUniqueKey === existingUniqueKey) {
                writeLog('WARN', '重複イベントを検出、スキップ: ' + group.name + ' (ユニークキー: ' + eventUniqueKey + ')');
                isDuplicate = true;
                break;
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
        var roomReservation = null; // 会議室予約情報を保持する変数

        if (group.needsRoom) {
            try {
                // 会議室定員は講師＋参加者の総数で計算（講師も席が必要）
                var lecturerCount = group.lecturerEmails ? group.lecturerEmails.length : 1; // 講師数を取得（デフォルト1）
                var totalAttendeeCount = participantCount + lecturerCount; // 講師分を追加
                writeLog('DEBUG', '会議室確保試行: 研修=' + group.name + ', 総参加者数=' + totalAttendeeCount + ' (講師' + lecturerCount + '名+参加者' + participantCount + '名), 時間=' + 
                         Utilities.formatDate(eventTime.start, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + 
                         Utilities.formatDate(eventTime.end, 'Asia/Tokyo', 'HH:mm'));
                
                roomReservation = findAndReserveRoom(totalAttendeeCount, eventTime.start, eventTime.end, group.name);
                roomName = roomReservation.roomName; // 実際の会議室名
                writeLog('INFO', '会議室確保成功: ' + roomName + ' (研修: ' + group.name + ', 総参加者: ' + totalAttendeeCount + '名)');
            } catch (e) {
                writeLog('ERROR', '会議室確保失敗: ' + e.message + ' (研修: ' + group.name + ')');
                roomName = '会議室未確保';
                roomError = '会議室確保失敗: ' + e.message;
            }
        } else {
            roomName = 'オンライン';
        }
        
        // カレンダーイベント作成
        try {
            // roomReservationオブジェクトを渡す
            createSingleCalendarEvent(group, roomReservation, eventTime.start, eventTime.end);
            
            // カレンダーイベントIDが正しく設定されているかチェック
            if (!group.calendarEventId) {
                writeLog('ERROR', 'カレンダーイベント作成後にIDが設定されていません: ' + group.name);
                throw new Error('カレンダーイベントIDが設定されませんでした');
            }
            
            // スケジュール管理配列に追加
            var scheduledEvent = {
                name: group.name,
                lecturer: group.lecturer,
                lecturerEmails: group.lecturerEmails,
                implementationDay: group.implementationDay,
                sequence: group.sequence,
                startTime: eventTime.start,
                endTime: eventTime.end,
                attendees: group.attendees.slice(),
                uniqueKey: eventUniqueKey,
                roomName: roomName,
                calendarEventId: group.calendarEventId
            };
            scheduledEvents.push(scheduledEvent);
            
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
 * 指定した時間枠で利用可能な会議室があるかチェックする（予約はしない）- 新予約管理システム対応
 * @param {number} numberOfAttendees - 参加者数
 * @param {Date} startTime - 開始時間
 * @param {Date} endTime - 終了時間
 * @returns {boolean} 利用可能な会議室があるかどうか
 */
function isAnyRoomAvailable(numberOfAttendees, startTime, endTime) {
    writeLog('DEBUG', '利用可能な会議室の存在チェック: 必要人数=' + numberOfAttendees + ', 時間=' + 
             Utilities.formatDate(startTime, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + 
             Utilities.formatDate(endTime, 'Asia/Tokyo', 'HH:mm'));
    
    var roomManager = RoomReservationManager.getInstance();
    
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
                // 予約管理システムで内部重複チェック
                if (!roomManager.isRoomAvailable(roomName, startTime, endTime)) {
                    writeLog('DEBUG', '予約管理システムで使用中のためスキップ: ' + roomName);
                    continue;
                }

                // Googleカレンダー上で空いているかチェック
                if (calendarId && isGoogleCalendarRoomAvailable(calendarId, startTime, endTime)) {
                    writeLog('DEBUG', '利用可能な会議室を発見: ' + roomName + ' (定員: ' + capacity + ', 必要人数: ' + numberOfAttendees + ')');
                    return true;
                } else {
                    writeLog('DEBUG', 'Googleカレンダーで使用中のためスキップ: ' + roomName);
                }
            } else {
                writeLog('DEBUG', '定員不足のためスキップ: ' + roomName + ' (定員: ' + capacity + ', 必要人数: ' + numberOfAttendees + ')');
            }
        }
        
        writeLog('DEBUG', '利用可能な会議室が見つかりませんでした（チェック対象: ' + data.length + '室）');
        return false;
    } catch (e) {
        writeLog('WARN', '会議室の空き状況チェックでエラー: ' + e.message);
        return false; // エラー時は利用不可とみなす
    }
}

/**
 * 講師の空き時間を考慮して適切な時間枠を見つける（実施日・実施順を考慮）
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
    var implementationDay = trainingGroup.implementationDay || 50;
    var sequence = trainingGroup.sequence || 100;

    writeLog('INFO', '時間枠検索開始（実施順厳密版）: ' + trainingGroup.name + ' (実施日: ' + implementationDay + ', 実施順: ' + sequence + ', 時間: ' + durationMinutes + '分)');

    // 実施日を計算
    var targetDate = calculateImplementationDate(hireDate, implementationDay);
    if (!targetDate || targetDate.getDay() === 0 || targetDate.getDay() === 6) {
        var reason = '対象日(' + Utilities.formatDate(targetDate, 'Asia/Tokyo', 'yyyy/MM/dd') + ')が週末または無効です';
        writeLog('WARN', reason);
        return null;
    }

    // 基準開始時間を設定
    var baseHour, baseMinute;
    if (implementationDay === 1) {
        baseHour = 15; baseMinute = 0;
    } else if (implementationDay === 2) {
        baseHour = 16; baseMinute = 0;
    } else {
        baseHour = 9; baseMinute = 0;
    }

    // 実施順に基づく厳密な開始時間計算
    var calculatedStartTime = calculateSequenceBasedStartTime(targetDate, implementationDay, sequence, durationMinutes, baseHour, baseMinute);
    
    writeLog('DEBUG', '実施順' + sequence + 'の計算開始時間: ' + Utilities.formatDate(calculatedStartTime, 'Asia/Tokyo', 'MM/dd HH:mm'));

    var proposedStart = new Date(calculatedStartTime.getTime());
    var proposedEnd = new Date(proposedStart.getTime() + (durationMinutes * 60 * 1000));

    // 営業時間チェック（すべての日で19時まで統一）
    var maxEndTime = 19;
    if (proposedEnd.getHours() > maxEndTime || (proposedEnd.getHours() === maxEndTime && proposedEnd.getMinutes() > 0)) {
        var reason = '計算された終了時刻(' + Utilities.formatDate(proposedEnd, 'Asia/Tokyo', 'HH:mm') + ')が営業時間外です';
        writeLog('WARN', reason);
        return null;
    }

    // 空き状況チェック
    var availability = isProposedTimeSlotAvailable(proposedStart, proposedEnd, trainingGroup);
    if (availability.available) {
        writeLog('INFO', '時間枠確保成功（実施順厳密）: ' + Utilities.formatDate(proposedStart, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + Utilities.formatDate(proposedEnd, 'Asia/Tokyo', 'HH:mm'));
        return { start: proposedStart, end: proposedEnd };
    } else {
        writeLog('WARN', '計算時間枠利用不可（' + availability.reason + '）、フォールバック検索を開始: ' + trainingGroup.name);
        
        // フォールバック: 15分間隔で次の空き時間を検索
        var fallbackResult = findFallbackTimeSlot(targetDate, implementationDay, sequence, durationMinutes, proposedStart, trainingGroup);
        if (fallbackResult) {
            writeLog('INFO', 'フォールバック時間枠確保成功: ' + Utilities.formatDate(fallbackResult.start, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + Utilities.formatDate(fallbackResult.end, 'Asia/Tokyo', 'HH:mm'));
            return fallbackResult;
        } else {
            writeLog('WARN', '時間枠確保失敗（フォールバック含む）: ' + trainingGroup.name);
            return null;
        }
    }
}

/**
 * 実施順に基づく厳密な開始時間を計算する
 * @param {Date} targetDate - 対象日
 * @param {number} implementationDay - 実施日
 * @param {number} sequence - 実施順
 * @param {number} durationMinutes - 研修時間（分）
 * @param {number} baseHour - 基準開始時間（時）
 * @param {number} baseMinute - 基準開始時間（分）
 * @returns {Date} 計算された開始時間
 */
function calculateSequenceBasedStartTime(targetDate, implementationDay, sequence, durationMinutes, baseHour, baseMinute) {
    var baseStartTime = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate(), baseHour, baseMinute);
    
    // 実施順1番の場合は基準時間から開始
    if (sequence === 1) {
        writeLog('DEBUG', '実施順1番のため基準時間から開始: ' + Utilities.formatDate(baseStartTime, 'Asia/Tokyo', 'HH:mm'));
        return baseStartTime;
    }

    // 同一日同一実施日の既存研修を実施順でソート
    var sameDayEvents = scheduledEvents.filter(function(event) {
        var eventDateStr = Utilities.formatDate(event.startTime, 'Asia/Tokyo', 'yyyy-MM-dd');
        var targetDateStr = Utilities.formatDate(targetDate, 'Asia/Tokyo', 'yyyy-MM-dd');
        return eventDateStr === targetDateStr && event.implementationDay === implementationDay;
    });

    if (sameDayEvents.length === 0) {
        writeLog('DEBUG', '同一日に既存研修なし、基準時間から開始: ' + Utilities.formatDate(baseStartTime, 'Asia/Tokyo', 'HH:mm'));
        return baseStartTime;
    }

    // 実施順でソート
    sameDayEvents.sort(function(a, b) {
        var aSeq = a.sequence || 999;
        var bSeq = b.sequence || 999;
        return aSeq - bSeq;
    });

    writeLog('DEBUG', '同一日既存研修(' + sameDayEvents.length + '件):');
    for (var i = 0; i < sameDayEvents.length; i++) {
        var event = sameDayEvents[i];
        writeLog('DEBUG', '  実施順' + (event.sequence || '未設定') + ': ' + event.name + ' (' + 
                 Utilities.formatDate(event.startTime, 'Asia/Tokyo', 'HH:mm') + '-' + 
                 Utilities.formatDate(event.endTime, 'Asia/Tokyo', 'HH:mm') + ')');
    }

    // 現在の実施順の直前の研修を探す
    var previousEvent = null;
    for (var i = 0; i < sameDayEvents.length; i++) {
        var event = sameDayEvents[i];
        var eventSequence = event.sequence || 999;
        
        if (eventSequence === sequence - 1) {
            previousEvent = event;
            break;
        }
    }

    if (previousEvent) {
        // 直前の研修の終了時刻から即座に開始
        var startTime = new Date(previousEvent.endTime.getTime());
        
        // 昼休み時間帯（12:00-13:00）を跨ぐ場合の調整
        var lunchStart = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate(), 12, 0);
        var lunchEnd = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate(), 13, 0);
        
        if (startTime >= lunchStart && startTime < lunchEnd) {
            startTime = lunchEnd;
            writeLog('DEBUG', '実施順' + sequence + ': 直前研修終了が昼休み中のため13:00から開始');
        }
        
        writeLog('DEBUG', '実施順' + sequence + ': 直前研修（実施順' + (sequence - 1) + '）終了時刻から開始: ' + 
                 Utilities.formatDate(startTime, 'Asia/Tokyo', 'HH:mm'));
        return startTime;
    } else {
        // 直前の実施順がない場合、最も遅い終了時刻の研修の後から開始
        var latestEvent = sameDayEvents[sameDayEvents.length - 1];
        var startTime = new Date(latestEvent.endTime.getTime());
        
        // 昼休み時間帯調整
        var lunchStart = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate(), 12, 0);
        var lunchEnd = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate(), 13, 0);
        
        if (startTime >= lunchStart && startTime < lunchEnd) {
            startTime = lunchEnd;
            writeLog('DEBUG', '実施順' + sequence + ': 最新研修終了が昼休み中のため13:00から開始');
        }
        
        writeLog('DEBUG', '実施順' + sequence + ': 最新研修終了時刻から開始: ' + 
                 Utilities.formatDate(startTime, 'Asia/Tokyo', 'HH:mm'));
        return startTime;
    }
}

/**
 * フォールバック時間枠検索（15分間隔で次の空き時間を検索）
 * @param {Date} targetDate - 対象日
 * @param {number} implementationDay - 実施日
 * @param {number} sequence - 実施順
 * @param {number} durationMinutes - 研修時間（分）
 * @param {Date} originalStart - 元の開始時間
 * @param {Object} trainingGroup - 研修グループ
 * @returns {Object} {time: {start, end}|null, reason: string}
 */
function findFallbackTimeSlot(targetDate, implementationDay, sequence, durationMinutes, originalStart, trainingGroup) {
    writeLog('INFO', 'フォールバック時間枠検索開始: ' + trainingGroup.name + ' (元時間: ' + 
             Utilities.formatDate(originalStart, 'Asia/Tokyo', 'MM/dd HH:mm') + ')');
    
    // 検索範囲の設定（すべての日で19時まで統一）
    var searchStart = new Date(originalStart.getTime());
    var maxEndTime = 19; // すべての日で19時まで統一
    var dayEnd = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate(), maxEndTime, 0);
    
    // 昼休み時間帯の設定
    var lunchStart = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate(), 12, 0);
    var lunchEnd = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate(), 13, 0);
    
    // 15分間隔で検索
    var currentStart = new Date(searchStart.getTime());
    var maxIterations = 50; // 無限ループ防止
    var iteration = 0;
    
    writeLog('DEBUG', 'フォールバック検索範囲: ' + 
             Utilities.formatDate(searchStart, 'Asia/Tokyo', 'MM/dd HH:mm') + ' - ' + 
             Utilities.formatDate(dayEnd, 'Asia/Tokyo', 'MM/dd HH:mm'));
    
    while (iteration < maxIterations) {
        iteration++;
        
        // 15分進める
        currentStart.setTime(currentStart.getTime() + (15 * 60 * 1000));
        var currentEnd = new Date(currentStart.getTime() + (durationMinutes * 60 * 1000));
        
        writeLog('DEBUG', 'フォールバック試行' + iteration + ': ' + 
                 Utilities.formatDate(currentStart, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + 
                 Utilities.formatDate(currentEnd, 'Asia/Tokyo', 'HH:mm'));
        
        // 営業時間外チェック
        if (currentEnd > dayEnd) {
            writeLog('DEBUG', 'フォールバック検索終了: 営業時間外に到達');
            break;
        }
        
        // 昼休み時間帯をスキップ
        if (isLunchTimeOverlap(currentStart, currentEnd)) {
            writeLog('DEBUG', 'フォールバック昼休みスキップ: ' + 
                     Utilities.formatDate(currentStart, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + 
                     Utilities.formatDate(currentEnd, 'Asia/Tokyo', 'HH:mm'));
            
            // 昼休み後に調整
            if (currentStart < lunchEnd) {
                currentStart = new Date(lunchEnd.getTime());
                currentEnd = new Date(currentStart.getTime() + (durationMinutes * 60 * 1000));
                writeLog('DEBUG', 'フォールバック昼休み後調整: ' + 
                         Utilities.formatDate(currentStart, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + 
                         Utilities.formatDate(currentEnd, 'Asia/Tokyo', 'HH:mm'));
            }
            continue;
        }
        
        // 空き状況チェック
        var availability = isProposedTimeSlotAvailable(currentStart, currentEnd, trainingGroup);
        if (availability.available) {
            writeLog('INFO', 'フォールバック時間枠発見: ' + 
                     Utilities.formatDate(currentStart, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + 
                     Utilities.formatDate(currentEnd, 'Asia/Tokyo', 'HH:mm') + ' (試行回数: ' + iteration + ')');
            return { start: currentStart, end: currentEnd };
        } else {
            writeLog('DEBUG', 'フォールバック時間枠利用不可: ' + availability.reason);
        }
    }
    
    // フォールバック検索でも見つからなかった場合
    var reason = 'フォールバック検索失敗: ' + iteration + '回試行後、利用可能な時間枠が見つかりませんでした';
    writeLog('WARN', reason + ' (研修: ' + trainingGroup.name + ')');
    return null;
}

/**
 * 提案された時間枠が利用可能かチェック（会議室込み）
 * @param {Date} proposedStart - 提案開始時間
 * @param {Date} proposedEnd - 提案終了時間
 * @param {Object} trainingGroup - 研修グループ
 * @returns {boolean} 利用可能かどうか
 */
function isProposedTimeSlotAvailable(proposedStart, proposedEnd, trainingGroup) {
    writeLog('DEBUG', '提案時間枠チェック: ' + trainingGroup.name + ' (' + 
             Utilities.formatDate(proposedStart, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + 
             Utilities.formatDate(proposedEnd, 'Asia/Tokyo', 'HH:mm') + ')');
    
    // 1. 参加者の時間重複チェック
    var timeSlotCheck = isTimeSlotAvailable(proposedStart, proposedEnd, trainingGroup);
    if (!timeSlotCheck.available) {
        writeLog('DEBUG', '参加者の時間重複により利用不可');
        return timeSlotCheck;
    }
    
    // 2. 会議室の要否をチェック
    if (trainingGroup.needsRoom) {
        var totalAttendees = (trainingGroup.attendees ? trainingGroup.attendees.length : 0) + 
                             (trainingGroup.lecturerEmails ? trainingGroup.lecturerEmails.length : (trainingGroup.lecturer ? 1 : 0));
        
        if (!isAnyRoomAvailable(totalAttendees, proposedStart, proposedEnd)) {
            writeLog('DEBUG', '利用可能な会議室がないため利用不可');
            return { available: false, reason: '利用可能な会議室なし' };
        }
    }
    
    writeLog('DEBUG', '提案時間枠利用可能');
    return { available: true, reason: '' };
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
 * 指定された時間帯が昼休み時間（12:00-13:00）と重複するかチェック
 * @param {Date} startTime - 開始時間
 * @param {Date} endTime - 終了時間
 * @returns {boolean} 昼休みと重複する場合true
 */
function isLunchTimeOverlap(startTime, endTime) {
    var lunchStart = new Date(startTime.getFullYear(), startTime.getMonth(), startTime.getDate(), 12, 0);
    var lunchEnd = new Date(startTime.getFullYear(), startTime.getMonth(), startTime.getDate(), 13, 0);
    
    // 重複チェック: 開始時間 < 昼休み終了時間 AND 終了時間 > 昼休み開始時間
    return (startTime < lunchEnd && endTime > lunchStart);
}

/**
 * 研修グループの信頼性の高いユニークキーを生成
 * @param {Object} trainingGroup - 研修グループオブジェクト
 * @returns {string} ユニークキー
 */
function generateEventUniqueKey(trainingGroup) {
    if (!trainingGroup || !trainingGroup.name) {
        return 'invalid_' + Math.random().toString(36).substring(2, 15);
    }
    
    // 基本要素
    var keyComponents = [
        trainingGroup.name.trim(),
        trainingGroup.implementationDay || 0,
        trainingGroup.sequence || 0
    ];
    
    // 講師情報（複数講師対応）
    var lecturerKey = '';
    if (trainingGroup.lecturerEmails && trainingGroup.lecturerEmails.length > 0) {
        lecturerKey = trainingGroup.lecturerEmails.sort().join(',');
    } else if (trainingGroup.lecturer) {
        lecturerKey = trainingGroup.lecturer;
    }
    keyComponents.push(lecturerKey);
    
    // 参加者情報（ソートして一意性を保証）
    if (trainingGroup.attendees && trainingGroup.attendees.length > 0) {
        var sortedAttendees = trainingGroup.attendees.slice().sort().join(',');
        keyComponents.push(sortedAttendees);
    }
    
    var uniqueKey = keyComponents.join('_');
    
    // 長すぎる場合はハッシュ化（簡易版）
    if (uniqueKey.length > 200) {
        uniqueKey = 'hash_' + Utilities.base64Encode(uniqueKey).substring(0, 50);
    }
    
    return uniqueKey;
}

/**
 * 指定した時間枠が利用可能かチェック（重複問題修正版）
 * @param {Date} proposedStart - 提案開始時間
 * @param {Date} proposedEnd - 提案終了時間
 * @param {Object} trainingGroup - チェック対象の研修グループ
 * @returns {boolean} 利用可能かどうか
 */
function isTimeSlotAvailable(proposedStart, proposedEnd, trainingGroup) {
    writeLog('DEBUG', '時間枠利用可能性チェック: ' + trainingGroup.name + ' (' + 
             Utilities.formatDate(proposedStart, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + 
             Utilities.formatDate(proposedEnd, 'Asia/Tokyo', 'MM/dd HH:mm') + ')');
    
    var proposedAttendees = trainingGroup.attendees || [];
    var lecturerEmails = trainingGroup.lecturerEmails || (trainingGroup.lecturer ? [trainingGroup.lecturer] : []);

    // 1. 既にこの実行でスケジュールされたイベントとの重複チェック（厳密化）
    for (var i = 0; i < scheduledEvents.length; i++) {
        var scheduledEvent = scheduledEvents[i];
        
        // まず時間重複があるかをチェック
        if (!(proposedEnd <= scheduledEvent.startTime || proposedStart >= scheduledEvent.endTime)) {
            writeLog('DEBUG', '時間重複発見: ' + trainingGroup.name + ' vs ' + scheduledEvent.name + 
                     ' (' + Utilities.formatDate(proposedStart, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + 
                     Utilities.formatDate(proposedEnd, 'Asia/Tokyo', 'HH:mm') + ' vs ' +
                     Utilities.formatDate(scheduledEvent.startTime, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + 
                     Utilities.formatDate(scheduledEvent.endTime, 'Asia/Tokyo', 'HH:mm') + ')');
            
            // **厳密チェック: 同一開始時刻は絶対に許可しない**
            if (proposedStart.getTime() === scheduledEvent.startTime.getTime()) {
                var reason = '同一開始時刻の重複: ' + scheduledEvent.name;
                writeLog('DEBUG', '同一開始時刻により利用不可: ' + trainingGroup.name + ' vs ' + scheduledEvent.name);
                return { available: false, reason: reason };
            }
            
            // 時間重複がある場合、参加者重複もチェック
            var scheduledAttendees = scheduledEvent.attendees || [];
            var hasOverlappingAttendee = proposedAttendees.some(function(attendee) {
                return scheduledAttendees.indexOf(attendee) !== -1;
            });
            
            if (hasOverlappingAttendee) {
                var reason = '参加者重複: ' + scheduledEvent.name;
                writeLog('DEBUG', '時間・参加者重複により利用不可: ' + trainingGroup.name + ' vs ' + scheduledEvent.name);
                return { available: false, reason: reason };
            }

            // 講師の重複もチェック
            var scheduledLecturers = scheduledEvent.lecturerEmails || (scheduledEvent.lecturer ? [scheduledEvent.lecturer] : []);
            var hasOverlappingLecturer = lecturerEmails.some(function(lecturer) {
                return scheduledLecturers.indexOf(lecturer) !== -1;
            });

            if (hasOverlappingLecturer) {
                 var reason = '講師重複: ' + scheduledEvent.name;
                 writeLog('DEBUG', '時間・講師重複により利用不可: ' + trainingGroup.name + ' vs ' + scheduledEvent.name);
                 return { available: false, reason: reason };
            }
        }
    }
    
    // 2. 講師のGoogleカレンダーとの重複チェック
    for (var j = 0; j < lecturerEmails.length; j++) {
        var email = lecturerEmails[j];
        if (email && email.trim() !== '') {
            var lecturerCheck = isLecturerAvailable(email, proposedStart, proposedEnd);
            if (!lecturerCheck.available) {
                writeLog('DEBUG', '講師の予定と重複により利用不可: ' + email);
                return lecturerCheck;
            }
        }
    }
    
    writeLog('DEBUG', '時間枠利用可能: ' + trainingGroup.name);
    return { available: true, reason: '' };
}

/**
 * 講師が指定時間に空いているかチェック（終日イベント除外対応）
 * @param {string} lecturerEmail - 講師のメールアドレス
 * @param {Date} proposedStart - 提案開始時間
 * @param {Date} proposedEnd - 提案終了時間
 * @returns {Object} {available: boolean, reason: string}
 */
function isLecturerAvailable(lecturerEmail, proposedStart, proposedEnd) {
    try {
        writeLog('DEBUG', '講師カレンダーチェック開始: ' + lecturerEmail);
        
        var lecturerCalendar = CalendarApp.getCalendarById(lecturerEmail);
        if (!lecturerCalendar) {
            writeLog('WARN', '講師のカレンダーにアクセスできません: ' + lecturerEmail);
            return { available: true, reason: '' }; // アクセスできない場合は空いているとみなす
        }
        
        var existingEvents = lecturerCalendar.getEvents(proposedStart, proposedEnd);
        if (existingEvents.length === 0) {
            writeLog('DEBUG', '講師カレンダーに予定なし');
            return { available: true, reason: '' };
        }
        
        writeLog('DEBUG', '講師の既存予定候補: ' + existingEvents.length + '件');
        
        // 実際に時間重複している予定のみをチェック
        for (var i = 0; i < existingEvents.length; i++) {
            var event = existingEvents[i];
            var eventStart = event.getStartTime();
            var eventEnd = event.getEndTime();
            
            // 終日イベントを除外
            if (isAllDayEvent(event)) {
                writeLog('DEBUG', '終日イベントのためスキップ: ' + event.getTitle());
                continue;
            }
            
            // 実際の時間重複チェック
            if (!(proposedEnd <= eventStart || proposedStart >= eventEnd)) {
                var reason = '講師予定重複(' + lecturerEmail.split('@')[0] + '): ' + event.getTitle();
                writeLog('DEBUG', '講師予定と時間重複: ' + event.getTitle() + ' (' + 
                         Utilities.formatDate(eventStart, 'Asia/Tokyo', 'MM/dd HH:mm') + 
                         '-' + Utilities.formatDate(eventEnd, 'Asia/Tokyo', 'HH:mm') + ')');
                return { available: false, reason: reason };
            }
        }
        
        writeLog('DEBUG', '講師カレンダー重複なし');
        return { available: true, reason: '' };
        
    } catch (e) {
        writeLog('WARN', '講師カレンダーチェックでエラー: ' + e.message + ' (講師: ' + lecturerEmail + ')');
        // エラーの場合は空いているとみなす（厳格すぎるより寛容に）
        return { available: true, reason: '' };
    }
}

/**
 * イベントが終日イベントかどうかを判定
 * @param {GoogleAppsScript.Calendar.CalendarEvent} event - カレンダーイベント
 * @returns {boolean} 終日イベントの場合true
 */
function isAllDayEvent(event) {
    if (event.isAllDayEvent()) {
        return true;
    }
    
    var eventStart = event.getStartTime();
    var eventEnd = event.getEndTime();
    var eventDuration = eventEnd.getTime() - eventStart.getTime();
    
    // 24時間以上のイベントまたは00:00-00:00のイベントを終日イベントとみなす
    return (eventDuration >= 24 * 60 * 60 * 1000) ||
           (eventStart.getHours() === 0 && eventStart.getMinutes() === 0 &&
            eventEnd.getHours() === 0 && eventEnd.getMinutes() === 0);
}

/**
 * 単一のカレンダーイベントを作成する
 * @param {Object} trainingDetails - 研修情報
 * @param {string|null} roomName - 会議室名
 * @param {Date} startTime - 開始時間
 * @param {Date} endTime - 終了時間
 */
function createSingleCalendarEvent(trainingDetails, roomReservation, startTime, endTime) {
    var title = trainingDetails.name;
    var roomName = roomReservation ? roomReservation.roomName : 'オンライン';
    
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
    if (roomReservation && roomName && roomName !== 'オンライン' && roomName !== '会議室未確保' && roomName !== '実施不要') {
        // roomReservationから直接カレンダーIDを取得
        if (roomReservation.calendarId) {
            roomResourceEmail = roomReservation.calendarId;
            validEmails.push(roomResourceEmail);
            writeLog('INFO', '会議室カレンダーIDを参加者に追加: ' + roomResourceEmail);
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
    if (roomResourceEmail) {
        locationString = ''; // 会議室リソースが設定されていれば、locationは自動的に設定される
    }

    var description = trainingDetails.memo || '';
    
    var options = {
        description: description,
        guests: validEmails.join(','),
        location: locationString,
    };

    // ハイブリッド開催の場合、Google Meetのリンクを生成
    if (roomReservation && (roomReservation.isHybrid || roomName === 'オンライン')) {
        options.conferenceDataVersion = 1;
    }
    
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
 * 研修グループをインクリメンタルに処理してマッピングシートを更新する
 * @param {Array<Object>} trainingGroups - 研修グループの配列
 * @param {Array<Object>} allNewHires - 全入社者の配列
 * @param {Date} hireDate - 入社日
 * @param {Object} mappingSheet - 更新対象のマッピングシート
 */
function processTrainingGroupsIncrementally(trainingGroups, allNewHires, hireDate, mappingSheet) {
    writeLog('INFO', 'インクリメンタル処理開始: ' + trainingGroups.length + '件の研修を処理');
    
    // スケジュール管理用配列をリセット
    scheduledEvents = [];
    
    // 会議室予約管理システムをリセット
    var roomManager = RoomReservationManager.getInstance();
    roomManager.reset();
    
    var successCount = 0;
    var errorCount = 0;
    
    for (var i = 0; i < trainingGroups.length; i++) {
        var group = trainingGroups[i];
        var rowIndex = i + 2; // スプレッドシートの行番号（ヘッダー行の次から）
        
        writeLog('INFO', '処理中 (' + (i + 1) + '/' + trainingGroups.length + '): ' + group.name);
        
        try {
            // 処理状況を「処理中」に更新
            updateMappingSheetRow(mappingSheet, rowIndex, {
                status: '処理中...',
                roomName: '確保中...',
                schedule: '計算中...'
            });
            
            // 参加者数チェック
            var participantCount = 0;
            if (group.attendees && group.attendees.length > 0) {
                for (var k = 0; k < group.attendees.length; k++) {
                    if (group.attendees[k] !== group.lecturer) {
                        participantCount++;
                    }
                }
            }
            
            if (participantCount === 0) {
                writeLog('INFO', '参加者が0人のため研修をスキップ: ' + group.name);
                updateMappingSheetRow(mappingSheet, rowIndex, {
                    status: 'スキップ（参加者0名）',
                    roomName: '実施不要',
                    schedule: 'N/A'
                });
                continue;
            }
            
            // 時間枠を確保
            var eventTime = findAvailableTimeSlot(group, hireDate);
            if (!eventTime) {
                writeLog('ERROR', '時間枠確保失敗: ' + group.name);
                updateMappingSheetRow(mappingSheet, rowIndex, {
                    status: '失敗（時間枠未確保）',
                    roomName: group.needsRoom ? '会議室未確保' : 'オンライン',
                    schedule: '時間枠未確保',
                    errorReason: '利用可能な時間枠が見つかりませんでした'
                });
                errorCount++;
                continue;
            }
            
            // 会議室確保
            var roomName = null;
            var roomReservation = null;
            if (group.needsRoom) {
                try {
                    var lecturerCount = group.lecturerEmails ? group.lecturerEmails.length : 1;
                    var totalAttendeeCount = participantCount + lecturerCount;
                    roomReservation = findAndReserveRoom(totalAttendeeCount, eventTime.start, eventTime.end, group.name);
                    roomName = roomReservation.roomName;
                    writeLog('INFO', '会議室確保成功: ' + roomName);
                } catch (e) {
                    writeLog('ERROR', '会議室確保失敗: ' + e.message);
                    roomName = '会議室未確保';
                }
            } else {
                roomName = 'オンライン';
            }
            
            // カレンダーイベント作成
            createSingleCalendarEvent(group, roomReservation, eventTime.start, eventTime.end);
            
            // 成功時の更新
            var scheduleStr = Utilities.formatDate(eventTime.start, 'Asia/Tokyo', 'MM/dd(E) HH:mm') + 
                            '-' + Utilities.formatDate(eventTime.end, 'Asia/Tokyo', 'HH:mm');
            
            updateMappingSheetRow(mappingSheet, rowIndex, {
                status: '成功',
                roomName: roomName,
                schedule: scheduleStr,
                calendarId: group.calendarEventId || ''
            });
            
            successCount++;
            writeLog('INFO', '研修処理成功: ' + group.name + ' (ID: ' + group.calendarEventId + ')');
            
        } catch (e) {
            writeLog('ERROR', '研修処理失敗: ' + group.name + ' - ' + e.message);
            updateMappingSheetRow(mappingSheet, rowIndex, {
                status: '失敗: ' + e.message,
                roomName: group.needsRoom ? '会議室未確保' : 'オンライン',
                schedule: 'エラー'
            });
            errorCount++;
        }
        
        // 少し待機（API制限対策）
        if (i < trainingGroups.length - 1) {
            Utilities.sleep(1000); // 1秒待機
        }
    }
    
    // 最終サマリーをシートに追加（拡張検証付き）
    addProcessingSummary(mappingSheet, trainingGroups.length, successCount, errorCount, allNewHires, trainingGroups);
    
    writeLog('INFO', 'インクリメンタル処理完了: 成功=' + successCount + ', 失敗=' + errorCount + ', 総数=' + trainingGroups.length);
}

/**
 * マッピングシートの特定行を更新する
 * @param {Object} mappingSheet - 更新対象のシート
 * @param {number} rowIndex - 行番号
 * @param {Object} updates - 更新内容 {status, roomName, schedule, calendarId}
 */
function updateMappingSheetRow(mappingSheet, rowIndex, updates) {
    try {
        if (updates.roomName !== undefined) {
            mappingSheet.getRange(rowIndex, 6).setValue(updates.roomName); // F列: 会議室名
        }
        if (updates.schedule !== undefined) {
            mappingSheet.getRange(rowIndex, 7).setValue(updates.schedule); // G列: 研修実施日時
        }
        if (updates.calendarId !== undefined) {
            mappingSheet.getRange(rowIndex, 8).setValue(updates.calendarId); // H列: カレンダーID
        }
        if (updates.status !== undefined) {
            var statusCell = mappingSheet.getRange(rowIndex, 9); // I列: 処理状況
            statusCell.setValue(updates.status);
            
            // 処理状況に応じた背景色設定
            if (updates.status === '成功') {
                statusCell.setBackground('#d4edda').setFontColor('#155724');
            } else if (updates.status.indexOf('失敗') !== -1 || updates.status.indexOf('エラー') !== -1) {
                statusCell.setBackground('#f8d7da').setFontColor('#721c24');
            } else if (updates.status.indexOf('処理中') !== -1) {
                statusCell.setBackground('#fff3cd').setFontColor('#856404');
            } else if (updates.status.indexOf('スキップ') !== -1) {
                statusCell.setBackground('#e2e3e5').setFontColor('#383d41');
            }
        }
        
        // 表示を強制的に更新
        SpreadsheetApp.flush();
        
    } catch (e) {
        writeLog('ERROR', 'マッピングシート行更新でエラー: ' + e.message + ' (行: ' + rowIndex + ')');
    }
}

/**
 * 処理サマリーをシートに追加する（拡張版）
 * @param {Object} mappingSheet - 対象シート
 * @param {number} totalCount - 総数
 * @param {number} successCount - 成功数
 * @param {number} errorCount - エラー数
 * @param {Array<Object>} newHires - 入社者データ（検証用）
 * @param {Array<Object>} trainingGroups - 研修グループデータ（検証用）
 */
function addProcessingSummary(mappingSheet, totalCount, successCount, errorCount, newHires, trainingGroups) {
    var lastRow = mappingSheet.getLastRow();
    var summaryStartRow = lastRow + 2;
    
    // 基本サマリー
    mappingSheet.getRange(summaryStartRow, 1).setValue('【処理サマリー】');
    mappingSheet.getRange(summaryStartRow + 1, 1).setValue('総研修数:');
    mappingSheet.getRange(summaryStartRow + 1, 2).setValue(totalCount);
    mappingSheet.getRange(summaryStartRow + 2, 1).setValue('成功:');
    mappingSheet.getRange(summaryStartRow + 2, 2).setValue(successCount);
    mappingSheet.getRange(summaryStartRow + 3, 1).setValue('失敗:');
    mappingSheet.getRange(summaryStartRow + 3, 2).setValue(errorCount);
    mappingSheet.getRange(summaryStartRow + 4, 1).setValue('処理完了日時:');
    mappingSheet.getRange(summaryStartRow + 4, 2).setValue(new Date());
    
    // 基本サマリーのスタイル設定
    var basicSummaryRange = mappingSheet.getRange(summaryStartRow, 1, 5, 2);
    basicSummaryRange.setBackground('#f0f0f0');
    basicSummaryRange.setFontWeight('bold');
    
    // 拡張検証セクション
    var currentRow = summaryStartRow + 6;
    
    // 職位別研修数検証
    if (newHires && trainingGroups) {
        try {
            writeLog('INFO', '拡張検証レポート作成開始');
            
            // 職位別研修数検証
            var validationResults = validateTrainingCountByPosition(newHires, trainingGroups);
            
            mappingSheet.getRange(currentRow, 1).setValue('【職位別研修数検証】');
            currentRow++;
            
            var overallIcon = validationResults.overall ? '✅' : '❌';
            var overallStatus = validationResults.overall ? '正常' : '異常';
            mappingSheet.getRange(currentRow, 1).setValue('全体検証結果:');
            mappingSheet.getRange(currentRow, 2).setValue(overallIcon + ' ' + overallStatus);
            
            // 検証結果に応じて背景色を設定
            var statusCell = mappingSheet.getRange(currentRow, 2);
            if (validationResults.overall) {
                statusCell.setBackground('#d4edda').setFontColor('#155724');
            } else {
                statusCell.setBackground('#f8d7da').setFontColor('#721c24');
            }
            currentRow++;
            
            // 個人別詳細
            if (validationResults.details && validationResults.details.length > 0) {
                for (var i = 0; i < validationResults.details.length; i++) {
                    var detail = validationResults.details[i];
                    var icon = detail.isValid ? '✅' : '❌';
                    var detailText = detail.name + '(' + detail.rank + '/' + detail.experience + '): 期待数' + detail.expected + ' → 実際数' + detail.actual + ' ' + icon;
                    
                    mappingSheet.getRange(currentRow, 1).setValue('- ' + detail.name + ':');
                    mappingSheet.getRange(currentRow, 2).setValue('期待数' + detail.expected + ' → 実際数' + detail.actual + ' ' + icon);
                    
                    // 個別結果の背景色設定
                    var detailCell = mappingSheet.getRange(currentRow, 2);
                    if (detail.isValid) {
                        detailCell.setBackground('#e8f5e8');
                    } else {
                        detailCell.setBackground('#ffeaea');
                    }
                    currentRow++;
                }
            }
            
            currentRow++; // 空行
            
            // カレンダー重複チェック
            var conflictResults = validateCalendarTimeSlots(trainingGroups);
            
            mappingSheet.getRange(currentRow, 1).setValue('【カレンダー重複チェック】');
            currentRow++;
            
            var conflictIcon = conflictResults.hasConflicts ? '❌' : '✅';
            var conflictStatus = conflictResults.hasConflicts ? 'あり' : 'なし';
            mappingSheet.getRange(currentRow, 1).setValue('時間重複:');
            mappingSheet.getRange(currentRow, 2).setValue(conflictIcon + ' ' + conflictStatus);
            
            // 重複結果に応じて背景色を設定
            var conflictCell = mappingSheet.getRange(currentRow, 2);
            if (!conflictResults.hasConflicts) {
                conflictCell.setBackground('#d4edda').setFontColor('#155724');
            } else {
                conflictCell.setBackground('#f8d7da').setFontColor('#721c24');
            }
            currentRow++;
            
            // チェック対象数
            mappingSheet.getRange(currentRow, 1).setValue('チェック対象イベント数:');
            mappingSheet.getRange(currentRow, 2).setValue(conflictResults.checkedEvents + '/' + conflictResults.totalEvents);
            currentRow++;
            
            // 重複詳細
            if (conflictResults.conflicts && conflictResults.conflicts.length > 0) {
                mappingSheet.getRange(currentRow, 1).setValue('重複詳細:');
                currentRow++;
                
                for (var j = 0; j < Math.min(5, conflictResults.conflicts.length); j++) { // 最大5件まで表示
                    var conflict = conflictResults.conflicts[j];
                    mappingSheet.getRange(currentRow, 1).setValue('- ' + conflict.event1 + ' vs ' + conflict.event2);
                    mappingSheet.getRange(currentRow, 2).setValue(conflict.time1 + ' / ' + conflict.time2);
                    mappingSheet.getRange(currentRow, 2).setBackground('#ffeaea');
                    currentRow++;
                }
                
                if (conflictResults.conflicts.length > 5) {
                    mappingSheet.getRange(currentRow, 1).setValue('... 他' + (conflictResults.conflicts.length - 5) + '件');
                    currentRow++;
                }
            }
            
            // 拡張検証セクションのスタイル設定
            var validationHeaderRange = mappingSheet.getRange(summaryStartRow + 6, 1, currentRow - summaryStartRow - 6, 2);
            validationHeaderRange.setFontWeight('normal');
            
            // ヘッダー行を太字に
            mappingSheet.getRange(summaryStartRow + 6, 1).setFontWeight('bold'); // 【職位別研修数検証】
            if (currentRow > summaryStartRow + 8) {
                // 【カレンダー重複チェック】の行を探して太字に
                for (var row = summaryStartRow + 7; row < currentRow; row++) {
                    var cellValue = mappingSheet.getRange(row, 1).getValue();
                    if (cellValue && cellValue.toString().indexOf('【カレンダー重複チェック】') !== -1) {
                        mappingSheet.getRange(row, 1).setFontWeight('bold');
                        break;
                    }
                }
            }
            
            writeLog('INFO', '拡張検証レポート作成完了');
            
        } catch (e) {
            writeLog('ERROR', '拡張検証レポート作成でエラー: ' + e.message);
            
            // エラー時のフォールバック表示
            mappingSheet.getRange(currentRow, 1).setValue('【拡張検証】');
            currentRow++;
            mappingSheet.getRange(currentRow, 1).setValue('検証エラー:');
            mappingSheet.getRange(currentRow, 2).setValue(e.message);
            mappingSheet.getRange(currentRow, 2).setBackground('#f8d7da').setFontColor('#721c24');
        }
    }
}

// =========================================
// 拡張検証機能
// =========================================

/**
 * 職位別研修数を検証する
 * @param {Array<Object>} newHires - 入社者データ
 * @param {Array<Object>} trainingGroups - 研修グループデータ
 * @returns {Object} 検証結果
 */
function validateTrainingCountByPosition(newHires, trainingGroups) {
    writeLog('INFO', '=== 職位別研修数検証開始 ===');
    
    var validationResults = {
        overall: true,
        details: [],
        expectedCounts: {},
        actualCounts: {}
    };
    
    try {
        // 研修マスタから期待研修数を計算
        var expectedCounts = calculateExpectedTrainingCounts();
        validationResults.expectedCounts = expectedCounts;
        
        // 各入社者について検証
        for (var i = 0; i < newHires.length; i++) {
            var hire = newHires[i];
            var pattern = determineTrainingPattern(hire.rank, hire.experience);
            
            if (!pattern) {
                writeLog('WARN', '研修パターンが決定できません: ' + hire.name + ' (職位: ' + hire.rank + ', 経験: ' + hire.experience + ', 所属: ' + hire.department + ')');
                continue;
            }
            
            // この入社者が参加する研修数をカウント
            var actualCount = 0;
            for (var j = 0; j < trainingGroups.length; j++) {
                var group = trainingGroups[j];
                if (group.attendees && group.attendees.indexOf(hire.email) !== -1) {
                    actualCount++;
                }
            }
            
            var expectedCount = expectedCounts[pattern] || 0;
            var isValid = (actualCount === expectedCount);
            
            if (!isValid) {
                validationResults.overall = false;
            }
            
            validationResults.details.push({
                name: hire.name,
                rank: hire.rank,
                experience: hire.experience,
                pattern: pattern,
                expected: expectedCount,
                actual: actualCount,
                isValid: isValid
            });
            
            writeLog('INFO', '検証結果: ' + hire.name + ' (' + hire.rank + '/' + hire.experience + '/' + hire.department + ', パターン' + pattern + ') - 期待数: ' + expectedCount + ', 実際数: ' + actualCount + ', 判定: ' + (isValid ? '✅' : '❌'));
        }
        
        writeLog('INFO', '職位別研修数検証完了: 全体判定=' + (validationResults.overall ? '正常' : '異常'));
        
    } catch (e) {
        writeLog('ERROR', '職位別研修数検証でエラー: ' + e.message);
        validationResults.overall = false;
        validationResults.error = e.message;
    }
    
    return validationResults;
}

/**
 * 研修マスタから各パターンの期待研修数を計算する
 * @returns {Object} パターン別期待研修数
 */
function calculateExpectedTrainingCounts() {
    writeLog('DEBUG', '期待研修数計算開始');
    
    var expectedCounts = {
        'A': 0, // 未経験者向け
        'B': 0, // 経験C/SC向け
        'C': 0, // M向け
        'D': 0  // SMup向け
    };
    
    try {
        var masterSheet = SpreadsheetApp.openById(SPREADSHEET_IDS.TRAINING_MASTER).getSheetByName(SHEET_NAMES.TRAINING_MASTER);
        var lastRow = masterSheet.getLastRow();
        
        if (lastRow <= 4) {
            writeLog('WARN', '研修マスタにデータがありません');
            return expectedCounts;
        }
        
        var lastCol = masterSheet.getLastColumn();
        var maxCol = Math.min(lastCol, 21);
        var masterData = masterSheet.getRange(5, 1, lastRow - 4, maxCol).getValues();
        
        for (var i = 0; i < masterData.length; i++) {
            var row = masterData[i];
            
            var lv1 = row[0];           // A列: Lv.1
            var trainingName = row[2];  // C列: 研修名称
            var patternA = row[3];      // D列: Aパターン
            var patternB = row[4];      // E列: Bパターン
            var patternC = row[5];      // F列: Cパターン
            var patternD = row[6];      // G列: Dパターン
            
            // 対象外フィルタリング
            if (!trainingName || trainingName.trim() === '') continue;
            if (lv1 !== 'DX ONB' && lv1 !== 'ビジネススキル研修') continue;
            
            // 各パターンでカウント
            if (patternA === '●') expectedCounts['A']++;
            if (patternB === '●') expectedCounts['B']++;
            if (patternC === '●') expectedCounts['C']++;
            if (patternD === '●') expectedCounts['D']++;
        }
        
        writeLog('DEBUG', '期待研修数: A=' + expectedCounts['A'] + ', B=' + expectedCounts['B'] + ', C=' + expectedCounts['C'] + ', D=' + expectedCounts['D']);
        
    } catch (e) {
        writeLog('ERROR', '期待研修数計算でエラー: ' + e.message);
    }
    
    return expectedCounts;
}

/**
 * カレンダー時間枠の重複をチェックする
 * @param {Array<Object>} trainingGroups - 研修グループデータ
 * @returns {Object} 重複チェック結果
 */
function validateCalendarTimeSlots(trainingGroups) {
    writeLog('INFO', '=== カレンダー時間枠重複チェック開始 ===');
    
    var conflictResults = {
        hasConflicts: false,
        conflicts: [],
        totalEvents: 0,
        checkedEvents: 0
    };
    
    try {
        var scheduledEvents = [];
        
        // スケジュールされたイベントを収集
        for (var i = 0; i < trainingGroups.length; i++) {
            var group = trainingGroups[i];
            if (group.calendarEventId) {
                try {
                    var event = CalendarApp.getEventById(group.calendarEventId);
                    if (event) {
                        scheduledEvents.push({
                            name: group.name,
                            eventId: group.calendarEventId,
                            startTime: event.getStartTime(),
                            endTime: event.getEndTime(),
                            attendees: group.attendees || []
                        });
                        conflictResults.checkedEvents++;
                    }
                } catch (e) {
                    writeLog('WARN', 'カレンダーイベント取得失敗: ' + group.name + ' (ID: ' + group.calendarEventId + ') - ' + e.message);
                }
            }
        }
        
        conflictResults.totalEvents = scheduledEvents.length;
        writeLog('DEBUG', 'チェック対象イベント数: ' + conflictResults.totalEvents);
        
        // 時間重複チェック
        for (var i = 0; i < scheduledEvents.length; i++) {
            for (var j = i + 1; j < scheduledEvents.length; j++) {
                var event1 = scheduledEvents[i];
                var event2 = scheduledEvents[j];
                
                // 時間重複チェック
                var timeOverlap = !(event1.endTime <= event2.startTime || event1.startTime >= event2.endTime);
                
                if (timeOverlap) {
                    // 参加者重複チェック
                    var attendeeOverlap = false;
                    for (var k = 0; k < event1.attendees.length; k++) {
                        if (event2.attendees.indexOf(event1.attendees[k]) !== -1) {
                            attendeeOverlap = true;
                            break;
                        }
                    }
                    
                    if (attendeeOverlap) {
                        conflictResults.hasConflicts = true;
                        conflictResults.conflicts.push({
                            event1: event1.name,
                            event2: event2.name,
                            time1: Utilities.formatDate(event1.startTime, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + Utilities.formatDate(event1.endTime, 'Asia/Tokyo', 'HH:mm'),
                            time2: Utilities.formatDate(event2.startTime, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + Utilities.formatDate(event2.endTime, 'Asia/Tokyo', 'HH:mm'),
                            conflictType: '参加者重複あり'
                        });
                        
                        writeLog('WARN', '時間重複検出: ' + event1.name + ' vs ' + event2.name + ' (参加者重複あり)');
                    }
                }
            }
        }
        
        writeLog('INFO', 'カレンダー時間枠重複チェック完了: 重複=' + (conflictResults.hasConflicts ? 'あり(' + conflictResults.conflicts.length + '件)' : 'なし'));
        
    } catch (e) {
        writeLog('ERROR', 'カレンダー時間枠重複チェックでエラー: ' + e.message);
        conflictResults.error = e.message;
    }
    
    return conflictResults;
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