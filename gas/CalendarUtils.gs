// =========================================
// カレンダー関連ユーティリティ（リファクタリング版）
// =========================================

writeLog('INFO', 'CalendarUtils.gs の読み込みを開始しました');

// =========================================
// 実施順管理システム
// =========================================

/**
 * 実施順を管理するクラス（シングルトン）
 */
var SequenceManager = (function() {
    var instance = null;
    
    function createInstance() {
        var scheduledTrainings = []; // 実施順を管理するための配列
        
        return {
            /**
             * スケジュール配列をリセット
             */
            reset: function() {
                scheduledTrainings = [];
                writeLog('INFO', '実施順管理をリセットしました');
            },
            
            /**
             * 研修を実施順に追加
             * @param {Object} training - 研修情報
             */
            addTraining: function(training) {
                var trainingInfo = {
                    name: training.name,
                    implementationDay: training.implementationDay || 50,
                    sequence: training.sequence || 999,
                    startTime: training.startTime,
                    endTime: training.endTime,
                    attendees: training.attendees || [],
                    lecturerEmails: training.lecturerEmails || [],
                    calendarEventId: training.calendarEventId
                };
                
                scheduledTrainings.push(trainingInfo);
                
                // 実施日と実施順でソート
                scheduledTrainings.sort(function(a, b) {
                    if (a.implementationDay !== b.implementationDay) {
                        return a.implementationDay - b.implementationDay;
                    }
                    return a.sequence - b.sequence;
                });
                
                writeLog('DEBUG', '実施順管理に追加: ' + training.name + ' (実施日: ' + trainingInfo.implementationDay + ', 実施順: ' + trainingInfo.sequence + ')');
            },
            
            /**
             * 指定した実施日・実施順の直前の研修を取得
             * @param {number} implementationDay - 実施日
             * @param {number} sequence - 実施順
             * @returns {Object|null} 直前の研修情報
             */
            getPreviousTraining: function(implementationDay, sequence) {
                var previousTraining = null;
                
                for (var i = 0; i < scheduledTrainings.length; i++) {
                    var training = scheduledTrainings[i];
                    if (training.implementationDay === implementationDay && training.sequence === sequence - 1) {
                        previousTraining = training;
                        break;
                    }
                }
                
                writeLog('DEBUG', '直前研修検索: 実施日' + implementationDay + '実施順' + sequence + ' → ' + (previousTraining ? previousTraining.name : 'なし'));
                return previousTraining;
            },
            
            /**
             * 指定した実施日の最新研修を取得
             * @param {number} implementationDay - 実施日
             * @returns {Object|null} 最新の研修情報
             */
            getLatestTrainingInDay: function(implementationDay) {
                var latestTraining = null;
                var latestSequence = -1;
                
                for (var i = 0; i < scheduledTrainings.length; i++) {
                    var training = scheduledTrainings[i];
                    if (training.implementationDay === implementationDay && training.sequence > latestSequence) {
                        latestTraining = training;
                        latestSequence = training.sequence;
                    }
                }
                
                writeLog('DEBUG', '実施日' + implementationDay + 'の最新研修: ' + (latestTraining ? latestTraining.name + '(実施順' + latestTraining.sequence + ')' : 'なし'));
                return latestTraining;
            },
            
            /**
             * 指定した実施日・実施順の最新研修を取得（同じ実施順の研修がある場合用）
             * @param {number} implementationDay - 実施日
             * @param {number} sequence - 実施順
             * @returns {Object|null} 最新の研修情報
             */
            getLatestTrainingInSequence: function(implementationDay, sequence) {
                var latestTraining = null;
                
                // 同じ実施日・実施順の研修を全て取得
                var sameSequenceTrainings = [];
                for (var i = 0; i < scheduledTrainings.length; i++) {
                    var training = scheduledTrainings[i];
                    if (training.implementationDay === implementationDay && training.sequence === sequence) {
                        sameSequenceTrainings.push(training);
                    }
                }
                
                // 最後に追加された（配列の最後の）研修を取得
                if (sameSequenceTrainings.length > 0) {
                    latestTraining = sameSequenceTrainings[sameSequenceTrainings.length - 1];
                }
                
                writeLog('DEBUG', '実施日' + implementationDay + '実施順' + sequence + 'の最新研修: ' + 
                         (latestTraining ? latestTraining.name + '(同実施順研修数: ' + sameSequenceTrainings.length + ')' : 'なし'));
                return latestTraining;
            },
            
            /**
             * 現在のスケジュール状況をログ出力
             */
            logCurrentSchedule: function() {
                writeLog('DEBUG', '=== 現在の実施順スケジュール ===');
                if (scheduledTrainings.length === 0) {
                    writeLog('DEBUG', 'スケジュール済み研修なし');
                } else {
                    for (var i = 0; i < scheduledTrainings.length; i++) {
                        var training = scheduledTrainings[i];
                        writeLog('DEBUG', (i + 1) + '. ' + training.name + ' (実施日' + training.implementationDay + '実施順' + training.sequence + 
                                 ', ' + (training.startTime ? Utilities.formatDate(training.startTime, 'Asia/Tokyo', 'MM/dd HH:mm') : '時間未設定') + 
                                 '-' + (training.endTime ? Utilities.formatDate(training.endTime, 'Asia/Tokyo', 'HH:mm') : ''));
                    }
                }
                writeLog('DEBUG', '=== スケジュール状況終了 ===');
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

writeLog('INFO', 'CalendarUtils.gs - SequenceManager を読み込みました');

// =========================================
// 時間枠計算システム
// =========================================

/**
 * 時間枠計算を管理するクラス（シングルトン）
 */
var TimeSlotCalculator = (function() {
    var instance = null;
    
    function createInstance() {
        return {
            /**
             * 研修の適切な時間枠を計算
             * @param {Object} trainingGroup - 研修グループ
             * @param {Date} hireDate - 入社日
             * @returns {Object|null} {start: Date, end: Date} または null
             */
            calculateTimeSlot: function(trainingGroup, hireDate) {
                var durationMinutes = this.extractDurationMinutes(trainingGroup.time);
                var implementationDay = trainingGroup.implementationDay || 50;
                var sequence = trainingGroup.sequence || 100;

                writeLog('INFO', '時間枠計算開始: ' + trainingGroup.name + ' (実施日: ' + implementationDay + ', 実施順: ' + sequence + ', 時間: ' + durationMinutes + '分)');

                // 実施日を計算
                var targetDate = this.calculateImplementationDate(hireDate, implementationDay);
                if (!targetDate || targetDate.getDay() === 0 || targetDate.getDay() === 6) {
                    writeLog('WARN', '対象日が週末または無効です: ' + (targetDate ? Utilities.formatDate(targetDate, 'Asia/Tokyo', 'yyyy/MM/dd') : 'null'));
                    return null;
                }

                // 基準開始時間を設定
                var baseStartTime = this.getBaseStartTime(targetDate, implementationDay);
                
                // 実施順に基づく開始時間を計算
                var calculatedStartTime = this.calculateSequenceBasedStartTime(targetDate, implementationDay, sequence, baseStartTime);
                
                var proposedStart = new Date(calculatedStartTime.getTime());
                var proposedEnd = new Date(proposedStart.getTime() + (durationMinutes * 60 * 1000));

                // === 昼休み (12:00-13:00) との重複を回避 ===
                var lunchStart = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate(), 12, 0, 0);
                var lunchEnd   = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate(), 13, 0, 0);

                // 開始が昼休み前でも終了が昼休みを跨ぐ場合は 13:00 にシフト
                if (proposedStart < lunchEnd && proposedEnd > lunchStart) {
                    writeLog('DEBUG', '昼休み重複検知: ' + Utilities.formatDate(proposedStart, 'Asia/Tokyo', 'HH:mm') + '-' + Utilities.formatDate(proposedEnd, 'Asia/Tokyo', 'HH:mm') + ' → 13:00 にシフト');

                    // 新しい開始時刻を 13:00 に設定
                    proposedStart = new Date(lunchEnd.getTime());
                    proposedEnd   = new Date(proposedStart.getTime() + (durationMinutes * 60 * 1000));

                    // もし営業時間外になった場合は調整不可と判断
                    if (!this.isWithinBusinessHours(proposedEnd, implementationDay)) {
                        writeLog('WARN', '昼休み回避後の終了時刻が営業時間外です。時間枠計算を中止');
                        return null;
                    }
                }

                // 営業時間チェック
                if (!this.isWithinBusinessHours(proposedEnd, implementationDay)) {
                    writeLog('WARN', '計算された終了時刻が営業時間外です: ' + Utilities.formatDate(proposedEnd, 'Asia/Tokyo', 'HH:mm'));
                    return null;
                }

                writeLog('INFO', '時間枠計算完了: ' + Utilities.formatDate(proposedStart, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + Utilities.formatDate(proposedEnd, 'Asia/Tokyo', 'HH:mm'));
                return { start: proposedStart, end: proposedEnd };
            },
            
            /**
             * 研修時間文字列から分数を抽出
             * @param {string} timeString - 時間文字列
             * @returns {number} 分数
             */
            extractDurationMinutes: function(timeString) {
                var durationMinutes = 60; // デフォルト60分
                if (timeString && timeString.indexOf('分') !== -1) {
                    var timeMatch = timeString.match(/(\d+)分/);
                    if (timeMatch) {
                        durationMinutes = parseInt(timeMatch[1]);
                    }
                }
                return durationMinutes;
            },
            
            /**
             * 入社日から実施日を計算
             * @param {Date} hireDate - 入社日
             * @param {number} businessDays - 営業日数
             * @returns {Date} 計算された日付
             */
            calculateImplementationDate: function(hireDate, businessDays) {
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
            },
            
            /**
             * 基準開始時間を取得
             * @param {Date} targetDate - 対象日
             * @param {number} implementationDay - 実施日
             * @returns {Date} 基準開始時間
             */
            getBaseStartTime: function(targetDate, implementationDay) {
                var baseHour, baseMinute;
                if (implementationDay === 1) {
                    baseHour = 15; baseMinute = 0;
                } else if (implementationDay === 2) {
                    baseHour = 16; baseMinute = 0;
                } else {
                    baseHour = 10; baseMinute = 0; // 3営業日目以降は10:00開始に変更
                }
                
                return new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate(), baseHour, baseMinute);
            },
            
            /**
             * 実施順に基づく開始時間を計算（修正版）
             * @param {Date} targetDate - 対象日
             * @param {number} implementationDay - 実施日
             * @param {number} sequence - 実施順
             * @param {Date} baseStartTime - 基準開始時間
             * @returns {Date} 計算された開始時間
             */
            calculateSequenceBasedStartTime: function(targetDate, implementationDay, sequence, baseStartTime) {
                writeLog('DEBUG', '実施順計算開始: 実施日' + implementationDay + '実施順' + sequence);
                
                var sequenceManager = SequenceManager.getInstance();
                
                // 実施順1番の場合は基準時間から開始（ただし、同じ実施順の研修がある場合は調整）
                if (sequence === 1) {
                    // 同じ実施日・実施順の既存研修をチェック
                    var sameSequenceTraining = sequenceManager.getLatestTrainingInSequence(implementationDay, sequence);
                    if (sameSequenceTraining && sameSequenceTraining.endTime) {
                        var startTime = new Date(sameSequenceTraining.endTime.getTime());
                        startTime = this.adjustForLunchTime(targetDate, startTime);
                        writeLog('DEBUG', '実施順' + sequence + ': 同実施順の既存研修終了後から開始: ' + 
                                 Utilities.formatDate(startTime, 'Asia/Tokyo', 'HH:mm'));
                        return startTime;
                    } else {
                        writeLog('DEBUG', '実施順1番のため基準時間から開始: ' + Utilities.formatDate(baseStartTime, 'Asia/Tokyo', 'HH:mm'));
                        return baseStartTime;
                    }
                }

                // 直前の実施順の研修を探す
                var previousTraining = sequenceManager.getPreviousTraining(implementationDay, sequence);
                
                if (previousTraining && previousTraining.endTime) {
                    // 直前の研修の終了時刻から即座に開始
                    var startTime = new Date(previousTraining.endTime.getTime());
                    
                    // 昼休み時間帯（12:00-13:00）を跨ぐ場合の調整
                    startTime = this.adjustForLunchTime(targetDate, startTime);
                    
                    writeLog('DEBUG', '実施順' + sequence + ': 直前研修（実施順' + (sequence - 1) + '）終了後から開始: ' + 
                             Utilities.formatDate(startTime, 'Asia/Tokyo', 'HH:mm'));
                    return startTime;
                } else {
                    // 直前の実施順がない場合、同じ実施順の既存研修をチェック
                    var sameSequenceTraining = sequenceManager.getLatestTrainingInSequence(implementationDay, sequence);
                    if (sameSequenceTraining && sameSequenceTraining.endTime) {
                        var startTime = new Date(sameSequenceTraining.endTime.getTime());
                        startTime = this.adjustForLunchTime(targetDate, startTime);
                        writeLog('DEBUG', '実施順' + sequence + ': 同実施順の既存研修終了後から開始: ' + 
                                 Utilities.formatDate(startTime, 'Asia/Tokyo', 'HH:mm'));
                        return startTime;
                    }
                    
                    // 同実施日の最新研修の後から開始
                    var latestTraining = sequenceManager.getLatestTrainingInDay(implementationDay);
                    
                    if (latestTraining && latestTraining.endTime) {
                        var startTime = new Date(latestTraining.endTime.getTime());
                        startTime = this.adjustForLunchTime(targetDate, startTime);
                        
                        writeLog('DEBUG', '実施順' + sequence + ': 実施日' + implementationDay + 'の最新研修終了後から開始: ' + 
                                 Utilities.formatDate(startTime, 'Asia/Tokyo', 'HH:mm'));
                        return startTime;
                    } else {
                        // 既存研修がない場合は基準時間から開始
                        writeLog('DEBUG', '実施順' + sequence + ': 既存研修なし、基準時間から開始: ' + 
                                 Utilities.formatDate(baseStartTime, 'Asia/Tokyo', 'HH:mm'));
                        return baseStartTime;
                    }
                }
            },
            
            /**
             * 昼休み時間の調整
             * @param {Date} targetDate - 対象日
             * @param {Date} startTime - 開始時間
             * @returns {Date} 調整後の開始時間
             */
            adjustForLunchTime: function(targetDate, startTime) {
                var lunchStart = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate(), 12, 0);
                var lunchEnd = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate(), 13, 0);
                
                if (startTime >= lunchStart && startTime < lunchEnd) {
                    writeLog('DEBUG', '昼休み時間調整: ' + Utilities.formatDate(startTime, 'Asia/Tokyo', 'HH:mm') + ' → 13:00');
                    return lunchEnd;
                }
                return startTime;
            },
            
            /**
             * 営業時間内チェック
             * @param {Date} endTime - 終了時間
             * @param {number} [implementationDay] - 実施日（1,2営業日の場合は上限を24時に拡張）
             * @returns {boolean} 営業時間内かどうか
             */
            isWithinBusinessHours: function(endTime, implementationDay) {
                // 1・2営業日は 24:00 まで許容、それ以降は 19:00
                var maxEndHour = (implementationDay === 1 || implementationDay === 2) ? 24 : 19;
                // 24時上限の場合は常に true（23:59 以前）
                if (maxEndHour === 24) return true;
                return endTime.getHours() < maxEndHour || (endTime.getHours() === maxEndHour && endTime.getMinutes() === 0);
            },
            
            /**
             * 翌営業日を計算
             * @param {Date} date - 基準日
             * @returns {Date} 翌営業日
             */
            calculateNextWorkDay: function(date) {
                var nextDay = new Date(date);
                nextDay.setDate(nextDay.getDate() + 1);
                
                // 土日の場合は次の月曜日に調整
                if (nextDay.getDay() === 0) { // 日曜日
                    nextDay.setDate(nextDay.getDate() + 1);
                } else if (nextDay.getDay() === 6) { // 土曜日
                    nextDay.setDate(nextDay.getDate() + 2);
                }
                
                writeLog('DEBUG', '翌営業日を計算: ' + Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd') + 
                         ' → ' + Utilities.formatDate(nextDay, 'Asia/Tokyo', 'yyyy/MM/dd'));
                         
                return nextDay;
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

writeLog('INFO', 'CalendarUtils.gs - TimeSlotCalculator を読み込みました');

// =========================================
// 会議室管理システム（改良版）
// =========================================

/**
 * 会議室管理クラス（改良版）
 */
var RoomManager = (function() {
    var instance = null;
    
    function createInstance() {
        var reservations = [];
        
        return {
            /**
             * 予約をリセット
             */
            reset: function() {
                reservations = [];
                writeLog('INFO', '会議室管理をリセットしました');
            },
            
            /**
             * 会議室を予約
             * @param {Object} roomInfo - 会議室情報
             * @param {Date} startTime - 開始時間
             * @param {Date} endTime - 終了時間
             * @param {string} trainingName - 研修名
             * @returns {boolean} 予約成功かどうか
             */
            reserveRoom: function(roomInfo, startTime, endTime, trainingName) {
                if (!this.isRoomAvailable(roomInfo.name, startTime, endTime)) {
                    writeLog('WARN', '会議室予約失敗（時間重複）: ' + roomInfo.name);
                    return false;
                }
                
                var reservation = {
                    roomName: roomInfo.name,
                    calendarId: roomInfo.calendarId,
                    startTime: new Date(startTime),
                    endTime: new Date(endTime),
                    trainingName: trainingName,
                    reservedAt: new Date()
                };
                
                reservations.push(reservation);
                writeLog('INFO', '会議室予約成功: ' + roomInfo.name + ' (' + 
                         Utilities.formatDate(startTime, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + 
                         Utilities.formatDate(endTime, 'Asia/Tokyo', 'HH:mm') + ')');
                return true;
            },
            
            /**
             * 会議室の利用可能性をチェック
             * @param {string} roomName - 会議室名
             * @param {Date} startTime - 開始時間
             * @param {Date} endTime - 終了時間
             * @returns {boolean} 利用可能かどうか
             */
            isRoomAvailable: function(roomName, startTime, endTime) {
                for (var i = 0; i < reservations.length; i++) {
                    var reservation = reservations[i];
                    if (reservation.roomName === roomName) {
                        if (!(endTime <= reservation.startTime || startTime >= reservation.endTime)) {
                            return false;
                        }
                    }
                }
                return true;
            },
            
            /**
             * 適切な会議室を検索・予約
             * @param {number} numberOfAttendees - 参加者数
             * @param {Date} startTime - 開始時間
             * @param {Date} endTime - 終了時間
             * @param {string} trainingName - 研修名
             * @returns {Object} 会議室予約結果
             */
            findAndReserveRoom: function(numberOfAttendees, startTime, endTime, trainingName) {
                writeLog('DEBUG', '会議室検索・予約開始: 必要人数=' + numberOfAttendees);
                
                var availableRooms = this.getAvailableRooms(startTime, endTime);
                if (availableRooms.length === 0) {
                    throw new Error('指定時間に利用可能な会議室が見つかりませんでした');
                }

                // 参加者全員を収容できる会議室を検索
                var suitableRooms = availableRooms.filter(function(room) {
                    return room.capacity >= numberOfAttendees;
                });
                
                var selectedRoom;
                var isHybrid = false;
                var onlineCount = 0;

                if (suitableRooms.length > 0) {
                    // 最適な会議室を選択（定員の小さい順）
                    suitableRooms.sort(function(a, b) { return a.capacity - b.capacity; });
                    selectedRoom = suitableRooms[0];
                    writeLog('INFO', '最適会議室選択: ' + selectedRoom.name + ' (定員: ' + selectedRoom.capacity + ')');
                } else {
                    // ハイブリッド開催用に最大の会議室を選択
                    availableRooms.sort(function(a, b) { return b.capacity - a.capacity; });
                    selectedRoom = availableRooms[0];
                    isHybrid = true;
                    onlineCount = numberOfAttendees - selectedRoom.capacity;
                    writeLog('INFO', 'ハイブリッド開催会議室選択: ' + selectedRoom.name + ' (オンライン参加: ' + onlineCount + '名)');
                }
                
                // 会議室を予約
                if (!this.reserveRoom(selectedRoom, startTime, endTime, trainingName)) {
                    throw new Error('会議室の予約に失敗しました: ' + selectedRoom.name);
                }
                
                return {
                    roomName: selectedRoom.name,
                    isHybrid: isHybrid,
                    capacity: selectedRoom.capacity,
                    onlineCount: onlineCount,
                    calendarId: selectedRoom.calendarId
                };
            },
            
            /**
             * 利用可能な会議室を取得
             * @param {Date} startTime - 開始時間
             * @param {Date} endTime - 終了時間
             * @returns {Array} 利用可能な会議室リスト
             */
            getAvailableRooms: function(startTime, endTime) {
                try {
                    var sheet = SpreadsheetApp.openById(SPREADSHEET_IDS.ROOM_MASTER).getSheetByName(SHEET_NAMES.ROOM_MASTER);
                    var lastRow = sheet.getLastRow();
                    
                    if (lastRow <= 1) {
                        throw new Error('会議室マスタにデータがありません');
                    }
                    
                    var data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
                    var availableRooms = [];
                    
                    for (var i = 0; i < data.length; i++) {
                        var row = data[i];
                        var roomName = row[0];
                        var calendarId = row[1];
                        var capacity = row[2];
                        var onbAvailable = row[3];
                        
                        if (onbAvailable !== '利用可能') continue;
                        
                        // 内部予約システムでの重複チェック
                        if (!this.isRoomAvailable(roomName, startTime, endTime)) {
                            continue;
                        }
                        
                        // Googleカレンダーでの重複チェック
                        if (calendarId && this.isGoogleCalendarRoomAvailable(calendarId, startTime, endTime)) {
                            availableRooms.push({
                                name: roomName,
                                capacity: capacity,
                                calendarId: calendarId
                            });
                        }
                    }
                    
                    return availableRooms;
                } catch (e) {
                    writeLog('ERROR', '利用可能会議室取得でエラー: ' + e.message);
                    throw e;
                }
            },
            
            /**
             * Googleカレンダーでの会議室可用性チェック
             * @param {string} calendarId - カレンダーID
             * @param {Date} startTime - 開始時間
             * @param {Date} endTime - 終了時間
             * @returns {boolean} 利用可能かどうか
             */
            isGoogleCalendarRoomAvailable: function(calendarId, startTime, endTime) {
                try {
                    var resourceCalendar = CalendarApp.getCalendarById(calendarId);
                    if (!resourceCalendar) {
                        writeLog('WARN', 'リソースカレンダーにアクセスできません: ' + calendarId);
                        return false;
                    }
                    
                    var existingEvents = resourceCalendar.getEvents(startTime, endTime);
                    if (existingEvents.length === 0) {
                        return true;
                    }
                    
                    // 実際の時間重複チェック
                    for (var i = 0; i < existingEvents.length; i++) {
                        var event = existingEvents[i];
                        
                        // 終日イベントをスキップ
                        if (event.isAllDayEvent()) {
                            continue;
                        }
                        
                        var eventStart = event.getStartTime();
                        var eventEnd = event.getEndTime();
                        
                        // 時間重複チェック
                        if (!(endTime <= eventStart || startTime >= eventEnd)) {
                            return false;
                        }
                    }
                    
                    return true;
                } catch (e) {
                    writeLog('WARN', 'Googleカレンダー可用性チェックでエラー: ' + e.message);
                    return false;
                }
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

writeLog('INFO', 'CalendarUtils.gs - RoomManager を読み込みました');

// =========================================
// カレンダーイベント管理システム
// =========================================

/**
 * カレンダーイベント管理クラス（シングルトン）
 */
var CalendarEventManager = (function() {
    var instance = null;
    
    function createInstance() {
        return {
            /**
             * 全研修のカレンダーイベントを作成
             * @param {Array<Object>} trainingGroups - 研修グループ配列
             * @param {Date} hireDate - 入社日
             * @returns {Array} 処理結果配列
             */
            createAllCalendarEvents: function(trainingGroups, hireDate) {
                writeLog('INFO', '全研修カレンダーイベント作成開始: ' + trainingGroups.length + '件');
                
                // 管理システムをリセット
                var sequenceManager = SequenceManager.getInstance();
                var roomManager = RoomManager.getInstance();
                var calendarManager = CalendarEventManager.getInstance();
                
                sequenceManager.reset();
                roomManager.reset();
                
                // 研修グループを実施日・実施順でソート
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
                
                writeLog('INFO', '研修グループをソートしました（実施日・実施順）');
                
                var results = [];
                
                for (var i = 0; i < trainingGroups.length; i++) {
                    var group = trainingGroups[i];
                    
                    try {
                        writeLog('INFO', '研修処理開始 (' + (i + 1) + '/' + trainingGroups.length + '): ' + group.name + 
                                 ' (実施日: ' + (group.implementationDay || 'なし') + ', 実施順: ' + (group.sequence || 'なし') + ')');
                        
                        var result = this.processSingleTraining(group, hireDate);
                        results.push(result);
                        
                        // 成功した場合はスケジュール管理に追加
                        if (result.scheduled) {
                            group.startTime = result.eventTime.start;
                            group.endTime = result.eventTime.end;
                            sequenceManager.addTraining(group);
                        }
                        
                    } catch (e) {
                        writeLog('ERROR', '研修処理失敗: ' + group.name + ' - ' + e.message);
                        results.push({
                            training: group,
                            scheduled: false,
                            roomName: group.needsRoom ? '会議室未確保' : 'オンライン',
                            eventTime: null,
                            error: e.message
                        });
                    }
                }
                
                writeLog('INFO', '全研修カレンダーイベント作成完了');
                sequenceManager.logCurrentSchedule();
                
                return results;
            },
            
            /**
             * 単一研修を処理
             * @param {Object} trainingGroup - 研修グループ
             * @param {Date} hireDate - 入社日
             * @returns {Object} 処理結果
             */
            processSingleTraining: function(trainingGroup, hireDate) {
                // 参加者数チェック
                var participantCount = this.getParticipantCount(trainingGroup);
                if (participantCount === 0) {
                    writeLog('INFO', '参加者0名のため研修をスキップ: ' + trainingGroup.name);
                    return {
                        training: trainingGroup,
                        scheduled: false,
                        roomName: '実施不要',
                        eventTime: null,
                        error: '参加者0名のため実施不要'
                    };
                }
                
                // 時間枠を計算
                var timeSlotCalculator = TimeSlotCalculator.getInstance();
                var eventTime = timeSlotCalculator.calculateTimeSlot(trainingGroup, hireDate);
                
                if (!eventTime) {
                    throw new Error('適切な時間枠が見つかりませんでした');
                }
                
                // 参加者の時間重複チェック（フォールバック検索機能付き）
                var originalStartTime = eventTime.start;
                var originalEndTime = eventTime.end;
                
                // 最大5回まで代替時間を探す
                var shiftMinutes = 30;
                var scanCounter = 0; // 安全装置（無限ループ防止）
                var maxScan = 48;    // 当日内で最大 24 時間 / 30 分 = 48 回
                var durationMinutes = (eventTime.end.getTime() - eventTime.start.getTime()) / (60 * 1000);
                
                while (true) {
                    // 現在の時間枠で利用可能かチェック
                    if (this.isTimeSlotAvailable(eventTime.start, eventTime.end, trainingGroup)) {
                        if (scanCounter > 0) {
                            writeLog('INFO', '代替時間枠が見つかりました（シフト回数: ' + scanCounter + '）: ' + 
                                     Utilities.formatDate(eventTime.start, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + 
                                     Utilities.formatDate(eventTime.end, 'Asia/Tokyo', 'HH:mm'));
                        }
                        break;
                    }
                    
                    // 使用できない場合は 30 分後を試す
                    scanCounter++;
                    var newStartTime = new Date(eventTime.start.getTime() + (shiftMinutes * 60 * 1000));
                    var newEndTime = new Date(newStartTime.getTime() + (durationMinutes * 60 * 1000));
                    
                    // 当日 24:00 を超えたか（日付が変わったか）をチェック
                    if (newStartTime.getDate() !== eventTime.start.getDate()) {
                        writeLog('INFO', '当日内に空きなし → 翌営業日へ移行');
                        var nextWorkDay = timeSlotCalculator.calculateNextWorkDay(eventTime.start);
                        var baseStartTime = timeSlotCalculator.getBaseStartTime(nextWorkDay, trainingGroup.implementationDay);
                        newStartTime = baseStartTime;
                        newEndTime = new Date(newStartTime.getTime() + (durationMinutes * 60 * 1000));
                        scanCounter = 0; // カウンタリセット
                    }
                    
                    // 新しい時間枠を設定
                    eventTime.start = newStartTime;
                    eventTime.end = newEndTime;
                    
                    writeLog('INFO', '代替時間枠を試行中（シフト回数: ' + scanCounter + '）: ' + 
                            Utilities.formatDate(eventTime.start, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + 
                            Utilities.formatDate(eventTime.end, 'Asia/Tokyo', 'HH:mm'));

                    if (scanCounter > maxScan) {
                        writeLog('WARN', '安全装置: 当日のスキャン回数が上限を超えました');
                    }
                }
                
                // すべての試行で失敗した場合（翌営業日以降でも見つからず）
                if (!this.isTimeSlotAvailable(eventTime.start, eventTime.end, trainingGroup)) {
                    writeLog('ERROR', '適切な代替時間枠が見つかりませんでした。元の時間枠に戻します: ' + 
                            Utilities.formatDate(originalStartTime, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + 
                            Utilities.formatDate(originalEndTime, 'Asia/Tokyo', 'HH:mm'));
                    
                    eventTime.start = originalStartTime;
                    eventTime.end = originalEndTime;
                    
                    throw new Error('参加者または講師の時間が重複しています（代替時間枠も見つかりませんでした）');
                }
                
                // 会議室確保
                var roomReservation = null;
                if (trainingGroup.needsRoom) {
                    var roomManager = RoomManager.getInstance();
                    var lecturerCount = trainingGroup.lecturerEmails ? trainingGroup.lecturerEmails.length : 1;
                    var totalAttendees = participantCount + lecturerCount;
                    
                    roomReservation = roomManager.findAndReserveRoom(totalAttendees, eventTime.start, eventTime.end, trainingGroup.name);
                }
                
                // カレンダーイベント作成
                this.createSingleCalendarEvent(trainingGroup, roomReservation, eventTime.start, eventTime.end);
                
                return {
                    training: trainingGroup,
                    scheduled: true,
                    roomName: roomReservation ? roomReservation.roomName : 'オンライン',
                    eventTime: eventTime,
                    error: null,
                    calendarEventId: trainingGroup.calendarEventId
                };
            },
            
            /**
             * 参加者数を取得（講師を除く）
             * @param {Object} trainingGroup - 研修グループ
             * @returns {number} 参加者数
             */
            getParticipantCount: function(trainingGroup) {
                var count = 0;
                if (trainingGroup.attendees && trainingGroup.attendees.length > 0) {
                    for (var i = 0; i < trainingGroup.attendees.length; i++) {
                        if (trainingGroup.attendees[i] !== trainingGroup.lecturer) {
                            count++;
                        }
                    }
                }
                return count;
            },
            
            /**
             * 時間枠の利用可能性をチェック
             * @param {Date} startTime - 開始時間
             * @param {Date} endTime - 終了時間
             * @param {Object} trainingGroup - 研修グループ
             * @returns {boolean} 利用可能かどうか
             */
            isTimeSlotAvailable: function(startTime, endTime, trainingGroup) {
                var sequenceManager = SequenceManager.getInstance();
                
                // 既存スケジュールとの重複チェック（簡素化）
                // この部分は既存の複雑なロジックを簡略化
                
                // 講師の空き時間チェック
                var lecturerEmails = trainingGroup.lecturerEmails || (trainingGroup.lecturer ? [trainingGroup.lecturer] : []);
                
                // 各講師の予定をチェック
                for (var i = 0; i < lecturerEmails.length; i++) {
                    var lecturerEmail = lecturerEmails[i];
                    writeLog('DEBUG', '講師カレンダーチェック: ' + lecturerEmail);
                    
                    if (!this.isLecturerAvailable(lecturerEmail, startTime, endTime)) {
                        writeLog('INFO', '講師 ' + lecturerEmail + ' の時間が重複しています: ' + 
                                 Utilities.formatDate(startTime, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + 
                                 Utilities.formatDate(endTime, 'Asia/Tokyo', 'HH:mm'));
                        
                        return false;
                    }
                }
                
                return true;
            },
            
            /**
             * 講師の空き時間チェック
             * @param {string} lecturerEmail - 講師メールアドレス
             * @param {Date} startTime - 開始時間
             * @param {Date} endTime - 終了時間
             * @returns {boolean} 利用可能かどうか
             */
            isLecturerAvailable: function(lecturerEmail, startTime, endTime) {
                try {
                    var lecturerCalendar = CalendarApp.getCalendarById(lecturerEmail);
                    if (!lecturerCalendar) {
                        writeLog('WARN', '講師カレンダーにアクセスできません: ' + lecturerEmail + '（空き時間とみなします）');
                        return true; // アクセスできない場合は空いているとみなす
                    }
                    
                    var existingEvents = lecturerCalendar.getEvents(startTime, endTime);
                    for (var i = 0; i < existingEvents.length; i++) {
                        var event = existingEvents[i];
                        
                        if (event.isAllDayEvent()) {
                            continue;
                        }
                        
                        var eventStart = event.getStartTime();
                        var eventEnd = event.getEndTime();
                        
                        // 時間帯が重なるか判定
                        var isOverlap = !(endTime <= eventStart || startTime >= eventEnd);
                        if (isOverlap) {
                            var title = event.getTitle();
                            var visibility = (typeof event.getVisibility === 'function') ? event.getVisibility() : '';
                            var isPrivate = (visibility === CalendarApp.Visibility.PRIVATE || visibility === 'private');
                            // タイトルに「作業」または「社内」を含み、かつ非公開でない場合は重複を許容
                            if (/作業|社内/.test(title) && !isPrivate) {
                                writeLog('INFO', '重複だが許容タイトル（公開）: ' + title);
                                continue; // 無視して空きとみなす
                            }

                            // それ以外は重複NG
                            writeLog('DEBUG', '講師の予定と重複（不可）: ' + title + ' (' + 
                                    Utilities.formatDate(eventStart, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + 
                                    Utilities.formatDate(eventEnd, 'Asia/Tokyo', 'HH:mm') + ')');
                            return false;
                        }
                    }
                    
                    return true;
                } catch (e) {
                    writeLog('WARN', '講師カレンダーチェックでエラー: ' + e.message);
                    return true; // エラー時は空いているとみなす
                }
            },
            
            /**
             * 単一カレンダーイベントを作成
             * @param {Object} trainingDetails - 研修詳細
             * @param {Object} roomReservation - 会議室予約情報
             * @param {Date} startTime - 開始時間
             * @param {Date} endTime - 終了時間
             */
            createSingleCalendarEvent: function(trainingDetails, roomReservation, startTime, endTime) {
                var title = trainingDetails.name;
                var roomName = roomReservation ? roomReservation.roomName : 'オンライン';
                
                writeLog('INFO', 'カレンダーイベント作成: ' + title + ' (' + 
                         Utilities.formatDate(startTime, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + 
                         Utilities.formatDate(endTime, 'Asia/Tokyo', 'HH:mm') + ')');
                
                // 有効な参加者メールアドレスをフィルタリング
                var validEmails = [];
                for (var i = 0; i < trainingDetails.attendees.length; i++) {
                    var email = trainingDetails.attendees[i];
                    if (email && email.trim() !== '' && email.indexOf('@') > 0) {
                        validEmails.push(email.trim());
                    }
                }
                
                // 会議室リソースを参加者に追加
                if (roomReservation && roomReservation.calendarId) {
                    validEmails.push(roomReservation.calendarId);
                }
                
                var options = {
                    description: trainingDetails.memo || '',
                    guests: validEmails.join(','),
                    location: roomName === 'オンライン' ? '' : roomName
                };

                // ハイブリッド開催またはオンラインの場合はGoogle Meetを追加
                if (!roomReservation || roomReservation.isHybrid || roomName === 'オンライン') {
                    options.conferenceDataVersion = 1;
                }
                
                try {
                    var calendarEvent = CalendarApp.createEvent(title, startTime, endTime, options);
                    trainingDetails.calendarEventId = calendarEvent.getId();
                    
                    writeLog('INFO', 'カレンダーイベント作成成功: ' + title + ' (ID: ' + trainingDetails.calendarEventId + ')');
                } catch (e) {
                    writeLog('ERROR', 'カレンダーイベント作成失敗: ' + title + ' - ' + e.message);
                    throw e;
                }
            },
            
            /**
             * カレンダーイベントを削除
             * @param {string} eventId - イベントID
             * @returns {boolean} 削除成功かどうか
             */
            deleteSingleCalendarEvent: function(eventId) {
                // IDの形式チェックと前処理
                try {
                    if (!eventId) {
                        writeLog('WARN', 'イベントIDが空です');
                        return false;
                    }
                    
                    // 基本的な形式検証
                    if (typeof eventId !== 'string') {
                        writeLog('WARN', '無効なイベントID（文字列でない）: ' + eventId);
                        return false;
                    }
                    
                    // イベントIDの形式確認（フォーマット：<calendarId>_<eventUnique>）
                    if (!eventId.includes('@') || !eventId.includes('_')) {
                        writeLog('WARN', 'イベントIDのフォーマットが正しくありません: ' + eventId);
                        return false;
                    }
                    
                    var event = CalendarApp.getEventById(eventId);
                    if (!event) {
                        writeLog('WARN', 'イベントが見つかりません: ' + eventId);
                        return false;
                    }
                    
                    var eventTitle = event.getTitle();
                    event.deleteEvent();
                    
                    writeLog('INFO', 'カレンダーイベント削除成功: ' + eventTitle);
                    return true;
                } catch (e) {
                    writeLog('ERROR', 'カレンダーイベント削除失敗: ' + eventId + ' - ' + e.message);
                    return false;
                }
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

writeLog('INFO', 'CalendarUtils.gs - CalendarEventManager を読み込みました');

// =========================================
// インクリメンタル更新処理
// =========================================

/**
 * 研修グループをインクリメンタルに処理してマッピングシートを更新する
 * @param {Array<Object>} trainingGroups - 研修グループの配列
 * @param {Array<Object>} allNewHires - 全入社者の配列
 * @param {Date} hireDate - 入社日
 * @param {Object} mappingSheet - 更新対象のマッピングシート
 */
function processTrainingGroupsIncrementally(trainingGroups, allNewHires, hireDate, mappingSheet) {
    writeLog('INFO', 'インクリメンタル処理開始: ' + trainingGroups.length + '件の研修を処理');
    
    // 管理システムをリセット
    var sequenceManager = SequenceManager.getInstance();
    var roomManager = RoomManager.getInstance();
    var calendarManager = CalendarEventManager.getInstance();
    
    sequenceManager.reset();
    roomManager.reset();
    
    // 研修グループを実施日・実施順でソート
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
    
    writeLog('INFO', '研修グループをソートしました（実施日・実施順）');
    
    var successCount = 0;
    var errorCount = 0;
    
    for (var i = 0; i < trainingGroups.length; i++) {
        var group = trainingGroups[i];
        var rowIndex = i + 2; // スプレッドシートの行番号（ヘッダー行の次から）
        
        writeLog('INFO', '処理中 (' + (i + 1) + '/' + trainingGroups.length + '): ' + group.name + 
                 ' (実施日: ' + (group.implementationDay || 'なし') + ', 実施順: ' + (group.sequence || 'なし') + ')');
        
        try {
            // 処理状況を「処理中」に更新
            updateMappingSheetRow(mappingSheet, rowIndex, {
                status: '処理中...',
                roomName: '確保中...',
                schedule: '計算中...'
            });
            
            var result = calendarManager.processSingleTraining(group, hireDate);
            
            if (result.scheduled) {
                // 成功時の更新
                var scheduleStr = Utilities.formatDate(result.eventTime.start, 'Asia/Tokyo', 'MM/dd(E) HH:mm') + 
                                '-' + Utilities.formatDate(result.eventTime.end, 'Asia/Tokyo', 'HH:mm');
                
                updateMappingSheetRow(mappingSheet, rowIndex, {
                    status: '成功',
                    roomName: result.roomName,
                    schedule: scheduleStr,
                    calendarId: group.calendarEventId || ''
                });
                
                // 成功した研修をSequenceManagerに追加
                group.startTime = result.eventTime.start;
                group.endTime = result.eventTime.end;
                sequenceManager.addTraining(group);
                
                successCount++;
                writeLog('INFO', '研修処理成功: ' + group.name + ' (ID: ' + group.calendarEventId + ')');
            } else {
                updateMappingSheetRow(mappingSheet, rowIndex, {
                    status: result.error ? '失敗' : 'スキップ',
                    roomName: result.roomName,
                    schedule: result.error ? 'エラー' : 'N/A',
                    errorReason: result.error || ''
                });
                
                if (result.error && result.error !== '参加者0名のため実施不要') {
                    errorCount++;
                }
            }
        } catch (e) {
            writeLog('ERROR', '研修処理失敗: ' + group.name + ' - ' + e.message);
            updateMappingSheetRow(mappingSheet, rowIndex, {
                status: '失敗',
                roomName: group.needsRoom ? '会議室未確保' : 'オンライン',
                schedule: 'エラー',
                errorReason: e.message
            });
            errorCount++;
        }
        
        // 少し待機（API制限対策）
        if (i < trainingGroups.length - 1) {
            Utilities.sleep(1000); // 1秒待機
        }
    }
    
    // 最終的なスケジュール状況をログ出力
    sequenceManager.logCurrentSchedule();
    
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
            mappingSheet.getRange(rowIndex, MAPPING_COLS.ROOM_NAME).setValue(updates.roomName); // F列: 会議室名
        }
        if (updates.schedule !== undefined) {
            mappingSheet.getRange(rowIndex, MAPPING_COLS.SCHEDULE).setValue(updates.schedule); // G列: 研修実施日時
        }
        if (updates.calendarId !== undefined) {
            mappingSheet.getRange(rowIndex, MAPPING_COLS.CAL_ID).setValue(updates.calendarId); // H列: カレンダーID
        }
        if (updates.status !== undefined) {
            var statusCell = mappingSheet.getRange(rowIndex, MAPPING_COLS.STATUS);
            var statusVal = updates.status;
            if (statusVal && statusVal.indexOf('失敗:') === 0) {
                statusVal = '失敗';
            }
            statusCell.setValue(statusVal);
            
            // 処理状況に応じた背景色設定
            if (statusVal === '成功') {
                statusCell.setBackground('#d4edda').setFontColor('#155724');
            } else if (statusVal.indexOf('失敗') !== -1 || statusVal.indexOf('エラー') !== -1) {
                statusCell.setBackground('#f8d7da').setFontColor('#721c24');
            } else if (statusVal.indexOf('処理中') !== -1) {
                statusCell.setBackground('#fff3cd').setFontColor('#856404');
            } else if (statusVal.indexOf('スキップ') !== -1) {
                statusCell.setBackground('#e2e3e5').setFontColor('#383d41');
            }
        }
        if (updates.errorReason !== undefined) {
            var errorCell = mappingSheet.getRange(rowIndex, MAPPING_COLS.ERROR);
            errorCell.setValue(updates.errorReason);
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
    
    // === 成功研修のマップを生成 ===
    var successMap = {};
    var lastDataRow = mappingSheet.getLastRow();
    if (lastDataRow > 1) {
        var statusColCount = MAPPING_COLS.STATUS;
        var rawData = mappingSheet.getRange(2, 1, lastDataRow - 1, statusColCount).getValues();
        for (var ri = 0; ri < rawData.length; ri++) {
            var tName = rawData[ri][MAPPING_COLS.NAME - 1];
            var stat  = rawData[ri][MAPPING_COLS.STATUS - 1];
            if (tName) {
                successMap[tName] = (stat === '成功');
            }
        }
    }

    // --- 職位別研修数詳細 ---
    var rankSummary = {}; // { rank: { hireCount, required, registered } }
    var perHireRows = [];  // 個人別レポート保持

    for (var hi = 0; hi < newHires.length; hi++) {
        var hire = newHires[hi];
        var rank = hire.rank || '未分類';
        if (!rankSummary[rank]) {
            rankSummary[rank] = { hireCount: 0, required: 0, registered: 0};
        }
        rankSummary[rank].hireCount++;

        var requiredCnt = 0;
        var registeredCnt = 0;
        for (var tg = 0; tg < trainingGroups.length; tg++) {
            var grp = trainingGroups[tg];
            if (grp.attendees && grp.attendees.indexOf(hire.email) !== -1) {
                requiredCnt++;
                if (successMap[grp.name]) {
                    registeredCnt++;
                }
            }
        }
        rankSummary[rank].required += requiredCnt;
        rankSummary[rank].registered += registeredCnt;

        // 不足研修名リスト
        var missingNames = [];
        for (var tg2 = 0; tg2 < trainingGroups.length; tg2++) {
            var grp2 = trainingGroups[tg2];
            if (grp2.attendees && grp2.attendees.indexOf(hire.email) !== -1) {
                if (!successMap[grp2.name]) missingNames.push(grp2.name);
            }
        }

        perHireRows.push([
            hire.name,
            rank,
            requiredCnt,
            registeredCnt,
            Math.max(0, requiredCnt - registeredCnt),
            missingNames.join('\n')
        ]);

        // 職位別 missing list も蓄積
        if (!rankSummary[rank].missing) rankSummary[rank].missing = [];
        rankSummary[rank].missing = rankSummary[rank].missing.concat(missingNames.filter(function(n){return rankSummary[rank].missing.indexOf(n)==-1;}));
    }

    // --- 対象者別研修数検証 ---
    mappingSheet.getRange(currentRow,1).setValue('【対象者別研修数検証】');
    mappingSheet.getRange(currentRow,1).setFontWeight('bold');
    currentRow++;
    mappingSheet.getRange(currentRow,1,1,6).setValues([[ '氏名','職位','必要研修数','登録研修数','不足数','不足研修一覧' ]]);
    var headerHireRange = mappingSheet.getRange(currentRow,1,1,6);
    headerHireRange.setFontWeight('bold').setBackground('#4a90e2').setFontColor('white');
    currentRow++;
    if (perHireRows.length > 0) {
        var dataRangeHire = mappingSheet.getRange(currentRow,1,perHireRows.length,6);
        dataRangeHire.setValues(perHireRows);
        dataRangeHire.setBorder(true,true,true,true,true,true);
        dataRangeHire.setWrap(true).setVerticalAlignment('top');
        currentRow += perHireRows.length + 2;
    }

    // --- カレンダー重複チェックテーブル ---
    mappingSheet.getRange(currentRow,1).setValue('【カレンダー重複チェック】');
    mappingSheet.getRange(currentRow,1).setFontWeight('bold');
    currentRow++;

    var dupHeader = ['研修名①','研修名②','重複開始','重複終了'];
    mappingSheet.getRange(currentRow,1,1,dupHeader.length).setValues([dupHeader]);
    var dupHeaderRange = mappingSheet.getRange(currentRow,1,1,dupHeader.length);
    dupHeaderRange.setFontWeight('bold').setBackground('#4a90e2').setFontColor('white');
    currentRow++;

    var overlapRows = [];
    // 成功した研修で開始・終了時刻があるものを抽出
    var successEvents = [];
    for (var si=0; si<trainingGroups.length; si++) {
        var sg = trainingGroups[si];
        if (sg.startTime && sg.endTime && successMap[sg.name]) {
            successEvents.push(sg);
        }
    }
    for (var i1=0; i1<successEvents.length; i1++) {
        for (var i2=i1+1; i2<successEvents.length; i2++) {
            var e1 = successEvents[i1];
            var e2 = successEvents[i2];
            // 同日かつ時間帯重複判定
            var sameDate = Utilities.formatDate(e1.startTime,'Asia/Tokyo','yyyyMMdd') === Utilities.formatDate(e2.startTime,'Asia/Tokyo','yyyyMMdd');
            var overlap = (e1.startTime < e2.endTime) && (e2.startTime < e1.endTime);
            if (sameDate && overlap) {
                overlapRows.push([
                    e1.name,
                    e2.name,
                    Utilities.formatDate(new Date(Math.max(e1.startTime,e2.startTime)),'Asia/Tokyo','MM/dd HH:mm'),
                    Utilities.formatDate(new Date(Math.min(e1.endTime,e2.endTime)),'Asia/Tokyo','HH:mm')
                ]);
            }
        }
    }

    if (overlapRows.length > 0) {
        var dupRange = mappingSheet.getRange(currentRow,1,overlapRows.length,dupHeader.length);
        dupRange.setValues(overlapRows);
        dupRange.setBorder(true,true,true,true,true,true);
        currentRow += overlapRows.length;
    } else {
        mappingSheet.getRange(currentRow,1).setValue('重複なし');
        currentRow += 1;
    }

    // --- 研修マッピング検証（実施日・実施順） ---
    mappingSheet.getRange(currentRow,1).setValue('【研修マッピング検証】');
    mappingSheet.getRange(currentRow,1).setFontWeight('bold');
    currentRow++;

    var mapHeader = ['研修名','実施日(生成)','実施日(マスタ)','実施順(生成)','実施順(マスタ)'];
    mappingSheet.getRange(currentRow,1,1,mapHeader.length).setValues([mapHeader]);
    var mapHeaderRange = mappingSheet.getRange(currentRow,1,1,mapHeader.length);
    mapHeaderRange.setFontWeight('bold').setBackground('#4a90e2').setFontColor('white');
    currentRow++;

    // masterMap 再生成（上部 rank 解析時に存在しない場合に備え）
    var masterMap = {};
    try {
        var mSheet = SpreadsheetApp.openById(SPREADSHEET_IDS.TRAINING_MASTER).getSheetByName(SHEET_NAMES.TRAINING_MASTER);
        var mLast = mSheet.getLastRow();
        if (mLast > 4) {
            var mVals = mSheet.getRange(5,1,mLast-4,9).getValues();
            for (var mi=0; mi<mVals.length; mi++) {
                var mName = mVals[mi][2];
                if (mName && mName.toString().trim()!=='') {
                    masterMap[mName.trim()] = { day: mVals[mi][7], seq: mVals[mi][8] };
                }
            }
        }
    } catch(e) {
        writeLog('ERROR','マスタ読み込み失敗: '+e.message);
    }

    var diffRows = [];
    for (var gi=0; gi<trainingGroups.length; gi++) {
        var tg = trainingGroups[gi];
        var mi = masterMap[tg.name];
        if (!mi) continue;
        if (tg.implementationDay != mi.day || tg.sequence != mi.seq) {
            diffRows.push([tg.name, tg.implementationDay, mi.day, tg.sequence, mi.seq]);
        }
    }

    if (diffRows.length > 0) {
        var diffRange = mappingSheet.getRange(currentRow,1,diffRows.length,mapHeader.length);
        diffRange.setValues(diffRows);
        diffRange.setBorder(true,true,true,true,true,true);
        currentRow += diffRows.length;
    } else {
        mappingSheet.getRange(currentRow,1).setValue('すべて一致しました');
        currentRow += 1;
    }

    // --- 昼休憩枠チェック (12:00-13:00) ---
    mappingSheet.getRange(currentRow,1).setValue('【昼休憩枠チェック】');
    mappingSheet.getRange(currentRow,1).setFontWeight('bold');
    currentRow++;

    var lunchHeader = ['研修名','開始','終了'];
    mappingSheet.getRange(currentRow,1,1,lunchHeader.length).setValues([lunchHeader]);
    var lunchHeaderRange = mappingSheet.getRange(currentRow,1,1,lunchHeader.length);
    lunchHeaderRange.setFontWeight('bold').setBackground('#4a90e2').setFontColor('white');
    currentRow++;

    var lunchRows = [];
    var lunchStartHour = 12;
    var lunchEndHour = 13;
    for (var li=0; li<trainingGroups.length; li++) {
        var lg = trainingGroups[li];
        if (!lg.startTime || !lg.endTime) continue;
        // 開始が 13:00 未満かつ終了が 12:00 超過 → 重複
        if (lg.startTime.getHours() < lunchEndHour && lg.endTime.getHours() >= lunchStartHour) {
            lunchRows.push([
                lg.name,
                Utilities.formatDate(lg.startTime, 'Asia/Tokyo', 'MM/dd HH:mm'),
                Utilities.formatDate(lg.endTime, 'Asia/Tokyo', 'HH:mm')
            ]);
        }
    }

    if (lunchRows.length > 0) {
        var lunchRange = mappingSheet.getRange(currentRow,1,lunchRows.length,lunchHeader.length);
        lunchRange.setValues(lunchRows);
        lunchRange.setBorder(true,true,true,true,true,true);
        currentRow += lunchRows.length;
    } else {
        mappingSheet.getRange(currentRow,1).setValue('該当なし');
        currentRow += 1;
    }
}

// =========================================
// 公開API（既存コードとの互換性のため）
// =========================================

/**
 * 全研修のカレンダーイベントを作成（メイン関数 - リファクタリング版）
 * @param {Array<Object>} trainingGroups - 研修グループ配列
 * @param {Date} hireDate - 入社日
 * @returns {Array} 処理結果配列
 */
function createAllCalendarEvents_New(trainingGroups, hireDate) {
    var calendarManager = CalendarEventManager.getInstance();
    return calendarManager.createAllCalendarEvents(trainingGroups, hireDate);
}

/**
 * 会議室を確保（リファクタリング版）
 * @param {number} numberOfAttendees - 参加者数
 * @param {Date} startTime - 開始時間
 * @param {Date} endTime - 終了時間
 * @param {string} trainingName - 研修名
 * @returns {Object} 会議室予約結果
 */
function findAndReserveRoom_New(numberOfAttendees, startTime, endTime, trainingName) {
    var roomManager = RoomManager.getInstance();
    return roomManager.findAndReserveRoom(numberOfAttendees, startTime, endTime, trainingName);
}

/**
 * カレンダーイベントを削除（リファクタリング版）
 * @param {string} eventId - イベントID
 * @returns {boolean} 削除成功かどうか
 */
function deleteSingleCalendarEvent_New(eventId) {
    var calendarManager = CalendarEventManager.getInstance();
    return calendarManager.deleteSingleCalendarEvent(eventId);
}

// =========================================
// 従来のRoomReservationManager互換性
// =========================================

/**
 * 従来のRoomReservationManagerのシングルトン（互換性のため）
 */
var RoomReservationManager_New = (function() {
    return {
        getInstance: function() {
            return RoomManager.getInstance();
        }
    };
})();

writeLog('INFO', 'CalendarUtils.gs 完全読み込み完了');

// =========================================
// カレンダーイベント削除機能
// =========================================

/**
 * マッピングシートから全カレンダーイベントを削除する
 * @param {boolean} confirmDeletion - 削除前確認を行うかどうか（デフォルト: true）
 * @returns {Object} 削除結果の統計
 */
function deleteCalendarEventsFromMappingSheet(confirmDeletion) {
    writeLog('INFO', 'マッピングシートからカレンダーイベント削除を開始');
    
    // デフォルトで確認を行う
    if (confirmDeletion === undefined) {
        confirmDeletion = true;
    }
    
    try {
        // 最新のマッピングシートを取得
        var mappingSheet = getMostRecentMappingSheet();
        if (!mappingSheet) {
            throw new Error('マッピングシートが見つかりませんでした');
        }
        
        writeLog('INFO', 'マッピングシートを取得しました: ' + mappingSheet.getName());
        
        // カレンダーIDが格納されている列を特定
        var lastRow = mappingSheet.getLastRow();
        var calendarIdCol = MAPPING_COLS.CAL_ID;
        var eventNameCol = MAPPING_COLS.NAME;
        var resultStatCol = MAPPING_COLS.STATUS;
        
        if (lastRow <= 1) {
            throw new Error('マッピングシートにデータがありません');
        }
        
        // ヘッダー行を除いてデータ範囲を取得
        var dataRange = mappingSheet.getRange(2, 1, lastRow - 1, calendarIdCol + 1);
        var data = dataRange.getValues();
        
        // 削除対象のイベントを事前にカウント
        var deletionTargets = [];
        for (var i = 0; i < data.length; i++) {
            var row = data[i];
            var eventName = row[eventNameCol - 1];
            var calendarId = row[calendarIdCol - 1];
            
            if (calendarId && calendarId !== '' && calendarId !== '削除済み' && 
                calendarId !== '無効なIDのため削除済み' && calendarId !== 'エラーのため削除済み' &&
                calendarId !== '削除済み（既に存在しない）') {
                deletionTargets.push({
                    name: eventName,
                    id: calendarId,
                    rowIndex: i + 2
                });
            }
        }
        
        writeLog('INFO', '削除対象イベント数: ' + deletionTargets.length + '件');
        
        // 削除対象がない場合
        if (deletionTargets.length === 0) {
            writeLog('INFO', '削除対象のカレンダーイベントがありません');
            return {
                sheetName: mappingSheet.getName(),
                total: 0,
                success: 0,
                failed: 0,
                errors: [],
                message: '削除対象のイベントがありませんでした'
            };
        }
        
        // 削除前確認（UI表示）
        if (confirmDeletion) {
            var confirmMessage = '以下の' + deletionTargets.length + '件のカレンダーイベントを削除します:\n\n';
            for (var i = 0; i < Math.min(5, deletionTargets.length); i++) {
                confirmMessage += '• ' + deletionTargets[i].name + '\n';
            }
            if (deletionTargets.length > 5) {
                confirmMessage += '... 他' + (deletionTargets.length - 5) + '件\n';
            }
            confirmMessage += '\n実行しますか？';
            
            var ui = SpreadsheetApp.getUi();
            var response = ui.alert('カレンダーイベント削除確認', confirmMessage, ui.ButtonSet.YES_NO);
            
            if (response !== ui.Button.YES) {
                writeLog('INFO', 'ユーザーによりカレンダーイベント削除がキャンセルされました');
                return {
                    sheetName: mappingSheet.getName(),
                    total: 0,
                    success: 0,
                    failed: 0,
                    errors: [],
                    message: 'ユーザーによりキャンセルされました'
                };
            }
        }
        
        // 結果を格納する変数
        var result = {
            sheetName: mappingSheet.getName(),
            total: deletionTargets.length,
            success: 0,
            failed: 0,
            errors: []
        };
        
        // 各削除対象に対して処理（実際のGoogle Calendar API呼び出し）
        for (var i = 0; i < deletionTargets.length; i++) {
            var target = deletionTargets[i];
            var eventName = target.name;
            var calendarId = target.id;
            var rowIndex = target.rowIndex;
            
            // カレンダーIDの情報をログに記録（デバッグ）
            writeLog('DEBUG', 'カレンダーID検証 [' + rowIndex + '行目]: "' + eventName + 
                    '" - ID=[' + calendarId + '], 型=' + typeof calendarId);
            
            try {
                // カレンダーIDの基本的な形式チェック
                if (typeof calendarId !== 'string' || calendarId.length < 10) {
                    throw new Error('無効なカレンダーID形式: ' + calendarId);
                }
                
                // 実際にGoogle Calendar APIを呼び出してイベントを削除
                writeLog('INFO', 'カレンダーイベント削除開始: "' + eventName + '" (ID: ' + calendarId + ')');
                
                var event = CalendarApp.getEventById(calendarId);
                if (!event) {
                    writeLog('WARN', 'カレンダーイベントが見つかりません（既に削除済みの可能性）: ' + calendarId);
                    // 見つからない場合も成功として扱う（既に削除済み）
                    mappingSheet.getRange(rowIndex, calendarIdCol).setValue('削除済み（既に存在しない）');
                    mappingSheet.getRange(rowIndex, resultStatCol).setValue('削除済み');
                    result.success++;
                    continue;
                }
                
                // イベントの詳細をログに記録（削除前確認）
                var eventTitle = event.getTitle();
                var eventStart = event.getStartTime();
                var eventEnd = event.getEndTime();
                writeLog('INFO', '削除対象イベント詳細: タイトル="' + eventTitle + '", 時間=' + 
                         Utilities.formatDate(eventStart, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + 
                         Utilities.formatDate(eventEnd, 'Asia/Tokyo', 'HH:mm'));
                
                // 実際にイベントを削除
                event.deleteEvent();
                
                // シート上で削除済みとマーク
                mappingSheet.getRange(rowIndex, calendarIdCol).setValue('削除済み');
                mappingSheet.getRange(rowIndex, resultStatCol).setValue('削除済み');
                
                writeLog('INFO', 'カレンダーイベント削除成功: "' + eventName + '" (元ID: ' + calendarId + ')');
                result.success++;
                
                // API制限対策で少し待機
                Utilities.sleep(500); // 0.5秒待機
                
            } catch (e) {
                result.failed++;
                var errorMessage = 'カレンダーイベント削除エラー: "' + eventName + '" - ' + e.message;
                result.errors.push(errorMessage);
                writeLog('ERROR', errorMessage);
                
                // エラー発生時の処理
                try {
                    if (e.message.indexOf('無効な引数') !== -1 || e.message.indexOf('Invalid argument') !== -1) {
                        mappingSheet.getRange(rowIndex, calendarIdCol).setValue('無効なIDのため削除済み');
                        mappingSheet.getRange(rowIndex, resultStatCol).setValue('無効ID: ' + e.message);
                    } else {
                        mappingSheet.getRange(rowIndex, calendarIdCol).setValue('エラーのため削除済み');
                        mappingSheet.getRange(rowIndex, resultStatCol).setValue('エラー: ' + e.message);
                    }
                } catch (err) {
                    writeLog('ERROR', 'シート更新でさらにエラー: ' + err.message);
                }
            }
        }
        
        writeLog('INFO', 'カレンダーイベント削除処理完了: 成功=' + result.success + '/' + result.total + ', 失敗=' + result.failed);
        return result;
        
    } catch (e) {
        writeLog('ERROR', 'deleteCalendarEventsFromMappingSheet でエラー: ' + e.message);
        throw e;
    }
}

/**
 * 最新のマッピングシートを取得する
 * @returns {Object|null} シートオブジェクトまたはnull
 */
function getMostRecentMappingSheet() {
    try {
        // マッピングシートはEXECUTIONスプレッドシート内に作成されるため、EXECUTION IDを使用
        var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_IDS.EXECUTION);
        var sheets = spreadsheet.getSheets();
        
        writeLog('DEBUG', 'スプレッドシート内のシート数: ' + sheets.length);
        
        // マッピングシート名は通常「マッピング結果_YYYYMMDD」の形式
        var mappingSheets = [];
        
        for (var i = 0; i < sheets.length; i++) {
            var sheetName = sheets[i].getName();
            writeLog('DEBUG', 'シート名確認: ' + sheetName);
            if (sheetName.indexOf('マッピング結果_') === 0) {
                mappingSheets.push(sheets[i]);
                writeLog('DEBUG', 'マッピングシートを発見: ' + sheetName);
            }
        }
        
        if (mappingSheets.length === 0) {
            writeLog('WARN', 'マッピングシートが見つかりませんでした');
            return null;
        }
        
        // 最新のシートを取得（名前でソート）
        mappingSheets.sort(function(a, b) {
            return b.getName().localeCompare(a.getName());
        });
        
        var latestSheet = mappingSheets[0];
        writeLog('INFO', '最新のマッピングシートを選択: ' + latestSheet.getName());
        return latestSheet;
        
    } catch (e) {
        writeLog('ERROR', 'getMostRecentMappingSheet でエラー: ' + e.message);
        throw new Error('マッピングシートの取得に失敗しました: ' + e.message);
    }
}

/**
 * 特定の研修に関連するカレンダーイベントを削除する
 * @param {string} trainingName - 研修名
 * @returns {boolean} 削除成功かどうか
 */
function deleteSpecificTrainingEvent(trainingName) {
    writeLog('INFO', '特定研修のカレンダーイベント削除を開始: "' + trainingName + '"');
    
    try {
        // マッピングシートを取得
        var mappingSheet = getMostRecentMappingSheet();
        if (!mappingSheet) {
            throw new Error('マッピングシートが見つかりませんでした');
        }
        
        // ヘッダー行を除いて全データを取得
        var lastRow = mappingSheet.getLastRow();
        var dataRange = mappingSheet.getRange(2, 1, lastRow - 1, 9); // A列（研修名）からI列（処理状況）まで
        var data = dataRange.getValues();
        
        var found = false;
        var success = false;
        
        for (var i = 0; i < data.length; i++) {
            var currentTrainingName = data[i][0]; // A列：研修名
            var calendarId = data[i][MAPPING_COLS.CAL_ID -1]; // zero index
            
            if (currentTrainingName === trainingName && calendarId && 
                calendarId !== '削除済み' && calendarId !== '無効なIDのため削除済み' && 
                calendarId !== 'エラーのため削除済み') {
                
                found = true;
                
                // カレンダーIDの情報をデバッグログに出力
                writeLog('DEBUG', '特定研修ID検証: "' + currentTrainingName + '" - ID=[' + calendarId + '], 型=' + typeof calendarId);
                
                try {
                    // カレンダーIDの基本的な形式チェック
                    if (typeof calendarId !== 'string' || calendarId.length < 10) {
                        throw new Error('無効なカレンダーID形式: ' + calendarId);
                    }
                    
                    // 実際にGoogle Calendar APIを呼び出してイベントを削除
                    writeLog('INFO', '特定研修のカレンダーイベント削除開始: "' + trainingName + '" (ID: ' + calendarId + ')');
                    
                    var event = CalendarApp.getEventById(calendarId);
                    if (!event) {
                        writeLog('WARN', '特定研修のカレンダーイベントが見つかりません（既に削除済みの可能性）: ' + calendarId);
                        // 見つからない場合も成功として扱う（既に削除済み）
                        mappingSheet.getRange(i + 2, MAPPING_COLS.CAL_ID).setValue('削除済み（既に存在しない）');
                        mappingSheet.getRange(i + 2, MAPPING_COLS.STATUS).setValue('削除済み');
                        success = true;
                        break;
                    }
                    
                    // イベントの詳細をログに記録（削除前確認）
                    var eventTitle = event.getTitle();
                    var eventStart = event.getStartTime();
                    var eventEnd = event.getEndTime();
                    writeLog('INFO', '削除対象イベント詳細: タイトル="' + eventTitle + '", 時間=' + 
                             Utilities.formatDate(eventStart, 'Asia/Tokyo', 'MM/dd HH:mm') + '-' + 
                             Utilities.formatDate(eventEnd, 'Asia/Tokyo', 'HH:mm'));
                    
                    // 実際にイベントを削除
                    event.deleteEvent();
                    
                    // シート上で削除済みとマーク
                    mappingSheet.getRange(i + 2, MAPPING_COLS.CAL_ID).setValue('削除済み');
                    mappingSheet.getRange(i + 2, MAPPING_COLS.STATUS).setValue('削除済み');
                    
                    writeLog('INFO', '特定研修のカレンダーイベント削除成功: "' + trainingName + '" (ID: ' + calendarId + ')');
                    success = true;
                    break;
                    
                } catch (e) {
                    writeLog('ERROR', '特定研修のカレンダーイベント削除エラー: "' + trainingName + '" - ' + e.message);
                    // エラー時の処理
                    try {
                        if (e.message.indexOf('無効な引数') !== -1 || e.message.indexOf('Invalid argument') !== -1) {
                            mappingSheet.getRange(i + 2, MAPPING_COLS.CAL_ID).setValue('無効なIDのため削除済み');
                            mappingSheet.getRange(i + 2, MAPPING_COLS.STATUS).setValue('無効ID: ' + e.message);
                        } else {
                            mappingSheet.getRange(i + 2, MAPPING_COLS.CAL_ID).setValue('エラーのため削除済み');
                            mappingSheet.getRange(i + 2, MAPPING_COLS.STATUS).setValue('エラー: ' + e.message);
                        }
                    } catch (err) {
                        writeLog('ERROR', 'シート更新でさらにエラー: ' + err.message);
                    }
                    return false;
                }
            }
        }
        
        if (!found) {
            writeLog('WARN', '指定された研修が見つからないか、既に削除されています: "' + trainingName + '"');
            return false;
        }
        
        return success;
        
    } catch (e) {
        writeLog('ERROR', 'deleteSpecificTrainingEvent でエラー: ' + e.message);
        throw e;
    }
}

/**
 * カレンダーIDを指定してイベントを削除する
 * @param {string} calendarId - カレンダーイベントID
 * @returns {boolean} 削除成功かどうか
 */
function deleteEventByCalendarId(calendarId) {
    writeLog('INFO', 'カレンダーID指定削除を開始: ' + calendarId);
    try {
        if (!calendarId || typeof calendarId !== 'string') {
            throw new Error('カレンダーIDが無効です');
        }

        // マッピングシート取得
        var mappingSheet = getMostRecentMappingSheet();
        if (!mappingSheet) {
            throw new Error('マッピングシートが見つかりません');
        }

        var lastRow = mappingSheet.getLastRow();
        if (lastRow <= 1) {
            throw new Error('マッピングシートにデータがありません');
        }

        // データ取得（ヘッダー除く）
        var dataRange = mappingSheet.getRange(2, 1, lastRow - 1, MAPPING_COLS.CAL_ID);
        var data = dataRange.getValues();
        var foundRow = -1;
        for (var i = 0; i < data.length; i++) {
            var idCell = data[i][MAPPING_COLS.CAL_ID - 1];
            if (idCell === calendarId) {
                foundRow = i + 2; // sheet row index
                break;
            }
        }

        if (foundRow === -1) {
            writeLog('WARN', '指定されたカレンダーIDがマッピングシートに存在しません: ' + calendarId);
            // それでも削除を試みる
        }

        var event = CalendarApp.getEventById(calendarId);
        if (!event) {
            writeLog('WARN', 'カレンダーイベントが見つかりません（既に削除済みの可能性）: ' + calendarId);
            if (foundRow !== -1) {
                mappingSheet.getRange(foundRow, MAPPING_COLS.CAL_ID).setValue('削除済み（既に存在しない）');
                mappingSheet.getRange(foundRow, MAPPING_COLS.STATUS).setValue('削除済み');
            }
            return true;
        }

        // 削除実行
        event.deleteEvent();
        writeLog('INFO', 'カレンダーID削除成功: ' + calendarId);

        if (foundRow !== -1) {
            mappingSheet.getRange(foundRow, MAPPING_COLS.CAL_ID).setValue('削除済み');
            mappingSheet.getRange(foundRow, MAPPING_COLS.STATUS).setValue('削除済み');
        }

        return true;
    } catch (e) {
        writeLog('ERROR', 'deleteEventByCalendarId でエラー: ' + e.message);
        return false;
    }
} 