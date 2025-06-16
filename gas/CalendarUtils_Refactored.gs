// =========================================
// カレンダー関連ユーティリティ（リファクタリング版）
// =========================================

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

writeLog('INFO', 'CalendarUtils.gs (リファクタリング版) - SequenceManager を読み込みました');

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

                // 営業時間チェック
                if (!this.isWithinBusinessHours(proposedEnd)) {
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
                    baseHour = 9; baseMinute = 0;
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
                
                // 実施順1番の場合は基準時間から開始
                if (sequence === 1) {
                    writeLog('DEBUG', '実施順1番のため基準時間から開始: ' + Utilities.formatDate(baseStartTime, 'Asia/Tokyo', 'HH:mm'));
                    return baseStartTime;
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
                    // 直前の実施順がない場合、同実施日の最新研修の後から開始
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
             * @returns {boolean} 営業時間内かどうか
             */
            isWithinBusinessHours: function(endTime) {
                var maxEndHour = 19; // 19時まで
                return endTime.getHours() < maxEndHour || (endTime.getHours() === maxEndHour && endTime.getMinutes() === 0);
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

writeLog('INFO', 'CalendarUtils.gs (リファクタリング版) - TimeSlotCalculator を読み込みました');

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

writeLog('INFO', 'CalendarUtils.gs (リファクタリング版) - RoomManager を読み込みました');

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
                
                var results = [];
                
                for (var i = 0; i < trainingGroups.length; i++) {
                    var group = trainingGroups[i];
                    
                    try {
                        writeLog('INFO', '研修処理開始 (' + (i + 1) + '/' + trainingGroups.length + '): ' + group.name);
                        
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
                
                // 参加者の時間重複チェック
                if (!this.isTimeSlotAvailable(eventTime.start, eventTime.end, trainingGroup)) {
                    throw new Error('参加者または講師の時間が重複しています');
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
                // 講師の空き時間チェック
                var lecturerEmails = trainingGroup.lecturerEmails || (trainingGroup.lecturer ? [trainingGroup.lecturer] : []);
                for (var i = 0; i < lecturerEmails.length; i++) {
                    if (!this.isLecturerAvailable(lecturerEmails[i], startTime, endTime)) {
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
                        
                        if (!(endTime <= eventStart || startTime >= eventEnd)) {
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
                try {
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

writeLog('INFO', 'CalendarUtils.gs (リファクタリング版) - CalendarEventManager を読み込みました');

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

writeLog('INFO', 'CalendarUtils.gs (リファクタリング版) 完全読み込み完了');

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

                // 営業時間チェック
                if (!this.isWithinBusinessHours(proposedEnd)) {
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
                    baseHour = 9; baseMinute = 0;
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
                
                // 実施順1番の場合は基準時間から開始
                if (sequence === 1) {
                    writeLog('DEBUG', '実施順1番のため基準時間から開始: ' + Utilities.formatDate(baseStartTime, 'Asia/Tokyo', 'HH:mm'));
                    return baseStartTime;
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
                    // 直前の実施順がない場合、同実施日の最新研修の後から開始
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
             * @returns {boolean} 営業時間内かどうか
             */
            isWithinBusinessHours: function(endTime) {
                var maxEndHour = 19; // 19時まで
                return endTime.getHours() < maxEndHour || (endTime.getHours() === maxEndHour && endTime.getMinutes() === 0);
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
                
                var results = [];
                
                for (var i = 0; i < trainingGroups.length; i++) {
                    var group = trainingGroups[i];
                    
                    try {
                        writeLog('INFO', '研修処理開始 (' + (i + 1) + '/' + trainingGroups.length + '): ' + group.name);
                        
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
                
                // 参加者の時間重複チェック
                if (!this.isTimeSlotAvailable(eventTime.start, eventTime.end, trainingGroup)) {
                    throw new Error('参加者または講師の時間が重複しています');
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
                for (var i = 0; i < lecturerEmails.length; i++) {
                    if (!this.isLecturerAvailable(lecturerEmails[i], startTime, endTime)) {
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
                        
                        if (!(endTime <= eventStart || startTime >= eventEnd)) {
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
                try {
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

// =========================================
// 公開API（既存コードとの互換性のため）
// =========================================

/**
 * 全研修のカレンダーイベントを作成（メイン関数）
 * @param {Array<Object>} trainingGroups - 研修グループ配列
 * @param {Date} hireDate - 入社日
 * @returns {Array} 処理結果配列
 */
function createAllCalendarEvents(trainingGroups, hireDate) {
    var calendarManager = CalendarEventManager.getInstance();
    return calendarManager.createAllCalendarEvents(trainingGroups, hireDate);
}

/**
 * 会議室を確保（互換性のため）
 * @param {number} numberOfAttendees - 参加者数
 * @param {Date} startTime - 開始時間
 * @param {Date} endTime - 終了時間
 * @param {string} trainingName - 研修名
 * @returns {Object} 会議室予約結果
 */
function findAndReserveRoom(numberOfAttendees, startTime, endTime, trainingName) {
    var roomManager = RoomManager.getInstance();
    return roomManager.findAndReserveRoom(numberOfAttendees, startTime, endTime, trainingName);
}

/**
 * カレンダーイベントを削除（互換性のため）
 * @param {string} eventId - イベントID
 * @returns {boolean} 削除成功かどうか
 */
function deleteSingleCalendarEvent(eventId) {
    var calendarManager = CalendarEventManager.getInstance();
    return calendarManager.deleteSingleCalendarEvent(eventId);
}

// =========================================
// 従来のRoomReservationManager互換性
// =========================================

/**
 * 従来のRoomReservationManagerのシングルトン（互換性のため）
 */
var RoomReservationManager = (function() {
    return {
        getInstance: function() {
            return RoomManager.getInstance();
        }
    };
})(); 