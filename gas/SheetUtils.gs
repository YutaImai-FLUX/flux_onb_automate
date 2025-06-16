// =========================================
// スプレッドシート関連ユーティリティ
// =========================================

/**
 * `入社者リスト`から未処理のデータを取得する
 * @param {Date} targetHireDate - 対象の入社日（指定された場合、この日付の入社者のみを取得）
 * @returns {Array<Object>} 入社者情報の配列
 */
function getNewHires(targetHireDate) {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_IDS.NEW_HIRES).getSheetByName(SHEET_NAMES.NEW_HIRES);
  var lastRow = sheet.getLastRow();
  writeLog('DEBUG', '入社者リストの最終行: ' + lastRow);
  
  if (lastRow <= 1) {
    writeLog('DEBUG', 'データが存在しません');
    return [];
  }
  
  var data = sheet.getRange(2, 1, lastRow - 1, 15).getValues(); // 15列取得（所属列追加後）
  writeLog('DEBUG', '取得したデータ行数: ' + data.length);
  
  var newHires = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var rank = row[0];           // A列: 職位
    var experience = row[1];     // B列: 経験有無
    var name = row[2];           // C列: 名前
    var email = row[3];          // D列: メールアドレス
    var department = row[4];     // E列: 所属
    var hireDate = row[5];       // F列: 入社日
    
    writeLog('DEBUG', '【デバッグ】行' + (i + 2) + ' 生データ: ' + JSON.stringify(row));
    writeLog('DEBUG', '行' + (i + 2) + ': 職位=' + rank + ', 経験=' + experience + ', 氏名=' + name + ', メール=' + email + ', 所属=' + department + ', 入社日=' + hireDate);
    
    // A~F列（職位、経験有無、名前、メールアドレス、所属、入社日）が全て入力されている行のみを対象とする
    if (rank && rank.trim() !== '' && 
        experience && experience.trim() !== '' && 
        name && name.trim() !== '' && 
        email && email.trim() !== '' && 
        department && department.trim() !== '' &&
        hireDate && hireDate !== '') {
      
      // 入社日の形式チェック（Date型への変換チェック）
      var parsedHireDate = null;
      try {
        if (hireDate instanceof Date) {
          parsedHireDate = new Date(hireDate);
        } else if (typeof hireDate === 'string' && hireDate.trim() !== '') {
          parsedHireDate = new Date(hireDate.trim());
        } else if (typeof hireDate === 'number') {
          // Excelのシリアル値の場合
          parsedHireDate = new Date(hireDate);
        }
        
        // 有効な日付かチェック
        if (!parsedHireDate || isNaN(parsedHireDate.getTime())) {
          writeLog('WARN', '行' + (i + 2) + ': 無効な入社日形式のためスキップ: ' + hireDate);
          continue;
        }
      } catch (e) {
        writeLog('WARN', '行' + (i + 2) + ': 入社日解析エラーのためスキップ: ' + hireDate + ' - ' + e.message);
        continue;
      }
      
      // 対象入社日による絞り込み（指定された場合のみ）
      if (targetHireDate) {
        var targetDateOnly = new Date(targetHireDate.getFullYear(), targetHireDate.getMonth(), targetHireDate.getDate());
        var parsedDateOnly = new Date(parsedHireDate.getFullYear(), parsedHireDate.getMonth(), parsedHireDate.getDate());
        
        if (targetDateOnly.getTime() !== parsedDateOnly.getTime()) {
          writeLog('DEBUG', '行' + (i + 2) + ': 入社日不一致のためスキップ: 対象日=' + 
                   Utilities.formatDate(targetHireDate, 'Asia/Tokyo', 'yyyy/MM/dd') + 
                   ', 実際=' + Utilities.formatDate(parsedHireDate, 'Asia/Tokyo', 'yyyy/MM/dd'));
          continue;
        }
      }
      
      newHires.push({
        rowNum: i + 2, // 実際の行番号
        rank: rank.trim(),         // A列: 職位
        experience: experience.trim(), // B列: 経験有無
        name: name.trim(),         // C列: 名前
        email: email.trim(),       // D列: メールアドレス
        department: department.trim(), // E列: 所属
        hireDate: parsedHireDate,  // F列: 入社日
      });
      
      writeLog('INFO', '入社者追加: ' + name + ' (職位:' + rank + ', 経験:' + experience + ', 所属:' + department + ', メール:' + email + ', 入社日:' + Utilities.formatDate(parsedHireDate, 'Asia/Tokyo', 'yyyy/MM/dd') + ')');
    } else {
      var missingFields = [];
      if (!rank || rank.trim() === '') missingFields.push('職位');
      if (!experience || experience.trim() === '') missingFields.push('経験有無');
      if (!name || name.trim() === '') missingFields.push('名前');
      if (!email || email.trim() === '') missingFields.push('メールアドレス');
      if (!department || department.trim() === '') missingFields.push('所属');
      if (!hireDate || hireDate === '') missingFields.push('入社日');
      
      writeLog('DEBUG', '行' + (i + 2) + ': 必須項目未入力のためスキップ (未入力: ' + missingFields.join(', ') + ')');
    }
  }
  writeLog('DEBUG', '対象入社者数: ' + newHires.length);
  
  // フォールバック: 指定入社日の対象者が0名の場合、日付フィルタを外して全対象を再取得
  if (newHires.length === 0 && targetHireDate) {
    writeLog('WARN', '指定入社日(' + Utilities.formatDate(targetHireDate, 'Asia/Tokyo', 'yyyy/MM/dd') + ')の対象入社者が0名でした。日付フィルタを外して再検索します。');
    return getNewHires(null); // fallbackUsedは呼び出し側で判定できるようにログで通知
  }
  return newHires;
}

/**
 * 処理が完了した入社者のステータスを「済」に更新する
 * @param {Array<Object>} processedHires - 処理済み入社者情報の配列
 */
function updateStatuses(processedHires) {
    // 処理状況列が存在しないため、ステータス更新をスキップ
    writeLog('INFO', '処理完了: ' + processedHires.length + '名の入社者のカレンダー招待を作成しました');
}

/**
 * 実行ログを`実行シート`に記録する
 * 注意：2行目はヘッダー行のため、3行目以降に記録する
 * 列構成：A=実行者, B=入社日, C=処理結果, D=詳細メッセージ, E=実行日時
 * @param {Object} params - 実行パラメータ
 * @param {string} status - 処理結果 (成功/エラー)
 * @param {string} message - 詳細メッセージ
 */
function logExecution(params, status, message) {
    // 3行目（データ行）に記録
    // 2行目はヘッダー行のため使用しない
    var logRow = 3; // データ行の開始位置
    
    writeLog('DEBUG', '実行ログ記録: 行=' + logRow + ', 状態=' + status);
    
    // 新しい列構成に合わせて記録
    params.sheet.getRange(logRow, 1).setValue(params.user);       // A列: 実行者
    params.sheet.getRange(logRow, 2).setValue(params.hireDate);   // B列: 入社日
    params.sheet.getRange(logRow, 3).setValue(status);            // C列: 処理結果
    params.sheet.getRange(logRow, 4).setValue(message);           // D列: 詳細メッセージ
    params.sheet.getRange(logRow, 5).setValue(new Date());        // E列: 実行日時
}

/**
 * インクリメンタル処理用のマッピングシートを作成する
 * @param {Array<Object>} trainingGroups - 研修グループの配列
 * @param {Array<Object>} allNewHires - 全入社者の配列
 * @param {Date} periodStart - 研修期間開始日
 * @param {Date} periodEnd - 研修期間終了日
 * @returns {Object} 作成されたマッピングシート
 */
function createIncrementalMappingSheet(trainingGroups, allNewHires, periodStart, periodEnd) {
    writeLog('INFO', 'インクリメンタルマッピングシート作成開始');
    writeLog('DEBUG', '引数確認: trainingGroups=' + trainingGroups.length + ', allNewHires=' + allNewHires.length);
    
    var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_IDS.EXECUTION);
    var sheetName = 'マッピング結果_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd_HH-mm');
    
    // 既存のマッピングシートがあれば削除
    var existingSheets = spreadsheet.getSheets();
    for (var i = 0; i < existingSheets.length; i++) {
        if (existingSheets[i].getName().indexOf('マッピング結果_') === 0) {
            spreadsheet.deleteSheet(existingSheets[i]);
        }
    }
    
    // 新しいシートを作成
    var mappingSheet = spreadsheet.insertSheet(sheetName);
    writeLog('INFO', '新しいシート作成: ' + sheetName);
    
    // ヘッダー行を作成
    var headers = ['研修名', '対象者', '講師', '参加者数', '会議室要否', '会議室名', '実施日(営業日)', '実施順', '研修実施日時', 'カレンダーID', '処理状況', 'エラー詳細'];
    
    // ヘッダーを設定
    mappingSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // ヘッダー行のスタイル設定
    var headerRange = mappingSheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#4a90e2');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    headerRange.setWrap(true);
    headerRange.setVerticalAlignment('middle');
    
    // 各研修の初期行を作成（処理状況は「待機中」に設定）
    var dataRows = [];
    for (var i = 0; i < trainingGroups.length; i++) {
        var group = trainingGroups[i];
        
        if (!group || !group.name) {
            writeLog('ERROR', '研修グループ[' + i + ']が無効です。スキップします。');
            continue;
        }
        
        // 参加者リスト作成
        var participants = [];
        for (var j = 0; j < allNewHires.length; j++) {
            var hire = allNewHires[j];
            if (group.attendees && group.attendees.indexOf(hire.email) !== -1) {
                participants.push(hire.name + '(' + hire.rank + '/' + hire.experience + '/' + hire.department + ')');
            }
        }
        var participantsList = participants.join('\n');
        
        // 参加者数計算
        var lecturerCount = (group.lecturerEmails && group.lecturerEmails.length > 0) ? group.lecturerEmails.length : (group.lecturer ? 1 : 0);
        var participantCount = (participants ? participants.length : 0) + lecturerCount;
        
        var row = [
            group.name,
            participantsList,
            group.lecturerNames ? group.lecturerNames.join('\n') : '',
            participants.length + '/' + (group.needsRoom ? group.attendees.length : 0),
            group.needsRoom ? '必要' : '不要',
            '', // 会議室名
            group.implementationDay || '', // 実施日(営業日)
            group.sequence || '', // 実施順
            '', // 研修実施日時
            '', // カレンダーID
            '待機中',
            ''  // エラー詳細
        ];
        
        dataRows.push(row);
    }
    
    // データを設定
    if (dataRows.length > 0) {
        mappingSheet.getRange(2, 1, dataRows.length, dataRows[0].length).setValues(dataRows);
        writeLog('INFO', 'データ書き込み成功: ' + dataRows.length + '行');
        
        // スタイル設定
        var dataRange = mappingSheet.getRange(2, 1, dataRows.length, headers.length);
        dataRange.setBorder(true, true, true, true, true, true);
        dataRange.setVerticalAlignment('top');
        
        // 列幅設定
        mappingSheet.setColumnWidth(1, 250);  // 研修名
        mappingSheet.setColumnWidth(2, 300);  // 対象者
        mappingSheet.setColumnWidth(3, 200);  // 講師
        mappingSheet.setColumnWidth(4, 80);   // 参加者数
        mappingSheet.setColumnWidth(5, 100);  // 会議室要否
        mappingSheet.setColumnWidth(6, 150);  // 会議室名
        mappingSheet.setColumnWidth(7, 180);  // 実施日(営業日)
        mappingSheet.setColumnWidth(8, 100);  // 実施順
        mappingSheet.setColumnWidth(9, 120);  // 研修実施日時
        mappingSheet.setColumnWidth(10, 300); // カレンダーID
        mappingSheet.setColumnWidth(11, 120);  // 処理状況
        mappingSheet.setColumnWidth(12, 300); // エラー詳細
        
        // 列の固定
        mappingSheet.setFrozenColumns(1);
    }
    
    writeLog('INFO', 'インクリメンタルマッピングシート作成完了: ' + trainingGroups.length + '件の研修');
    return mappingSheet;
}

/**
 * 研修と入社者のマッピング結果をシートに表示する（従来版）
 * @param {Array<Object>} trainingGroups - 研修グループの配列
 * @param {Array<Object>} allNewHires - 全入社者の配列
 * @param {Date} periodStart - 研修期間開始日
 * @param {Date} periodEnd - 研修期間終了日
 */
function createMappingSheet(trainingGroups, allNewHires, periodStart, periodEnd) {
    writeLog('INFO', 'マッピングシート作成開始 - 新フォーマット版');
    writeLog('DEBUG', '引数確認: trainingGroups=' + trainingGroups.length + ', allNewHires=' + allNewHires.length + 
             ', periodStart=' + periodStart + ', periodEnd=' + periodEnd);
    
    var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_IDS.EXECUTION);
    var sheetName = 'マッピング結果_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd_HH-mm');
    
    // 既存のマッピングシートがあれば削除
    var existingSheets = spreadsheet.getSheets();
    for (var i = 0; i < existingSheets.length; i++) {
        if (existingSheets[i].getName().indexOf('マッピング結果_') === 0) {
            spreadsheet.deleteSheet(existingSheets[i]);
        }
    }
    
    // 新しいシートを作成
    var mappingSheet = spreadsheet.insertSheet(sheetName);
    writeLog('INFO', '新しいシート作成: ' + sheetName);
    
    // ヘッダー行を作成（講師列をB列の次に移動）
    var headers = ['研修名', '対象者', '講師', '参加者数', '会議室要否', '会議室名', '実施日(営業日)', '実施順', '研修実施日時', 'カレンダーID'];
    
    // ヘッダーを設定
    mappingSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // ヘッダー行のスタイル設定
    var headerRange = mappingSheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#4a90e2');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    headerRange.setWrap(true);
    headerRange.setVerticalAlignment('middle');
    
    // データ行を作成
    writeLog('DEBUG', '新フォーマットでデータ行を作成します');
    writeLog('DEBUG', 'trainingGroups配列の長さ: ' + trainingGroups.length);
    
    var dataRows = [];
    for (var i = 0; i < trainingGroups.length; i++) {
        var group = trainingGroups[i];
        
        // グループオブジェクトの詳細ログ出力（デバッグ用）
        writeLog('DEBUG', '研修グループ[' + i + ']の内容: ' + JSON.stringify(group));
        
        // group が null または undefined の場合をチェック
        if (!group) {
            writeLog('ERROR', '研修グループ[' + i + ']が null または undefined です。スキップします。');
            continue;
        }
        
        // group.name が存在しない場合をチェック
        if (!group.name) {
            writeLog('ERROR', '研修グループ[' + i + ']にnameプロパティがありません。スキップします。グループ内容: ' + JSON.stringify(group));
            continue;
        }
        
        writeLog('DEBUG', '研修グループ処理中: ' + group.name + ' (順番: ' + (i + 1) + '/' + trainingGroups.length + ')');
        
        // attendees が存在しない場合をチェック
        if (!group.attendees || !Array.isArray(group.attendees)) {
            writeLog('ERROR', '研修グループ[' + i + '](' + group.name + ')にattendeesプロパティがないか配列ではありません。スキップします。');
            continue;
        }
        
        // 参加者リスト作成（講師を除く）
        var participants = [];
        for (var j = 0; j < allNewHires.length; j++) {
            var hire = allNewHires[j];
            if (group.attendees.indexOf(hire.email) !== -1) {
                participants.push(hire.name + '(' + hire.rank + '/' + hire.experience + '/' + hire.department + ')');
            }
        }
        var participantsList = participants.join('\n');
        
        // 参加者数（入社者＋講師の合計）
        var lecturerCount = (group.lecturerEmails && group.lecturerEmails.length > 0) ? group.lecturerEmails.length : (group.lecturer ? 1 : 0);
        var participantCount = (participants ? participants.length : 0) + lecturerCount;
        
        writeLog('DEBUG', '参加者数計算 (表示用): ' + group.name + ' - 入社者' + (participants ? participants.length : 0) + '名, 講師' + lecturerCount + '名, 合計' + participantCount + '名');
        
        // 会議室名を取得（実際の会議室名を表示）
        var roomNameInfo = { roomName: '' };
        var needsRoom = group.needsRoom || false;
        if (needsRoom && participantCount > 0) {
            try {
                // 実際の会議室確保を試行（participantCountには既に講師数が含まれている）
                var roomResult = findAvailableRoomName(participantCount);
                // 戻り値がオブジェクトの場合とstring の場合を考慮
                if (typeof roomResult === 'object' && roomResult.roomName) {
                    roomNameInfo.roomName = roomResult.roomName;
                } else if (typeof roomResult === 'string') {
                    roomNameInfo.roomName = roomResult;
                } else {
                    roomNameInfo.roomName = 'オンライン';
                }
                writeLog('DEBUG', '会議室確保: ' + roomNameInfo.roomName + ' (総参加者: ' + participantCount + '名, 内訳: 入社者' + participants.length + '名+講師' + lecturerCount + '名)');
            } catch (e) {
                roomNameInfo.roomName = '会議室確保失敗';
                writeLog('WARN', '会議室確保失敗: ' + e.message + ' (研修: ' + group.name + ')');
            }
        } else {
            roomNameInfo.roomName = needsRoom ? '参加者なし' : 'オンライン';
        }
        
        // 研修実施日時を計算
        var trainingDateTime = '';
        var groupTime = group.time || '60分';
        writeLog('DEBUG', 'G列日時計算開始: ' + group.name + ' (implementationDay: ' + (group.implementationDay || 'なし') + 
                 ', sequence: ' + (group.sequence || 'なし') + ', periodStart: ' + 
                 (periodStart ? Utilities.formatDate(periodStart, 'Asia/Tokyo', 'yyyy/MM/dd') : 'null') + 
                 ', periodEnd: ' + (periodEnd ? Utilities.formatDate(periodEnd, 'Asia/Tokyo', 'yyyy/MM/dd') : 'null') + ')');
        
        try {
            if (periodStart && periodEnd) {
                var eventDate = new Date(periodStart);
                // implementationDayが設定されていればそれを使用、なければsequenceを使用
                var implementationDay = group.implementationDay || group.sequence || 1;
                
                // 平日ベースで日付を計算（1営業日目は入社日当日）
                var businessDaysToAdd = implementationDay - 1;
                var daysAdded = 0;
                while (daysAdded < businessDaysToAdd) {
                    eventDate.setDate(eventDate.getDate() + 1);
                    if (eventDate.getDay() !== 0 && eventDate.getDay() !== 6) {
                        daysAdded++;
                    }
                }
                
                // 期間内チェック
                if (eventDate > periodEnd) {
                    writeLog('WARN', '計算された日付が期間終了日を超過、期間終了日に調整: ' + group.name);
                    eventDate = new Date(periodEnd);
                }
                
                var timeMatch = groupTime.match(/(\d+)分/);
                var durationMinutes = timeMatch ? parseInt(timeMatch[1]) : 60;
                var endHour = 9 + Math.floor(durationMinutes / 60);
                var endMinute = durationMinutes % 60;
                
                trainingDateTime = Utilities.formatDate(eventDate, 'Asia/Tokyo', 'MM/dd(E)') + 
                                 ' 9:00-' + 
                                 (endHour < 10 ? '0' : '') + endHour + ':' + 
                                 (endMinute < 10 ? '0' : '') + endMinute + 
                                 ' (' + groupTime + ')';
                
                writeLog('DEBUG', 'G列日時計算結果: ' + group.name + ' → ' + trainingDateTime + 
                         ' (実施日: ' + implementationDay + '営業日目)');
            } else {
                trainingDateTime = '日程未定 (' + groupTime + ')';
                writeLog('WARN', 'G列日時計算: 期間情報が不足のため日程未定に設定: ' + group.name);
            }
        } catch (e) {
            writeLog('ERROR', 'G列日時計算エラー: ' + e.message + ' (研修: ' + group.name + ')');
            trainingDateTime = '日時計算エラー (' + groupTime + ')';
        }
        
        var row = [
            group.name,                        // A列: 研修名
            participantsList,                  // B列: 対象者
            // C列: 講師情報 - 複数講師対応（改行区切り）
            (group.lecturerNames && group.lecturerNames.length > 0) ? 
                group.lecturerNames.join('\n') : 
                (group.lecturer || ''),
            participantCount,                  // D列: 参加者数
            needsRoom ? '必要' : '不要',        // E列: 会議室要否
            roomNameInfo.roomName,              // F列: 会議室名
            group.implementationDay || '',       // G列: 実施日(営業日)
            group.sequence || '',                // H列: 実施順
            trainingDateTime,                  // I列: 研修実施日時
            ''                                 // J列: カレンダーID（初期作成時は空）
        ];
        
        writeLog('DEBUG', '行データ作成完了: ' + group.name + ' (列数: ' + row.length + ')');
        dataRows.push(row);
    }
    
    // データを設定
    writeLog('DEBUG', 'データ行数: ' + dataRows.length);
    if (dataRows.length > 0) {
        writeLog('DEBUG', '最初のデータ行: ' + JSON.stringify(dataRows[0]));
        writeLog('DEBUG', 'データ行の列数: ' + dataRows[0].length);
        writeLog('DEBUG', 'ヘッダーの列数: ' + headers.length);
        
        try {
            mappingSheet.getRange(2, 1, dataRows.length, dataRows[0].length).setValues(dataRows);
            writeLog('INFO', 'データ書き込み成功: ' + dataRows.length + '行');
        } catch (e) {
            writeLog('ERROR', 'データ書き込みエラー: ' + e.message);
            throw e;
        }
        
        // データ部分のスタイル設定
        var dataRange = mappingSheet.getRange(2, 1, dataRows.length, headers.length);
        dataRange.setBorder(true, true, true, true, true, true);
        dataRange.setVerticalAlignment('top');
        
        // 対象者列（B列）の改行設定
        var participantsRange = mappingSheet.getRange(2, 2, dataRows.length, 1);
        participantsRange.setWrap(true);
        participantsRange.setVerticalAlignment('top');
        
        // 講師列（C列）の改行設定 - 複数講師対応
        var lecturerRange = mappingSheet.getRange(2, 3, dataRows.length, 1);
        lecturerRange.setWrap(true);
        lecturerRange.setVerticalAlignment('top');
        
        // 参加者数列（D列）の中央揃え
        var countRange = mappingSheet.getRange(2, 4, dataRows.length, 1);
        countRange.setHorizontalAlignment('center');
        countRange.setVerticalAlignment('middle');
    }
    
    // 列幅の最適化設定（講師列をC列に移動後）
    mappingSheet.setColumnWidth(1, 250);  // A列: 研修名 - 研修名称に適した幅
    mappingSheet.setColumnWidth(2, 300);  // B列: 対象者 - 複数参加者名に対応
    mappingSheet.setColumnWidth(3, 200);  // C列: 講師 - 複数講師名に対応
    mappingSheet.setColumnWidth(4, 80);   // D列: 参加者数 - 数値用の狭い幅
    mappingSheet.setColumnWidth(5, 100);  // E列: 会議室要否 - "必要/不要"用
    mappingSheet.setColumnWidth(6, 150);  // F列: 会議室名 - 会議室名称用
    mappingSheet.setColumnWidth(7, 180);  // G列: 実施日(営業日) - 日付フォーマット用
    mappingSheet.setColumnWidth(8, 100);  // H列: 実施順 - 数値用の狭い幅
    mappingSheet.setColumnWidth(9, 120);  // I列: 研修実施日時 - 日時フォーマット用
    mappingSheet.setColumnWidth(10, 100);  // J列: カレンダーID - システム用（最小限）
    
    // 列の固定（研修名列を固定してスクロールしやすくする）
    mappingSheet.setFrozenColumns(1);
    
    // 行の高さを調整（内容に応じて動的に設定）
    mappingSheet.setRowHeight(1, 60); // ヘッダー行
    for (var i = 2; i <= dataRows.length + 1; i++) {
        var rowIndex = i - 2; // dataRows配列のインデックス
        if (rowIndex < dataRows.length) {
            var participantsList = dataRows[rowIndex][1] || ''; // B列: 対象者
            var lecturerInfo = dataRows[rowIndex][2] || '';     // C列: 講師
            
            // 改行数をカウントして行の高さを決定
            var participantLines = (participantsList.toString().match(/\n/g) || []).length + 1;
            var lecturerLines = (lecturerInfo.toString().match(/\n/g) || []).length + 1;
            var maxLines = Math.max(participantLines, lecturerLines, 1);
            
            // 行の高さを内容に応じて調整（最小30px、改行1つにつき+20px）
            var rowHeight = Math.max(30, 25 + (maxLines - 1) * 20);
            mappingSheet.setRowHeight(i, rowHeight);
        } else {
            mappingSheet.setRowHeight(i, 25); // デフォルトの高さ
        }
    }
    
    // サマリー情報を追加
    var summaryStartRow = dataRows.length + 3;
    mappingSheet.getRange(summaryStartRow, 1).setValue('【サマリー】');
    mappingSheet.getRange(summaryStartRow + 1, 1).setValue('総研修数:');
    mappingSheet.getRange(summaryStartRow + 1, 2).setValue(trainingGroups.length);
    mappingSheet.getRange(summaryStartRow + 2, 1).setValue('総入社者数:');
    mappingSheet.getRange(summaryStartRow + 2, 2).setValue(allNewHires.length);
    mappingSheet.getRange(summaryStartRow + 3, 1).setValue('作成日時:');
    mappingSheet.getRange(summaryStartRow + 3, 2).setValue(new Date());
    
    // サマリー部分のスタイル設定
    var summaryRange = mappingSheet.getRange(summaryStartRow, 1, 4, 2);
    summaryRange.setBackground('#f0f0f0');
    summaryRange.setFontWeight('bold');
    
    writeLog('INFO', 'マッピングシート作成完了: ' + trainingGroups.length + '件の研修, ' + allNewHires.length + '名の入社者');
    
    return mappingSheet;
}

/**
 * スケジュール結果でマッピングシートを更新する
 * @param {Array<Object>} scheduleResults - スケジュール結果の配列
 * @param {Array<Object>} allNewHires - 全入社者の配列
 * @param {Date} periodStart - 研修期間開始日
 * @param {Date} periodEnd - 研修期間終了日
 */
function updateMappingSheetWithScheduleResults(scheduleResults, allNewHires, periodStart, periodEnd) {
    writeLog('INFO', 'スケジュール結果でマッピングシート更新開始');
    writeLog('DEBUG', 'スケジュール結果数: ' + (scheduleResults ? scheduleResults.length : 0));
    
    // スケジュール結果の詳細をログ出力（デバッグ用）
    if (scheduleResults && scheduleResults.length > 0) {
        for (var i = 0; i < Math.min(3, scheduleResults.length); i++) {
            var result = scheduleResults[i];
            writeLog('DEBUG', 'スケジュール結果[' + i + ']: ' + 
                     'training.name=' + (result.training ? result.training.name : 'null') + 
                     ', scheduled=' + result.scheduled + 
                     ', calendarEventId=' + (result.calendarEventId || 'null') + 
                     ', eventTime=' + (result.eventTime ? 'あり' : 'なし'));
        }
    }
    
    var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_IDS.EXECUTION);
    var existingSheets = spreadsheet.getSheets();
    var mappingSheet = null;
    
    // 最新のマッピングシートを探す
    for (var i = 0; i < existingSheets.length; i++) {
        if (existingSheets[i].getName().indexOf('マッピング結果_') === 0) {
            mappingSheet = existingSheets[i];
            break;
        }
    }
    
    if (!mappingSheet) {
        writeLog('WARN', 'マッピングシートが見つかりません。新規作成します。');
        // スケジュール結果を使って新しいマッピングシートを作成
        createMappingSheetWithScheduleResults(scheduleResults, allNewHires, periodStart, periodEnd);
        return;
    }
    
    writeLog('INFO', 'マッピングシート更新: ' + mappingSheet.getName());
    
    // スケジュール結果をトレーニング名でマップ化
    var scheduleMap = {};
    for (var i = 0; i < scheduleResults.length; i++) {
        var result = scheduleResults[i];
        var trainingName = result.training.name;
        scheduleMap[trainingName] = result;
    }
    
    // 既存のデータ行数を取得
    var lastRow = mappingSheet.getLastRow();
    if (lastRow <= 1) {
        writeLog('WARN', 'マッピングシートにデータがありません');
        return;
    }
    
    // データ行を更新（2行目から）
    for (var row = 2; row <= lastRow; row++) {
        var trainingName = mappingSheet.getRange(row, 1).getValue(); // A列: 研修名
        if (!trainingName || trainingName.toString().trim() === '') {
            continue;
        }
        
        var scheduleResult = scheduleMap[trainingName];
        if (scheduleResult) {
            // F列（会議室名）を更新
            var roomName = scheduleResult.roomName || '';
            if (roomName === '' && !scheduleResult.training.needsRoom) {
                roomName = 'オンライン';
            } else if (!scheduleResult.scheduled) {
                roomName = scheduleResult.error || '確保失敗';
            }
            
            mappingSheet.getRange(row, 6).setValue(roomName); // F列: 会議室名
            
            // G列（研修実施日時）を更新（実際のスケジュール時間があれば）
            if (scheduleResult.eventTime && scheduleResult.eventTime.start) {
                var actualDateTime = Utilities.formatDate(scheduleResult.eventTime.start, 'Asia/Tokyo', 'MM/dd(E) HH:mm') + 
                                   '-' + Utilities.formatDate(scheduleResult.eventTime.end, 'Asia/Tokyo', 'HH:mm');
                mappingSheet.getRange(row, 7).setValue(actualDateTime); // G列: 研修実施日時
                writeLog('INFO', 'G列(研修実施日時)更新成功: ' + trainingName + ' → ' + actualDateTime);
            } else {
                writeLog('WARN', 'G列(研修実施日時)更新スキップ: ' + trainingName + ' (eventTime: ' + 
                         (scheduleResult.eventTime ? 'あり' : 'なし') + ', scheduled: ' + scheduleResult.scheduled + ')');
            }
            
            // H列（カレンダーID）を更新
            if (scheduleResult.calendarEventId) {
                // ヘッダー数を確認してカレンダーID列の位置を特定
                var lastCol = mappingSheet.getLastColumn();
                var headerRow = mappingSheet.getRange(1, 1, 1, lastCol).getValues()[0];
                var calendarIdCol = -1;
                for (var col = 0; col < headerRow.length; col++) {
                    if (headerRow[col] === 'カレンダーID') {
                        calendarIdCol = col + 1; // 1ベースのインデックス
                        break;
                    }
                }
                
                if (calendarIdCol > 0) {
                    mappingSheet.getRange(row, calendarIdCol).setValue(scheduleResult.calendarEventId);
                    writeLog('INFO', 'カレンダーID更新成功: ' + trainingName + ' → ' + scheduleResult.calendarEventId + ' (列: ' + calendarIdCol + ')');
                } else {
                    writeLog('ERROR', 'カレンダーID列が見つかりません: ' + trainingName + ' (ヘッダー: ' + headerRow.join(', ') + ')');
                }
            } else {
                writeLog('WARN', 'カレンダーIDが空です: ' + trainingName + ' (scheduled: ' + scheduleResult.scheduled + ', error: ' + (scheduleResult.error || 'なし') + ')');
            }
            
            writeLog('DEBUG', '更新完了: ' + trainingName + ' → 会議室: ' + roomName + ' (F列に設定)');
        } else {
            writeLog('DEBUG', 'スケジュール結果なし: ' + trainingName);
        }
    }
    
    writeLog('INFO', 'マッピングシート更新完了');
}

/**
 * スケジュール結果を使って新しいマッピングシートを作成する
 * @param {Array<Object>} scheduleResults - スケジュール結果の配列
 * @param {Array<Object>} allNewHires - 全入社者の配列
 * @param {Date} periodStart - 研修期間開始日
 * @param {Date} periodEnd - 研修期間終了日
 */
function createMappingSheetWithScheduleResults(scheduleResults, allNewHires, periodStart, periodEnd) {
    writeLog('INFO', 'スケジュール結果でマッピングシート作成開始');
    
    var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_IDS.EXECUTION);
    var sheetName = 'マッピング結果_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd_HH-mm');
    
    // 既存のマッピングシートがあれば削除
    var existingSheets = spreadsheet.getSheets();
    for (var i = 0; i < existingSheets.length; i++) {
        if (existingSheets[i].getName().indexOf('マッピング結果_') === 0) {
            spreadsheet.deleteSheet(existingSheets[i]);
        }
    }
    
    // 新しいシートを作成
    var mappingSheet = spreadsheet.insertSheet(sheetName);
    writeLog('INFO', '新しいシート作成: ' + sheetName);
    
    // ヘッダー行を作成（講師列をB列の次に移動）
    var headers = ['研修名', '対象者', '講師', '参加者数', '会議室要否', '会議室名', '実施日(営業日)', '実施順', '研修実施日時', 'カレンダーID', '処理状況', 'エラー詳細'];
    
    // ヘッダーを設定
    mappingSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // ヘッダー行のスタイル設定
    var headerRange = mappingSheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#4a90e2');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    headerRange.setWrap(true);
    headerRange.setVerticalAlignment('middle');
    
    // データ行を作成
    var dataRows = [];
    for (var i = 0; i < scheduleResults.length; i++) {
        var result = scheduleResults[i];
        var training = result.training;
        
        // 参加者リスト作成（講師を除く）
        var participants = [];
        for (var j = 0; j < allNewHires.length; j++) {
            var hire = allNewHires[j];
            if (training.attendees.indexOf(hire.email) !== -1) {
                participants.push(hire.name + '(' + hire.rank + '/' + hire.experience + '/' + hire.department + ')');
            }
        }
        var participantsList = participants.join('\n');
        // 参加者数（入社者＋講師の合計）
        var lecturerCount = (training.lecturerEmails && training.lecturerEmails.length > 0) ? training.lecturerEmails.length : (training.lecturer ? 1 : 0);
        var participantCount = (participants ? participants.length : 0) + lecturerCount;
        
        // 会議室名を設定
        var roomName = result.roomName || (training.needsRoom ? '会議室未確保' : 'オンライン');
        var trainingDateTime = 'N/A';
        if (result.eventTime && result.eventTime.start) {
            trainingDateTime = Utilities.formatDate(result.eventTime.start, 'Asia/Tokyo', 'MM/dd(E) HH:mm') + 
                             '-' + Utilities.formatDate(result.eventTime.end, 'Asia/Tokyo', 'HH:mm');
        } else {
            trainingDateTime = '日程未定 (' + training.time + ')';
        }
        
        // スケジュール状況
        var scheduleStatus = result.scheduled ? '成功' : ('失敗: ' + (result.error || '不明なエラー'));
        
        var row = [
            training.name,                          // A列: 研修名
            participantsList,                       // B列: 対象者
            // C列: 講師情報 - 複数講師対応（改行区切り）
            (training.lecturerNames && training.lecturerNames.length > 0) ? 
                training.lecturerNames.join('\n') : 
                (training.lecturer || ''),
            participantCount,                       // D列: 参加者数
            training.needsRoom ? '必要' : '不要',    // E列: 会議室要否
            roomName,                               // F列: 会議室名
            training.implementationDay || '',         // G列: 実施日(営業日)
            training.sequence || '',                 // H列: 実施順
            trainingDateTime,                       // I列: 研修実施日時
            result.calendarEventId || '',            // J列: カレンダーID
            scheduleStatus,                         // K列: 処理状況
            result.error || ''                     // L列: エラー詳細
        ];
        
        dataRows.push(row);
        writeLog('DEBUG', '行データ作成: ' + training.name + ' → 会議室: ' + roomName + ', 状況: ' + scheduleStatus);
    }
    
    // データを設定
    if (dataRows.length > 0) {
        mappingSheet.getRange(2, 1, dataRows.length, dataRows[0].length).setValues(dataRows);
        
        // データ部分のスタイル設定
        var dataRange = mappingSheet.getRange(2, 1, dataRows.length, headers.length);
        dataRange.setBorder(true, true, true, true, true, true);
        dataRange.setVerticalAlignment('top');
        
        // 対象者列（B列）の改行設定
        var participantsRange = mappingSheet.getRange(2, 2, dataRows.length, 1);
        participantsRange.setWrap(true);
        participantsRange.setVerticalAlignment('top');
        
        // 講師列（C列）の改行設定 - 複数講師対応
        var lecturerRange = mappingSheet.getRange(2, 3, dataRows.length, 1);
        lecturerRange.setWrap(true);
        lecturerRange.setVerticalAlignment('top');
        
        // 参加者数列（D列）の中央揃え
        var countRange = mappingSheet.getRange(2, 4, dataRows.length, 1);
        countRange.setHorizontalAlignment('center');
        countRange.setVerticalAlignment('middle');
        
        // スケジュール状況列（K列）のスタイル設定
        var statusRange = mappingSheet.getRange(2, 12, dataRows.length, 1);
        for (var i = 0; i < dataRows.length; i++) {
            var status = dataRows[i][11]; // スケジュール状況
            var cellRange = mappingSheet.getRange(i + 2, 12);
            if (status.indexOf('成功') !== -1) {
                cellRange.setBackground('#d4edda');
                cellRange.setFontColor('#155724');
            } else {
                cellRange.setBackground('#f8d7da');
                cellRange.setFontColor('#721c24');
            }
        }
    }
    
    // 列幅の最適化設定（スケジュール結果版 - 12列構成、講師列をC列に移動後）
    mappingSheet.setColumnWidth(1, 250);  // A列: 研修名 - 研修名称に適した幅
    mappingSheet.setColumnWidth(2, 300);  // B列: 対象者 - 複数参加者名に対応
    mappingSheet.setColumnWidth(3, 200);  // C列: 講師 - 複数講師名に対応
    mappingSheet.setColumnWidth(4, 80);   // D列: 参加者数 - 数値用の狭い幅
    mappingSheet.setColumnWidth(5, 100);  // E列: 会議室要否 - "必要/不要"用
    mappingSheet.setColumnWidth(6, 150);  // F列: 会議室名 - 会議室名称用
    mappingSheet.setColumnWidth(7, 180);  // G列: 実施日(営業日) - 日付フォーマット用
    mappingSheet.setColumnWidth(8, 100);  // H列: 実施順 - 数値用の狭い幅
    mappingSheet.setColumnWidth(9, 120);  // I列: 研修実施日時 - 日時フォーマット用
    mappingSheet.setColumnWidth(10, 100);  // J列: カレンダーID - システム用（最小限）
    mappingSheet.setColumnWidth(11, 120);  // K列: 処理状況 - 成功/失敗状況
    mappingSheet.setColumnWidth(12, 300);  // L列: エラー詳細 - エラー理由
    
    // 列の固定（研修名列を固定してスクロールしやすくする）
    mappingSheet.setFrozenColumns(1);
    
    // 行の高さを調整（内容に応じて動的に設定）
    mappingSheet.setRowHeight(1, 60); // ヘッダー行
    for (var i = 2; i <= dataRows.length + 1; i++) {
        var rowIndex = i - 2; // dataRows配列のインデックス
        if (rowIndex < dataRows.length) {
            var participantsList = dataRows[rowIndex][1] || ''; // B列: 対象者
            var lecturerInfo = dataRows[rowIndex][2] || '';     // C列: 講師
            
            // 改行数をカウントして行の高さを決定
            var participantLines = (participantsList.toString().match(/\n/g) || []).length + 1;
            var lecturerLines = (lecturerInfo.toString().match(/\n/g) || []).length + 1;
            var maxLines = Math.max(participantLines, lecturerLines, 1);
            
            // 行の高さを内容に応じて調整（最小30px、改行1つにつき+20px）
            var rowHeight = Math.max(30, 25 + (maxLines - 1) * 20);
            mappingSheet.setRowHeight(i, rowHeight);
        } else {
            mappingSheet.setRowHeight(i, 25); // デフォルトの高さ
        }
    }
    
    // サマリー情報を追加
    var summaryStartRow = dataRows.length + 3;
    mappingSheet.getRange(summaryStartRow, 1).setValue('【サマリー】');
    mappingSheet.getRange(summaryStartRow + 1, 1).setValue('総研修数:');
    mappingSheet.getRange(summaryStartRow + 1, 2).setValue(scheduleResults.length);
    mappingSheet.getRange(summaryStartRow + 2, 1).setValue('成功した研修数:');
    
    var successCount = 0;
    for (var i = 0; i < scheduleResults.length; i++) {
        if (scheduleResults[i].scheduled) successCount++;
    }
    mappingSheet.getRange(summaryStartRow + 2, 2).setValue(successCount);
    
    mappingSheet.getRange(summaryStartRow + 3, 1).setValue('総入社者数:');
    mappingSheet.getRange(summaryStartRow + 3, 2).setValue(allNewHires.length);
    mappingSheet.getRange(summaryStartRow + 4, 1).setValue('作成日時:');
    mappingSheet.getRange(summaryStartRow + 4, 2).setValue(new Date());
    
    // サマリー部分のスタイル設定
    var summaryRange = mappingSheet.getRange(summaryStartRow, 1, 5, 2);
    summaryRange.setBackground('#f0f0f0');
    summaryRange.setFontWeight('bold');
    
    writeLog('INFO', 'スケジュール結果マッピングシート作成完了: ' + scheduleResults.length + '件の研修, 成功: ' + successCount + '件');
    
    return mappingSheet;
}

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
            if (updates.status.indexOf('成功') !== -1) {
                statusCell.setBackground('#d4edda').setFontColor('#155724');
            } else if (updates.status.indexOf('失敗') !== -1) {
                statusCell.setBackground('#f8d7da').setFontColor('#721c24');
            } else if (updates.status.indexOf('スキップ') !== -1) {
                statusCell.setBackground('#e2e3e5').setFontColor('#383d41');
            }
        }
        if (updates.errorReason !== undefined) {
            mappingSheet.getRange(rowIndex, 10).setValue(updates.errorReason); // J列: エラー詳細
        }
        
        // 表示を強制的に更新
        SpreadsheetApp.flush();
    } catch (e) {
        writeLog('ERROR', 'マッピングシート行更新エラー: ' + e.message);
    }
} 