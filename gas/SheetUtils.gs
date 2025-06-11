// =========================================
// スプレッドシート関連ユーティリティ
// =========================================

/**
 * `入社者リスト`から未処理のデータを取得する
 * @returns {Array<Object>} 入社者情報の配列
 */
function getNewHires() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_IDS.NEW_HIRES).getSheetByName(SHEET_NAMES.NEW_HIRES);
  var lastRow = sheet.getLastRow();
  writeLog('DEBUG', '入社者リストの最終行: ' + lastRow);
  
  if (lastRow <= 1) {
    writeLog('DEBUG', 'データが存在しません');
    return [];
  }
  
  var data = sheet.getRange(2, 1, lastRow - 1, 14).getValues(); // 14列取得
  writeLog('DEBUG', '取得したデータ行数: ' + data.length);
  
  var newHires = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var rank = row[0];           // A列: 職位
    var experience = row[1];     // B列: 経験有無
    var name = row[2];           // C列: 名前
    var emailShort = row[3];     // D列: メールアドレス用略称
    var email = row[4];          // E列: メールアドレス
    
    writeLog('DEBUG', '行' + (i + 2) + ': 職位=' + rank + ', 経験=' + experience + ', 氏名=' + name + ', 略称=' + emailShort + ', メール=' + email);
    
    // 職位、氏名、メールアドレスまたは略称が入力されている行を対象とする
    if (rank && name && (email || emailShort)) {
      // メールアドレスが空の場合は略称@flux-g.comを使用
      var finalEmail = email || (emailShort + '@flux-g.com');
      
      newHires.push({
        rowNum: i + 2, // 実際の行番号
        rank: rank,     // A列: 職位
        experience: experience, // B列: 経験有無
        name: name,     // C列: 名前
        email: finalEmail, // E列またはD列+ドメイン
      });
    }
  }
  writeLog('DEBUG', '対象入社者数: ' + newHires.length);
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
 * @param {Object} params - 実行パラメータ
 * @param {string} status - 処理結果 (成功/エラー)
 * @param {string} message - 詳細メッセージ
 */
function logExecution(params, status, message) {
    // E,F列（5,6列目）にログを記載
    params.sheet.getRange(2, 5).setValue(status);   // E列2行目に状態
    params.sheet.getRange(2, 6).setValue(message);  // F列2行目にメッセージ
}

/**
 * 研修と入社者のマッピング結果をシートに表示する
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
    
    // ヘッダー行を作成
    var headers = ['研修名', '対象者', '参加者数', '会議室要否', '会議室名', '講師', '研修実施日時'];
    
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
    var dataRows = [];
    for (var i = 0; i < trainingGroups.length; i++) {
        var group = trainingGroups[i];
        writeLog('DEBUG', '研修グループ処理中: ' + group.name);
        
        // 参加者リスト作成（講師を除く）
        var participants = [];
        for (var j = 0; j < allNewHires.length; j++) {
            var hire = allNewHires[j];
            if (group.attendees.indexOf(hire.email) !== -1) {
                participants.push(hire.name + '(' + hire.rank + '/' + hire.experience + ')');
            }
        }
        var participantsList = participants.join('\n');
        
        // 参加者数（講師を除く）
        var participantCount = participants.length;
        
        // 会議室名を取得（仮）
        var roomName = '';
        if (group.needsRoom && participantCount > 0) {
            try {
                roomName = findAvailableRoomName(group.attendees.length);
            } catch (e) {
                roomName = 'エラー: ' + e.message;
            }
        }
        
        // 研修実施日時を計算
        var trainingDateTime = '';
        if (periodStart && periodEnd) {
            var eventDate = new Date(periodStart);
            var sequence = group.sequence || 1;
            
            // 平日ベースで日付を計算
            var businessDaysToAdd = sequence - 1;
            var daysAdded = 0;
            while (daysAdded < businessDaysToAdd) {
                eventDate.setDate(eventDate.getDate() + 1);
                if (eventDate.getDay() !== 0 && eventDate.getDay() !== 6) {
                    daysAdded++;
                }
            }
            
            // 期間内チェック
            if (eventDate > periodEnd) {
                eventDate = new Date(periodEnd);
            }
            
            var timeMatch = group.time.match(/(\d+)分/);
            var durationMinutes = timeMatch ? parseInt(timeMatch[1]) : 60;
            var endTime = new Date(new Date(eventDate).getTime() + (durationMinutes * 60 * 1000));
            endTime.setHours(9 + Math.floor(durationMinutes / 60), durationMinutes % 60);
            
            trainingDateTime = Utilities.formatDate(eventDate, 'Asia/Tokyo', 'MM/dd(E)') + 
                             ' 9:00-' + 
                             Utilities.formatDate(endTime, 'Asia/Tokyo', 'HH:mm') + 
                             ' (' + group.time + ')';
        } else {
            trainingDateTime = '日程未定 (' + group.time + ')';
        }
        
        var row = [
            group.name,           // 研修名
            participantsList,     // 対象者
            participantCount,     // 参加者数
            group.needsRoom ? '必要' : '不要', // 会議室要否
            roomName,             // 会議室名
            group.lecturer || '', // 講師
            trainingDateTime      // 研修実施日時
        ];
        
        dataRows.push(row);
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
        
        // 参加者数列（C列）の中央揃え
        var countRange = mappingSheet.getRange(2, 3, dataRows.length, 1);
        countRange.setHorizontalAlignment('center');
        countRange.setVerticalAlignment('middle');
    }
    
    // 列幅の自動調整
    for (var i = 1; i <= headers.length; i++) {
        mappingSheet.autoResizeColumn(i);
    }
    
    // 行の高さを調整
    mappingSheet.setRowHeight(1, 60); // ヘッダー行
    for (var i = 2; i <= dataRows.length + 1; i++) {
        mappingSheet.setRowHeight(i, 25); // データ行
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