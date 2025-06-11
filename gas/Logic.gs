// =========================================
// ビジネスロジック関連
// =========================================

/**
 * 職位と経験有無から対応する研修パターンを決定する
 * @param {string} rank - 職位
 * @param {string} experience - 経験有無
 * @returns {string} 研修パターン（A, B, C, D）
 */
function determineTrainingPattern(rank, experience) {
  // 経験有無が空白の場合は職位のみで判定
  if (!experience || experience.trim() === '') {
    // 職位のみでの判定
    if (rank === 'M') {
      return 'C'; // M向け
    }
    if (rank === 'SM' || rank === 'D') {
      return 'D'; // SMup向け
    }
    // S, C, CSで経験有無が空白の場合はパターン対象外
    writeLog('WARN', '職位 ' + rank + ' で経験有無が未入力のため、研修対象外です');
    return null;
  }
  
  // Aパターン（未経験者向け）：職位がS,C,CSで経験有無が「未経験者」
  if ((rank === 'S' || rank === 'C' || rank === 'CS') && experience === '未経験者') {
    return 'A';
  }
  
  // Bパターン（経験C/SC向け）：職位がC,CSで経験有無が「経験者」
  if ((rank === 'C' || rank === 'CS') && experience === '経験者') {
    return 'B';
  }
  
  // Cパターン（M向け）：職位がM
  if (rank === 'M') {
    return 'C';
  }
  
  // Dパターン（SMup向け）：職位がSM,D
  if (rank === 'SM' || rank === 'D') {
    return 'D';
  }
  
  // デフォルト（該当なし）
  writeLog('WARN', '職位と経験の組み合わせが研修パターンに該当しません: 職位=' + rank + ', 経験=' + experience);
  return null;
}

/**
 * 入社者データの必須項目を検証する
 * @param {Array<Object>} newHires - 入社者情報の配列
 */
function validateNewHires(newHires) {
  var errors = [];
  for (var i = 0; i < newHires.length; i++) {
    var hire = newHires[i];
    if (!hire.rank || !hire.email) {
      errors.push('行 ' + hire.rowNum + ': ' + hire.name + ' の職位またはメールアドレスが未入力です。');
    }
    // 経験有無は任意項目とする（空白の場合はスキップまたはデフォルト処理）
  }
  if (errors.length > 0) {
    throw new Error('データ不備エラー:\n' + errors.join('\n'));
  }
}

/**
 * 入社者と研修マスタを基に、研修ごとのグループを作成する
 * @param {Array<Object>} newHires - 入社者情報の配列
 * @returns {Array<Object>} 研修グループの配列
 */
function groupTrainingsForHires(newHires) {
  var masterSheet = SpreadsheetApp.openById(SPREADSHEET_IDS.TRAINING_MASTER).getSheetByName(SHEET_NAMES.TRAINING_MASTER);
  var lastRow = masterSheet.getLastRow();
  writeLog('DEBUG', '研修マスタの最終行: ' + lastRow);
  
  // 実際のヘッダー構造を確認（デバッグ用）
  if (lastRow >= 4) {
    var actualHeader = masterSheet.getRange(4, 1, 1, 19).getValues()[0];
    writeLog('DEBUG', '4行目ヘッダー（A-S列）: ' + actualHeader.join(' | '));
  }
  
  if (lastRow <= 4) {
    writeLog('DEBUG', '研修マスタにデータが存在しません（5行目以降にデータなし）');
    return [];
  }
  
  var masterData = masterSheet.getRange(5, 1, lastRow - 4, 19).getValues(); // 5行目から19列取得（A-S列）
  writeLog('DEBUG', '研修マスタの取得データ行数: ' + masterData.length);

  var trainingGroups = {};

  // 最初の3行のデータを詳細出力（デバッグ用）- 全列表示
  for (var debugIdx = 0; debugIdx < Math.min(3, masterData.length); debugIdx++) {
    writeLog('DEBUG', '研修データ行' + (debugIdx + 5) + ' 全列: ' + masterData[debugIdx].join(' | '));
  }

  for (var i = 0; i < masterData.length; i++) {
    var row = masterData[i];
    
    // A-S列（19列）の正しい構成
    // A列:Lv.1 | B列:Lv.2 | C列:研修名称 | D列:Aパターン（未経験者向け） | E列:Bパターン（経験C/SC向け） | F列:Cパターン（M向け） | G列:Dパターン（SMup向け） | H列:1日 | I列:15日 | J列:実施日（n営業日）/コンテンツ実施順 | K列:時間（単位：分） | L列:時間備考 | M列:担当者 | N列:メールアドレス用略称 | O列:メールアドレス | P列:会議室要否 | Q列:カレンダーメモ | R列:備考 | S列:（空白？）
    
    var lv1 = row[0];           // A列: Lv.1
    var lv2 = row[1];           // B列: Lv.2
    var trainingName = row[2];  // C列: 研修名称
    var patternA = row[3];      // D列: Aパターン（未経験者向け）
    var patternB = row[4];      // E列: Bパターン（経験C/SC向け）
    var patternC = row[5];      // F列: Cパターン（M向け）
    var patternD = row[6];      // G列: Dパターン（SMup向け）
    var timing1 = row[7];       // H列: 1日
    var timing15 = row[8];      // I列: 15日
    var sequence = row[9];      // J列: 実施日（n営業日）/コンテンツ実施順
    var timeMinutes = row[10];  // K列: 時間（単位：分）
    var timeNote = row[11];     // L列: 時間備考
    var instructor = row[12];   // M列: 担当者
    var emailShort = row[13];   // N列: メールアドレス用略称
    var email = row[14];        // O列: メールアドレス
    var needsRoom = row[15];    // P列: 会議室要否
    var memo = row[16];         // Q列: カレンダーメモ
    var note = row[17];         // R列: 備考
    var extraCol = row[18];     // S列: （追加列）
    
    // 最初の5行のみ詳細ログを出力
    if (i < 5) {
      writeLog('DEBUG', '研修行' + (i + 5) + ': A列(Lv1)=' + lv1 + ', B列(Lv2)=' + lv2 + ', C列(研修名)=' + trainingName);
      writeLog('DEBUG', '  対象職位: D列(A)=' + patternA + ', E列(B)=' + patternB + ', F列(C)=' + patternC + ', G列(D)=' + patternD);
      writeLog('DEBUG', '  その他: M列(担当者)=' + instructor + ', P列(会議室)=' + needsRoom);
    } else {
      writeLog('DEBUG', '研修行' + (i + 5) + ': 研修名=' + trainingName + ', 担当者=' + instructor);
    }
    
    // 研修名が空の場合はスキップ
    if (!trainingName) {
      continue;
    }
    
    // 対象パターンを確認（●マークがある列をチェック）- Lv.1,Lv.2は使用しない
    var targetPatterns = [];
    if (patternA === '●') targetPatterns.push('A'); // 未経験者向け
    if (patternB === '●') targetPatterns.push('B'); // 経験C/SC向け
    if (patternC === '●') targetPatterns.push('C'); // M向け
    if (patternD === '●') targetPatterns.push('D'); // SMup向け
    
    writeLog('DEBUG', '【新コード実行中】対象パターン: ' + targetPatterns.join(', '));
    
    for (var j = 0; j < newHires.length; j++) {
      var hire = newHires[j];
      var hirePattern = determineTrainingPattern(hire.rank, hire.experience);
      writeLog('DEBUG', '入社者: ' + hire.name + ', 職位: ' + hire.rank + ', 経験: ' + hire.experience + ', パターン: ' + hirePattern);
      writeLog('DEBUG', '対象パターンに含まれるか: ' + (hirePattern && targetPatterns.indexOf(hirePattern) !== -1));
      
      if (hirePattern && targetPatterns.indexOf(hirePattern) !== -1) {
        if (!trainingGroups[trainingName]) {
          // 講師のメールアドレスを決定
          var lecturerEmail = email || (emailShort ? emailShort + '@flux-g.com' : '');
          
          trainingGroups[trainingName] = {
            name: trainingName,
            sequence: sequence || 999, // 順番（未設定の場合は999）
            time: timeMinutes ? timeMinutes + '分' : '60分', // 時間
            lecturer: lecturerEmail,
            needsRoom: needsRoom === '必要',
            memo: memo || '',
            attendees: lecturerEmail ? [lecturerEmail] : [], // 講師を追加
          };
          writeLog('DEBUG', '新しい研修グループ作成: ' + trainingName + ' (講師: ' + lecturerEmail + ')');
        }
        if (trainingGroups[trainingName].attendees.indexOf(hire.email) === -1) {
            trainingGroups[trainingName].attendees.push(hire.email);
            writeLog('DEBUG', '参加者追加: ' + hire.name + '(' + hire.rank + '/' + hire.experience + ') -> ' + trainingName);
            writeLog('DEBUG', '現在の参加者リスト: ' + trainingGroups[trainingName].attendees.join(', '));
        }
      }
    }
  }
  
  // オブジェクトから配列に変換し、セット順番でソート
  var groups = [];
  for (var key in trainingGroups) {
    groups.push(trainingGroups[key]);
  }
  writeLog('INFO', '作成された研修グループ数: ' + groups.length);
  return groups.sort(function(a, b) { return a.sequence - b.sequence; });
} 