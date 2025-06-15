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
  writeLog('DEBUG', 'パターン判定開始: 職位=' + rank + ', 経験=' + experience);
  
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
  
  // 職位Sの場合は常にパターンC
  if (rank === 'S') {
    writeLog('DEBUG', 'Cパターンに分類（職位S専用）: 職位=' + rank + ', 経験=' + experience);
    return 'C';
  }
  
  // Aパターン（未経験者向け）：職位がC,SCで経験有無が「未経験者」
  if ((rank === 'C' || rank === 'SC') && experience === '未経験者') {
    writeLog('DEBUG', 'Aパターンに分類: 職位=' + rank + ', 経験=' + experience);
    return 'A';
  }
  
  // Bパターン（経験C/SC向け）：職位がC,SCで経験有無が「経験者」
  if ((rank === 'C' || rank === 'SC') && experience === '経験者') {
    writeLog('DEBUG', 'Bパターンに分類: 職位=' + rank + ', 経験=' + experience);
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
  writeLog('DEBUG', 'パターン判定結果: null (該当パターンなし)');
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
    if (!hire.rank || hire.rank.trim() === '') {
      errors.push('行 ' + hire.rowNum + ': ' + hire.name + ' の職位が未入力です。');
    }
    if (!hire.email || hire.email.trim() === '') {
      errors.push('行 ' + hire.rowNum + ': ' + hire.name + ' のメールアドレスが未入力です。');
    }
    // 経験有無は任意項目とする（空白の場合はスキップまたはデフォルト処理）
    writeLog('DEBUG', '検証完了: ' + hire.name + ' (職位:' + hire.rank + ', 経験:' + hire.experience + ', 所属:' + hire.department + ')');
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
  var lastCol = masterSheet.getLastColumn(); // 列数を先に取得
  writeLog('DEBUG', '研修マスタの最終行: ' + lastRow + ', 最終列: ' + lastCol);
  
  // 実際のヘッダー構造を確認（デバッグ用）M,P列削除後
  if (lastRow >= 4 && lastCol > 0) {
    var headerCols = Math.max(1, Math.min(lastCol, 19)); // 最低1列、最大19列
    var actualHeader = masterSheet.getRange(4, 1, 1, headerCols).getValues()[0];
    writeLog('DEBUG', '4行目ヘッダー（M,P列削除後、列数:' + headerCols + '）: ' + actualHeader.join(' | '));
  } else {
    writeLog('WARN', 'ヘッダー取得スキップ: lastRow=' + lastRow + ', lastCol=' + lastCol);
  }
  
  if (lastRow <= 4) {
    writeLog('DEBUG', '研修マスタにデータが存在しません（5行目以降にデータなし）');
    return [];
  }
  
  var maxCol = Math.max(1, Math.min(lastCol, 19)); // M,P列削除後は最大19列まで取得（最低1列）
  
  var masterData = masterSheet.getRange(5, 1, lastRow - 4, maxCol).getValues();
  writeLog('DEBUG', '研修マスタの取得データ行数: ' + masterData.length + ', 取得列数: ' + maxCol);

  var trainingGroups = {};

  // 最初の3行のデータを詳細出力（デバッグ用）- 全列表示
  for (var debugIdx = 0; debugIdx < Math.min(3, masterData.length); debugIdx++) {
    writeLog('DEBUG', '研修データ行' + (debugIdx + 5) + ' 全列: ' + masterData[debugIdx].join(' | '));
  }

  for (var i = 0; i < masterData.length; i++) {
    var row = masterData[i];
    
    // M列とP列（講師メールアドレス用略称）削除後の構成
    var lv1 = row[0];           // A列: Lv.1
    var lv2 = row[1];           // B列: Lv.2
    var trainingName = row[2];  // C列: 研修名称
    var patternA = row[3];      // D列: Aパターン（未経験者向け）
    var patternB = row[4];      // E列: Bパターン（経験C/SC向け）
    var patternC = row[5];      // F列: Cパターン（M向け）
    var patternD = row[6];      // G列: Dパターン（SMup向け）
    var implementationDay = row[7];  // H列: 実施日（n営業日）
    var sequence = row[8];      // I列: コンテンツ実施順
    var timeMinutes = row[9];   // J列: 時間（単位：分）
    var timeNote = row[10];     // K列: 時間備考
    var instructor1 = row[11];  // L列: 講師1:担当者
    var email1 = row[12];       // M列: 講師1:メールアドレス（元N列）
    var instructor2 = row.length > 13 ? row[13] : null;  // N列: 講師2:担当者（元O列）
    var email2 = row.length > 14 ? row[14] : null;       // O列: 講師2:メールアドレス（元Q列）
    
    // 会議室要否の判定（P列に移動）
    var needsRoom = row[15];    // P列: 会議室要否（元R列）
    
    var memo = row[16];         // Q列: カレンダーメモ（元S列）
    var note = row[17];         // R列: 備考（元T列）
    var extraCol = row[18];     // S列: （追加列、元U列）
    
    // 研修名が空の場合はスキップ
    if (!trainingName || trainingName.trim() === '') {
      writeLog('DEBUG', '研修行' + (i + 5) + ': 研修名が空のためスキップ');
      continue;
    }
    
    // A列（Lv.1）のフィルター：「DX ONB」または「ビジネススキル研修」以外はスキップ
    if (lv1 !== 'DX ONB' && lv1 !== 'ビジネススキル研修') {
      writeLog('DEBUG', '研修行' + (i + 5) + ': A列(' + lv1 + ')が対象外のためスキップ - ' + trainingName);
      continue;
    }
    
    // 研修項目ごとのループ処理開始ログ
    writeLog('INFO', '=== 研修項目処理開始 ===');
    writeLog('INFO', '研修名: ' + trainingName);
    writeLog('INFO', '行番号: ' + (i + 5));
    writeLog('INFO', '担当者1: ' + instructor1 + ', 担当者2: ' + instructor2);
    writeLog('INFO', '時間: ' + timeMinutes + '分');
    writeLog('INFO', '会議室要否: "' + needsRoom + '" (型: ' + typeof needsRoom + ', 判定結果: ' + (needsRoom === '必要') + ')');
    writeLog('INFO', '実施日（営業日）: ' + implementationDay);
    writeLog('INFO', 'コンテンツ実施順: ' + sequence);
    
    // 列データの詳細デバッグ（M,P列削除後の構成）
    writeLog('DEBUG', '行データ長: ' + row.length + '列');
    writeLog('DEBUG', '列データ詳細: M列(講師1Email)=' + (row[12] || 'undefined') + ', N列(講師2名)=' + (row[13] || 'undefined') + ', O列(講師2Email)=' + (row[14] || 'undefined') + ', P列(会議室要否)=' + (row[15] || 'undefined') + ', Q列(メモ)=' + (row[16] || 'undefined'));
    
    // 対象パターンを確認（●マークがある列をチェック）
    var targetPatterns = [];
    if (patternA === '●') targetPatterns.push('A'); // 未経験者向け
    if (patternB === '●') targetPatterns.push('B'); // 経験C/SC向け
    if (patternC === '●') targetPatterns.push('C'); // M向け
    if (patternD === '●') targetPatterns.push('D'); // SMup向け
    
    writeLog('INFO', '対象パターン: [' + targetPatterns.join(', ') + ']');
    
    if (targetPatterns.length === 0) {
      writeLog('WARN', '対象パターンが設定されていません。この研修はスキップされます。');
      continue;
    }
    
    // 各入社者との照合
    var matchedParticipants = [];
    for (var j = 0; j < newHires.length; j++) {
      var hire = newHires[j];
      var hirePattern = determineTrainingPattern(hire.rank, hire.experience);
      
      if (hirePattern && targetPatterns.indexOf(hirePattern) !== -1) {
        matchedParticipants.push(hire);
        writeLog('INFO', '  ✓ 参加者: ' + hire.name + ' (職位:' + hire.rank + ', 経験:' + hire.experience + ', 所属:' + hire.department + ', パターン:' + hirePattern + ')');
      } else {
        writeLog('DEBUG', '  ✗ 非対象: ' + hire.name + ' (職位:' + hire.rank + ', 経験:' + hire.experience + ', 所属:' + hire.department + ', パターン:' + hirePattern + ')');
      }
    }
    
    if (matchedParticipants.length > 0) {
      // 研修グループ作成 - 複数講師対応
      var lecturerEmails = [];
      var lecturerNames = [];
      
      // 講師1の情報処理（M,P列削除により略称列は使用しない）
      if (instructor1 && instructor1.trim() !== '') {
        var lecturerEmail1 = email1 || '';
        if (lecturerEmail1 && lecturerEmail1.trim() !== '') {
          lecturerEmails.push(lecturerEmail1.trim());
          lecturerNames.push(instructor1.trim());
          writeLog('DEBUG', '講師1追加: ' + instructor1 + ' (' + lecturerEmail1 + ')');
        } else {
          writeLog('WARN', '講師1のメールアドレスが設定されていません: ' + instructor1);
        }
      }
      
      // 講師2の情報処理（M,P列削除により略称列は使用しない）
      if (instructor2 && instructor2.trim() !== '') {
        var lecturerEmail2 = email2 || '';
        if (lecturerEmail2 && lecturerEmail2.trim() !== '') {
          lecturerEmails.push(lecturerEmail2.trim());
          lecturerNames.push(instructor2.trim());
          writeLog('DEBUG', '講師2追加: ' + instructor2 + ' (' + lecturerEmail2 + ')');
        } else {
          writeLog('WARN', '講師2のメールアドレスが設定されていません: ' + instructor2);
        }
      }
      
      // 後方互換性のため、lecturerEmailも維持
      var lecturerEmail = lecturerEmails.length > 0 ? lecturerEmails[0] : '';
      
      // 実施日を数値に正規化（簡素化版）
      var normalizedImplementationDay = 50; // デフォルト値を50に変更（999より小さく、通常の実施日より大きい値）
      writeLog('DEBUG', '実施日正規化開始: ' + trainingName + ', 元値=' + implementationDay + ', 型=' + typeof implementationDay);
      
      if (implementationDay) {
        if (typeof implementationDay === 'number' && !isNaN(implementationDay)) {
          normalizedImplementationDay = Math.max(1, Math.floor(implementationDay)); // 1以上の整数に正規化
          writeLog('DEBUG', '数値型実施日: ' + implementationDay + ' → ' + normalizedImplementationDay);
        } else {
          var dayStr = implementationDay.toString().trim();
          // 文字列から数値を抽出（例: "1営業日" → 1, "3" → 3, "第2日目" → 2）
          var dayMatch = dayStr.match(/(\d+)/); // 先頭でなくても良い
          if (dayMatch) {
            normalizedImplementationDay = Math.max(1, parseInt(dayMatch[1]));
            writeLog('DEBUG', '文字列型実施日: "' + dayStr + '" → ' + normalizedImplementationDay);
          } else {
            writeLog('WARN', '実施日のパターンが認識できませんでした: "' + dayStr + '" → デフォルト値(' + normalizedImplementationDay + ')を使用');
          }
        }
      } else {
        writeLog('DEBUG', '実施日が空のため、デフォルト値(' + normalizedImplementationDay + ')を使用');
      }
      
      writeLog('DEBUG', '実施日正規化完了: ' + trainingName + ' → ' + normalizedImplementationDay);
      
      // コンテンツ実施順を数値に正規化（簡素化版）
      var normalizedSequence = 999; // デフォルト値
      writeLog('DEBUG', '実施順正規化開始: ' + trainingName + ', 元値=' + sequence + ', 型=' + typeof sequence);
      
      if (sequence) {
        if (typeof sequence === 'number' && !isNaN(sequence)) {
          normalizedSequence = Math.max(1, Math.floor(sequence)); // 1以上の整数に正規化
          writeLog('DEBUG', '数値型実施順: ' + sequence + ' → ' + normalizedSequence);
        } else {
          var seqStr = sequence.toString().trim();
          writeLog('DEBUG', '文字列型実施順処理: "' + seqStr + '"');
          
          var found = false;
          
          // パターン1: 先頭の数字（例: "1", "3番目", "2日目"）- 最も一般的
          var numMatch = seqStr.match(/^(\d+)/);
          if (numMatch) {
            normalizedSequence = Math.max(1, parseInt(numMatch[1]));
            writeLog('DEBUG', '先頭数字マッチ: "' + seqStr + '" → ' + normalizedSequence);
            found = true;
          }
          
          // パターン2: 丸数字（例: "①", "②"）- よく使われる
          if (!found) {
            var circleNumbers = ['①', '②', '③', '④', '⑤', '⑥', '⑦', '⑧', '⑨', '⑩'];
            for (var ci = 0; ci < circleNumbers.length; ci++) {
              if (seqStr.indexOf(circleNumbers[ci]) !== -1) {
                normalizedSequence = ci + 1;
                writeLog('DEBUG', '丸数字マッチ: "' + seqStr + '" (' + circleNumbers[ci] + ') → ' + normalizedSequence);
                found = true;
                break;
              }
            }
          }
          
          // パターン3: カッコ付き数字（例: "(1)", "（2）"）
          if (!found) {
            var bracketMatch = seqStr.match(/[（(](\d+)[）)]/);
            if (bracketMatch) {
              normalizedSequence = Math.max(1, parseInt(bracketMatch[1]));
              writeLog('DEBUG', 'カッコ付き数字マッチ: "' + seqStr + '" → ' + normalizedSequence);
              found = true;
            }
          }
          
          // パターン4: 任意の位置の数字（最後の手段）
          if (!found) {
            var anyNumMatch = seqStr.match(/(\d+)/);
            if (anyNumMatch) {
              normalizedSequence = Math.max(1, parseInt(anyNumMatch[1]));
              writeLog('DEBUG', '任意位置数字マッチ: "' + seqStr + '" → ' + normalizedSequence);
              found = true;
            }
          }
          
          if (!found) {
            writeLog('WARN', '実施順のパターンが認識できませんでした: "' + seqStr + '" → デフォルト値(' + normalizedSequence + ')を使用');
          }
        }
      } else {
        writeLog('DEBUG', '実施順が空のため、デフォルト値(' + normalizedSequence + ')を使用');
      }
      
      writeLog('DEBUG', '実施順正規化完了: ' + trainingName + ' → ' + normalizedSequence);
      
      // ユニークキーは正規化後のデータで生成（複数講師対応）
      var lecturerKeyPart = lecturerEmails.length > 0 ? lecturerEmails.join(',') : '';
      var uniqueKey = trainingName + '_day' + normalizedImplementationDay + '_seq' + normalizedSequence + '_' + lecturerKeyPart;
      
      if (!trainingGroups[uniqueKey]) {
        
        trainingGroups[uniqueKey] = {
          name: trainingName,
          implementationDay: normalizedImplementationDay,
          implementationDayRaw: implementationDay, // 元の値も保持（デバッグ用）
          sequence: normalizedSequence,
          sequenceRaw: sequence, // 元の値も保持（デバッグ用）
          time: timeMinutes ? timeMinutes + '分' : '60分',
          lecturer: lecturerEmail, // 後方互換性のため第1講師
          lecturerEmails: lecturerEmails, // 全講師のメールアドレス配列
          lecturerNames: lecturerNames, // 全講師の名前配列
          needsRoom: needsRoom === '必要',
          memo: memo || '',
          attendees: lecturerEmails.slice(), // 講師全員を参加者に追加
          uniqueKey: uniqueKey,
        };
        
        writeLog('DEBUG', '実施日正規化: ' + implementationDay + ' → ' + normalizedImplementationDay);
        writeLog('DEBUG', '実施順正規化: ' + sequence + ' → ' + normalizedSequence);
        writeLog('DEBUG', 'ユニークキー生成: ' + uniqueKey);
        var lecturerInfo = lecturerEmails.length > 1 ? 
                          '講師: ' + lecturerNames.join(', ') + ' (' + lecturerEmails.length + '名)' :
                          '講師: ' + (lecturerNames.length > 0 ? lecturerNames[0] : '未設定');
        writeLog('INFO', '研修グループ作成: ' + trainingName + ' (ユニークキー: ' + uniqueKey + ', ' + lecturerInfo + ')');
      } else {
        writeLog('WARN', '既存の研修グループにマージ: ' + uniqueKey);
      }
      
      // 参加者追加
      for (var k = 0; k < matchedParticipants.length; k++) {
        var participant = matchedParticipants[k];
        if (trainingGroups[uniqueKey].attendees.indexOf(participant.email) === -1) {
          trainingGroups[uniqueKey].attendees.push(participant.email);
        }
      }
      
      writeLog('INFO', '参加者数: ' + matchedParticipants.length + '名');
      writeLog('INFO', '最終参加者リスト: ' + trainingGroups[uniqueKey].attendees.join(', '));
    } else {
      writeLog('INFO', '参加者なし: この研修に該当する入社者はいません');
    }
    
    writeLog('INFO', '=== 研修項目処理終了 ===');
    writeLog('INFO', ''); // 空行で区切り
  }
  
  // オブジェクトから配列に変換し、実施日とコンテンツ実施順でソート
  var groups = [];
  for (var key in trainingGroups) {
    groups.push(trainingGroups[key]);
  }
  writeLog('INFO', '作成された研修グループ数: ' + groups.length);
  
  // ソート前の状態をログ出力
  writeLog('INFO', '=== ソート前の研修順序 ===');
  for (var i = 0; i < groups.length; i++) {
    var group = groups[i];
    writeLog('INFO', (i + 1) + '. ' + group.name + 
             ' (実施日: ' + group.implementationDay + '営業日目' +
             ', 実施順: ' + group.sequence + 
             ', 元値: ' + group.implementationDayRaw + '/' + group.sequenceRaw + ')');
  }
  writeLog('INFO', '=== ソート前結果終了 ===');
  
  // 実施日（第1キー）とコンテンツ実施順（第2キー）でソート（改良版）
  var sortedGroups = groups.sort(function(a, b) { 
    // デフォルト値の扱いを改善
    var aDayValue = (a.implementationDay && a.implementationDay !== 999) ? a.implementationDay : 50;
    var bDayValue = (b.implementationDay && b.implementationDay !== 999) ? b.implementationDay : 50;
    var aSeqValue = (a.sequence && a.sequence !== 999) ? a.sequence : 100;
    var bSeqValue = (b.sequence && b.sequence !== 999) ? b.sequence : 100;
    
    writeLog('DEBUG', 'ソート比較: ' + a.name + '(day:' + aDayValue + ',seq:' + aSeqValue + ') vs ' + 
             b.name + '(day:' + bDayValue + ',seq:' + bSeqValue + ')');
    
    // 第1キー: 実施日比較
    var dayDiff = aDayValue - bDayValue;
    if (dayDiff !== 0) {
      writeLog('DEBUG', '実施日による順序決定: ' + dayDiff + ' (' + a.name + ' → ' + b.name + ')');
      return dayDiff;
    }
    
    // 第2キー: 実施順比較
    var seqDiff = aSeqValue - bSeqValue;
    if (seqDiff !== 0) {
      writeLog('DEBUG', '実施順による順序決定: ' + seqDiff + ' (' + a.name + ' → ' + b.name + ')');
      return seqDiff;
    }
    
    // 第3キー: 研修名による辞書順（同一実施日・同一順序の場合の安定性確保）
    var nameComparison = a.name.localeCompare(b.name, 'ja');
    writeLog('DEBUG', '研修名による順序決定: ' + nameComparison + ' (' + a.name + ' → ' + b.name + ')');
    return nameComparison;
  });
  
  // ソート結果をログ出力
  writeLog('INFO', '=== ソート後の研修順序 ===');
  for (var i = 0; i < sortedGroups.length; i++) {
    var group = sortedGroups[i];
    writeLog('INFO', (i + 1) + '. ' + group.name + 
             ' (実施日: ' + group.implementationDay + '営業日目' +
             ', 実施順: ' + group.sequence + 
             ', 元値: ' + group.implementationDayRaw + '/' + group.sequenceRaw + ')');
  }
  writeLog('INFO', '=== ソート結果終了 ===');
  
  return sortedGroups;
}

/**
 * 実施順処理のテスト・デバッグ関数
 */
function testSequenceProcessing() {
  try {
    var newHires = getNewHires();
    if (newHires.length === 0) {
      SpreadsheetApp.getUi().alert('警告', '入社者データがありません。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // 研修マスタから実施順データを直接チェック
    var masterSheet = SpreadsheetApp.openById(SPREADSHEET_IDS.TRAINING_MASTER).getSheetByName(SHEET_NAMES.TRAINING_MASTER);
    var lastRow = masterSheet.getLastRow();
    
    if (lastRow <= 4) {
      SpreadsheetApp.getUi().alert('エラー', '研修マスタにデータがありません。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // テスト関数でも実際の列数を確認
    var lastCol = masterSheet.getLastColumn();
    var maxCol = Math.min(lastCol, 21);
    var masterData = masterSheet.getRange(5, 1, lastRow - 4, maxCol).getValues();
    var sequenceTestResults = [];
    
    writeLog('INFO', '=== 実施順処理テスト開始 ===');
    
    // 実施順の正規化テスト
    for (var i = 0; i < Math.min(15, masterData.length); i++) {
      var row = masterData[i];
      var lv1 = row[0];           // A列: Lv.1 
      var trainingName = row[2];  // C列: 研修名称
      var implementationDay = row[7];  // H列: 実施日
      var sequence = row[8];      // I列: コンテンツ実施順
      
      if (!trainingName || trainingName.trim() === '') continue;
      
      // A列フィルタリング
      if (lv1 !== 'DX ONB' && lv1 !== 'ビジネススキル研修') {
        writeLog('DEBUG', 'テスト: A列フィルターによりスキップ - ' + trainingName + ' (A列: ' + lv1 + ')');
        continue;
      }
      
      writeLog('INFO', 'テスト対象: ' + trainingName + ' (実施日: ' + implementationDay + ', 実施順: ' + sequence + ')');
      
      // 実施日正規化
      var normalizedImplementationDay = 999;
      if (implementationDay) {
        if (typeof implementationDay === 'number') {
          normalizedImplementationDay = implementationDay;
        } else if (typeof implementationDay === 'string') {
          var dayMatch = implementationDay.toString().match(/^(\d+)/);
          if (dayMatch) {
            normalizedImplementationDay = parseInt(dayMatch[1]);
          }
        }
      }
      
      // 実施順正規化（Logic.gsと同じロジック）
      var normalizedSequence = 999;
      writeLog('DEBUG', '実施順正規化開始: ' + trainingName + ', 元値=' + sequence + ', 型=' + typeof sequence);
      
      if (sequence) {
        if (typeof sequence === 'number') {
          normalizedSequence = sequence;
          writeLog('DEBUG', '数値型実施順: ' + sequence + ' → ' + normalizedSequence);
        } else {
          var seqStr = sequence.toString().trim();
          writeLog('DEBUG', '文字列型実施順処理: "' + seqStr + '"');
          
          // パターン1: 先頭の数字
          var numMatch = seqStr.match(/^(\d+)/);
          if (numMatch) {
            normalizedSequence = parseInt(numMatch[1]);
            writeLog('DEBUG', '先頭数字マッチ: "' + seqStr + '" → ' + normalizedSequence);
          } 
          // パターン2: 丸数字
          else {
            var circleNumbers = ['①', '②', '③', '④', '⑤', '⑥', '⑦', '⑧', '⑨', '⑩'];
            var found = false;
            for (var ci = 0; ci < circleNumbers.length; ci++) {
              if (seqStr.indexOf(circleNumbers[ci]) !== -1) {
                normalizedSequence = ci + 1;
                writeLog('DEBUG', '丸数字マッチ: "' + seqStr + '" (' + circleNumbers[ci] + ') → ' + normalizedSequence);
                found = true;
                break;
              }
            }
            
            if (!found) {
              writeLog('WARN', '実施順パターン認識できず: "' + seqStr + '" → デフォルト値 999');
            }
          }
        }
      } else {
        writeLog('DEBUG', '実施順が空のため、デフォルト値 999 を使用');
      }
      
      sequenceTestResults.push({
        name: trainingName,
        originalDay: implementationDay,
        normalizedDay: normalizedImplementationDay,
        originalSeq: sequence,
        normalizedSeq: normalizedSequence
      });
    }
    
    writeLog('INFO', '=== ソート前の状態 ===');
    for (var i = 0; i < sequenceTestResults.length; i++) {
      var result = sequenceTestResults[i];
      writeLog('INFO', (i + 1) + '. ' + result.name + 
               ' (実施日: ' + result.normalizedDay + '営業日目' +
               ', 実施順: ' + result.normalizedSeq + 
               ', 元値: ' + result.originalDay + '/' + result.originalSeq + ')');
    }
    
    // 結果をソート（改良版）
    sequenceTestResults.sort(function(a, b) {
      // デフォルト値の扱いを改善
      var aDayValue = (a.normalizedDay && a.normalizedDay !== 999) ? a.normalizedDay : 50;
      var bDayValue = (b.normalizedDay && b.normalizedDay !== 999) ? b.normalizedDay : 50;
      var aSeqValue = (a.normalizedSeq && a.normalizedSeq !== 999) ? a.normalizedSeq : 100;
      var bSeqValue = (b.normalizedSeq && b.normalizedSeq !== 999) ? b.normalizedSeq : 100;
      
      var dayDiff = aDayValue - bDayValue;
      if (dayDiff !== 0) {
        writeLog('DEBUG', 'ソート: 実施日による比較 ' + a.name + '(' + aDayValue + ') vs ' + b.name + '(' + bDayValue + ') = ' + dayDiff);
        return dayDiff;
      }
      var seqDiff = aSeqValue - bSeqValue;
      if (seqDiff !== 0) {
        writeLog('DEBUG', 'ソート: 実施順による比較 ' + a.name + '(' + aSeqValue + ') vs ' + b.name + '(' + bSeqValue + ') = ' + seqDiff);
        return seqDiff;
      }
      // 第3キー: 研修名による辞書順
      return a.name.localeCompare(b.name, 'ja');
    });
    
    writeLog('INFO', '=== ソート後の状態 ===');
    for (var i = 0; i < sequenceTestResults.length; i++) {
      var result = sequenceTestResults[i];
      writeLog('INFO', (i + 1) + '. ' + result.name + 
               ' (実施日: ' + result.normalizedDay + '営業日目' +
               ', 実施順: ' + result.normalizedSeq + 
               ', 元値: ' + result.originalDay + '/' + result.originalSeq + ')');
    }
    
    // 結果表示
    var message = '実施順処理テスト結果（ソート後）:\n\n';
    for (var i = 0; i < Math.min(10, sequenceTestResults.length); i++) {
      var result = sequenceTestResults[i];
      message += (i + 1) + '. ' + result.name + '\n';
      message += '   実施日: ' + result.originalDay + ' → ' + result.normalizedDay + '\n';
      message += '   実施順: ' + result.originalSeq + ' → ' + result.normalizedSeq + '\n\n';
    }
    
    if (sequenceTestResults.length > 10) {
      message += '... 他' + (sequenceTestResults.length - 10) + '件\n\n';
    }
    
    message += '詳細なログは実行ログシートをご確認ください。';
    
    SpreadsheetApp.getUi().alert('テスト結果', message, SpreadsheetApp.getUi().ButtonSet.OK);
    writeLog('INFO', '=== 実施順処理テスト完了 ===');
    
  } catch (e) {
    writeLog('ERROR', '実施順処理テストでエラー: ' + e.message);
    SpreadsheetApp.getUi().alert('エラー', '実施順処理テストでエラーが発生しました:\n' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
} 