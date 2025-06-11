// =========================================
// 通知関連ユーティリティ
// =========================================

/**
 * 担当者へ通知メールを送信する
 * @param {string} subject - 件名
 * @param {string} body - 本文
 */
function sendNotificationEmail(subject, body) {
    GmailApp.sendEmail(NOTIFICATION_EMAIL, subject, body);
}

/**
 * GASログに情報を出力する
 * @param {string} level - ログレベル (INFO, ERROR, DEBUG)
 * @param {string} message - ログメッセージ
 * @param {Object} data - 追加データ（オプション）
 */
function writeLog(level, message, data) {
    var timestamp = new Date().toISOString();
    var logMessage = '[' + timestamp + '] ' + level + ': ' + message;
    
    if (data) {
        logMessage += ' | データ: ' + JSON.stringify(data);
    }
    
    console.log(logMessage);
}

/**
 * 対象者情報をログに出力する
 * @param {Array<Object>} newHires - 対象者配列
 */
function logTargetUsers(newHires) {
    writeLog('INFO', '=== 対象者情報 ===');
    writeLog('INFO', '対象者数: ' + newHires.length + '名');
    
    for (var i = 0; i < newHires.length; i++) {
        var hire = newHires[i];
        writeLog('INFO', '対象者' + (i + 1) + ': ' + hire.name + ' (' + hire.rank + ') - ' + hire.email);
    }
    writeLog('INFO', '==================');
}

/**
 * エラー情報を詳細にログ出力する
 * @param {Error} error - エラーオブジェクト
 * @param {string} context - エラー発生箇所
 * @param {Object} additionalData - 追加情報
 */
function logError(error, context, additionalData) {
    writeLog('ERROR', '=== エラー発生 ===');
    writeLog('ERROR', '発生箇所: ' + context);
    writeLog('ERROR', 'エラーメッセージ: ' + error.message);
    writeLog('ERROR', 'スタックトレース: ' + error.stack);
    
    if (additionalData) {
        writeLog('ERROR', '追加情報', additionalData);
    }
    writeLog('ERROR', '==================');
} 