/**
 * 🧪 Phase2テスト用 Google Apps Script
 * 
 * 対象項目:
 * 【発信元】会社名、担当者名、電話番号、FAX、メールアドレス (F～J列)
 * 【基本情報】宛名、敬称、工事名、工事住所 (K～N列)
 * 
 * 転記先: F28～N28（28行目のF列からN列まで）
 * スプレッドシートID: 1xfFlHJihYyhJ-CKP3Aj5veN9c9lanolsTj4kyvR_9R0
 */

// 設定
const TEST_SPREADSHEET_ID = '1xfFlHJihYyhJ-CKP3Aj5veN9c9lanolsTj4kyvR_9R0';
const TEST_SHEET_NAME = 'シート1';
const TEST_TARGET_ROW = 28;

// 列マッピング（1-based）
const COLUMN_MAP = {
    companyName: 6,        // F列: 会社名
    contactPerson: 7,      // G列: 担当者名
    phoneNumber: 8,        // H列: 電話番号
    faxNumber: 9,          // I列: FAX
    emailAddress: 10,      // J列: メールアドレス
    addressee: 11,         // K列: 宛名
    honorific: 12,         // L列: 敬称
    constructionName: 13,  // M列: 工事名
    constructionAddress: 14 // N列: 工事住所
};

/**
 * GET リクエスト処理
 */
function doGet(e) {
    return ContentService
        .createTextOutput('🧪 Phase2テスト用API - 発信元＋基本情報転記システム')
        .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * POST リクエスト処理（Phase2版）
 */
function doPost(e) {
    try {
        console.log('=== 🧪 Phase2テスト開始 ===');
        
        // 1. データ受信確認
        if (!e || !e.postData || !e.postData.contents) {
            throw new Error('POSTデータが空です');
        }
        
        const data = JSON.parse(e.postData.contents);
        console.log('Phase2受信データ:', JSON.stringify(data, null, 2));
        
        // 2. 必須項目チェック
        const requiredFields = ['companyName', 'contactPerson', 'phoneNumber', 'emailAddress', 'addressee', 'constructionName'];
        const missingFields = requiredFields.filter(field => !data[field] || !data[field].trim());
        
        if (missingFields.length > 0) {
            throw new Error(`必須項目が未入力です: ${missingFields.join(', ')}`);
        }
        
        // 3. スプレッドシート接続
        console.log('スプレッドシート接続中...');
        const spreadsheet = SpreadsheetApp.openById(TEST_SPREADSHEET_ID);
        const sheet = spreadsheet.getSheetByName(TEST_SHEET_NAME);
        
        if (!sheet) {
            throw new Error(`シート "${TEST_SHEET_NAME}" が見つかりません`);
        }
        
        console.log('スプレッドシート接続成功');
        console.log('シート名:', sheet.getName());
        console.log('対象行:', TEST_TARGET_ROW);
        
        // 4. 現在の値確認（デバッグ用）
        console.log('=== 転記前の現在値 ===');
        Object.entries(COLUMN_MAP).forEach(([field, col]) => {
            const currentValue = sheet.getRange(TEST_TARGET_ROW, col).getValue();
            console.log(`${field} (${getColumnLetter(col)}${TEST_TARGET_ROW}):`, currentValue);
        });
        
        // 5. データ転記実行
        console.log('=== データ転記開始 ===');
        const timestamp = new Date().toLocaleString('ja-JP');
        const results = {};
        
        // 各項目を個別に転記
        Object.entries(COLUMN_MAP).forEach(([field, col]) => {
            let value = data[field] || '';
            
            // 特別処理
            if (field === 'honorific' && !value) {
                value = '様'; // デフォルト値
            }
            
            // タイムスタンプ付きで転記（デバッグ用）
            const writeValue = value ? `${value} (${timestamp})` : '';
            
            console.log(`転記: ${field} → ${getColumnLetter(col)}${TEST_TARGET_ROW} = "${writeValue}"`);
            sheet.getRange(TEST_TARGET_ROW, col).setValue(writeValue);
            
            results[field] = {
                column: getColumnLetter(col),
                value: writeValue
            };
        });
        
        console.log('=== データ転記完了 ===');
        
        // 6. 転記後確認
        console.log('=== 転記後の確認 ===');
        Object.entries(COLUMN_MAP).forEach(([field, col]) => {
            const afterValue = sheet.getRange(TEST_TARGET_ROW, col).getValue();
            console.log(`${field} (${getColumnLetter(col)}${TEST_TARGET_ROW}):`, afterValue);
        });
        
        console.log('=== 🧪 Phase2テスト成功 ===');
        
        return ContentService
            .createTextOutput(JSON.stringify({
                result: 'success',
                message: 'Phase2テスト完了：発信元＋基本情報を転記しました',
                transferredItems: Object.keys(COLUMN_MAP).length,
                targetRow: TEST_TARGET_ROW,
                columnRange: `F${TEST_TARGET_ROW}:N${TEST_TARGET_ROW}`,
                timestamp: timestamp,
                details: results
            }))
            .setMimeType(ContentService.MimeType.JSON);
            
    } catch (error) {
        console.error('=== 🚨 Phase2テストエラー ===');
        console.error('エラー:', error.message);
        console.error('スタック:', error.stack);
        
        return ContentService
            .createTextOutput(JSON.stringify({
                result: 'error',
                message: `Phase2テストエラー: ${error.message}`,
                timestamp: new Date().toISOString()
            }))
            .setMimeType(ContentService.MimeType.JSON);
    }
}

/**
 * 列番号を列文字に変換（A=1, B=2, ...）
 */
function getColumnLetter(columnNumber) {
    let temp, letter = '';
    while (columnNumber > 0) {
        temp = (columnNumber - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        columnNumber = (columnNumber - temp - 1) / 26;
    }
    return letter;
}

/**
 * Phase2手動テスト用関数
 */
function manualTestPhase2() {
    try {
        console.log('=== Phase2手動テスト開始 ===');
        
        const testData = {
            // 発信元
            companyName: 'テスト会社Phase2',
            contactPerson: 'テスト太郎',
            phoneNumber: '03-1234-5678',
            faxNumber: '03-1234-5679',
            emailAddress: 'test@example.com',
            
            // 基本情報
            addressee: 'テスト株式会社',
            honorific: '御中',
            constructionName: 'テスト工事Phase2',
            constructionAddress: '東京都テスト区テスト町1-2-3',
            
            timestamp: new Date().toISOString(),
            testMode: 'phase2'
        };
        
        const mockEvent = {
            postData: {
                contents: JSON.stringify(testData)
            }
        };
        
        const result = doPost(mockEvent);
        console.log('Phase2テスト結果:', JSON.parse(result.getContent()));
        
        console.log('=== Phase2手動テスト完了 ===');
        
    } catch (error) {
        console.error('Phase2手動テストエラー:', error);
    }
}

/**
 * 列マッピング確認用関数
 */
function checkColumnMapping() {
    try {
        console.log('=== 列マッピング確認 ===');
        
        const spreadsheet = SpreadsheetApp.openById(TEST_SPREADSHEET_ID);
        const sheet = spreadsheet.getSheetByName(TEST_SHEET_NAME);
        
        console.log(`対象行: ${TEST_TARGET_ROW}`);
        console.log('列マッピング:');
        
        Object.entries(COLUMN_MAP).forEach(([field, col]) => {
            const colLetter = getColumnLetter(col);
            const currentValue = sheet.getRange(TEST_TARGET_ROW, col).getValue();
            console.log(`${field}: ${colLetter}${TEST_TARGET_ROW} (列${col}) = "${currentValue}"`);
        });
        
        console.log('=== 列マッピング確認完了 ===');
        
    } catch (error) {
        console.error('列マッピング確認エラー:', error);
    }
}

/**
 * Phase2用データクリア関数（テスト用）
 */
function clearPhase2Data() {
    try {
        console.log('=== Phase2データクリア開始 ===');
        
        const spreadsheet = SpreadsheetApp.openById(TEST_SPREADSHEET_ID);
        const sheet = spreadsheet.getSheetByName(TEST_SHEET_NAME);
        
        Object.values(COLUMN_MAP).forEach(col => {
            sheet.getRange(TEST_TARGET_ROW, col).setValue('');
        });
        
        console.log('=== Phase2データクリア完了 ===');
        
    } catch (error) {
        console.error('Phase2データクリアエラー:', error);
    }
}