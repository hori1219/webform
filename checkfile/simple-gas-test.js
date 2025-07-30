/**
 * 🧪 超シンプルテスト用 Google Apps Script
 * 
 * 目的: 会社名のみをF列28行目に転記
 * スプレッドシートID: 1xfFlHJihYyhJ-CKP3Aj5veN9c9lanolsTj4kyvR_9R0
 * シート名: シート1
 * 転記先: F28セル
 */

// 設定
const TEST_SPREADSHEET_ID = '1xfFlHJihYyhJ-CKP3Aj5veN9c9lanolsTj4kyvR_9R0';
const TEST_SHEET_NAME = 'シート1';
const TEST_TARGET_ROW = 28;    // 28行目
const TEST_TARGET_COL = 6;     // F列（1-based）

/**
 * GET リクエスト処理
 */
function doGet(e) {
  return ContentService
    .createTextOutput('🧪 スモールテスト用API - 動作中')
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * POST リクエスト処理（超シンプル版）
 */
function doPost(e) {
  try {
    console.log('=== 🧪 テスト開始 ===');
    
    // 1. データ受信確認
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('POSTデータが空です');
    }
    
    const data = JSON.parse(e.postData.contents);
    console.log('受信データ:', data);
    
    const companyName = data.companyName || '';
    if (!companyName) {
      throw new Error('会社名が空です');
    }
    
    console.log('会社名:', companyName);
    
    // 2. スプレッドシート接続
    console.log('スプレッドシートID:', TEST_SPREADSHEET_ID);
    console.log('シート名:', TEST_SHEET_NAME);
    
    const spreadsheet = SpreadsheetApp.openById(TEST_SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(TEST_SHEET_NAME);
    
    if (!sheet) {
      throw new Error(`シート "${TEST_SHEET_NAME}" が見つかりません`);
    }
    
    console.log('スプレッドシート接続成功');
    
    // 3. 現在の状況確認
    const currentValue = sheet.getRange(TEST_TARGET_ROW, TEST_TARGET_COL).getValue();
    console.log(`現在のF${TEST_TARGET_ROW}の値:`, currentValue);
    
    // 4. 転記実行
    const timestamp = new Date().toLocaleString('ja-JP');
    const writeValue = `${companyName} (${timestamp})`;
    
    console.log('転記する値:', writeValue);
    console.log(`転記先: F${TEST_TARGET_ROW}`);
    
    sheet.getRange(TEST_TARGET_ROW, TEST_TARGET_COL).setValue(writeValue);
    
    console.log('✅ 転記完了');
    
    // 5. 転記後の確認
    const afterValue = sheet.getRange(TEST_TARGET_ROW, TEST_TARGET_COL).getValue();
    console.log('転記後の値:', afterValue);
    
    console.log('=== 🧪 テスト成功 ===');
    
    return ContentService
      .createTextOutput(JSON.stringify({
        result: 'success',
        message: `会社名「${companyName}」をF${TEST_TARGET_ROW}に転記しました`,
        companyName: companyName,
        targetCell: `F${TEST_TARGET_ROW}`,
        timestamp: timestamp,
        writtenValue: writeValue
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    console.error('=== 🚨 テストエラー ===');
    console.error('エラー:', error.message);
    console.error('スタック:', error.stack);
    
    return ContentService
      .createTextOutput(JSON.stringify({
        result: 'error',
        message: error.message,
        timestamp: new Date().toISOString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 手動テスト用関数（Apps Scriptエディターから実行可能）
 */
function manualTest() {
  try {
    console.log('=== 手動テスト開始 ===');
    
    const testData = {
      companyName: 'テスト会社_' + new Date().getTime()
    };
    
    const mockEvent = {
      postData: {
        contents: JSON.stringify(testData)
      }
    };
    
    const result = doPost(mockEvent);
    console.log('テスト結果:', result.getContent());
    
    console.log('=== 手動テスト完了 ===');
    
  } catch (error) {
    console.error('手動テストエラー:', error);
  }
}

/**
 * スプレッドシート接続テスト
 */
function testConnection() {
  try {
    console.log('=== 接続テスト開始 ===');
    
    const spreadsheet = SpreadsheetApp.openById(TEST_SPREADSHEET_ID);
    console.log('スプレッドシート名:', spreadsheet.getName());
    
    const sheet = spreadsheet.getSheetByName(TEST_SHEET_NAME);
    console.log('シート名:', sheet.getName());
    console.log('最終行:', sheet.getLastRow());
    console.log('最終列:', sheet.getLastColumn());
    
    const currentF28 = sheet.getRange('F28').getValue();
    console.log('現在のF28の値:', currentF28);
    
    console.log('=== 接続テスト成功 ===');
    
  } catch (error) {
    console.error('接続テストエラー:', error);
  }
}