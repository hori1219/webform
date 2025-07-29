/**
 * 出荷証明書WebフォームからのデータをGoogle Sheetsに転記するGoogle Apps Script
 * 
 * セットアップ手順:
 * 1. Google Sheetsで「拡張機能」→「Apps Script」を選択
 * 2. このコードをCode.gsにコピー＆ペースト
 * 3. 「デプロイ」→「新しいデプロイ」でウェブアプリとして公開
 * 4. アクセス権限を「全員」に設定
 * 5. 取得したURLをHTMLファイルのGOOGLE_SCRIPT_URLに設定
 */

// 設定項目
const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE'; // スプレッドシートのIDを設定
const SHEET_NAME = '出荷証明書'; // シート名を設定
const NOTIFICATION_EMAIL = 'YOUR_NOTIFICATION_EMAIL_HERE'; // 通知メールアドレス（オプション）

/**
 * POST リクエストを処理する関数
 * Webフォームからのデータをスプレッドシートに追加
 */
function doPost(e) {
  try {
    // POSTデータを解析
    const data = JSON.parse(e.postData.contents);
    console.log('受信データ:', data);
    
    // スプレッドシートを開く
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = spreadsheet.getSheetByName(SHEET_NAME);
    
    // シートが存在しない場合は作成
    if (!sheet) {
      sheet = spreadsheet.insertSheet(SHEET_NAME);
      setupSheetHeaders(sheet);
    }
    
    // ヘッダーが設定されていない場合は設定
    if (sheet.getLastRow() === 0) {
      setupSheetHeaders(sheet);
    }
    
    // データを行として追加
    const rowData = [
      data.timestamp || new Date().toLocaleString('ja-JP'),
      data.senderCompany || '',
      data.senderName || '',
      data.senderEmail || '',
      data.senderPhone || '',
      data.receiverCompany || '',
      data.receiverName || '',
      data.receiverAddress || '',
      data.receiverEmail || '',
      data.receiverPhone || '',
      data.shippingDate || '',
      data.shippingMethod || '',
      data.trackingNumber || '',
      data.packageCount || '',
      data.productName || '',
      data.productCode || '',
      data.quantity || '',
      data.unit || '',
      data.weight || '',
      data.value || '',
      data.notes || ''
    ];
    
    sheet.appendRow(rowData);
    
    // 最新行のフォーマットを設定
    formatLatestRow(sheet);
    
    // 通知メールを送信（設定されている場合）
    if (NOTIFICATION_EMAIL && NOTIFICATION_EMAIL !== 'YOUR_NOTIFICATION_EMAIL_HERE') {
      sendNotificationEmail(data);
    }
    
    console.log('データ追加成功:', rowData);
    
    return ContentService
      .createTextOutput(JSON.stringify({
        result: 'success',
        message: '出荷証明書が正常に登録されました',
        timestamp: new Date().toISOString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    console.error('エラー発生:', error);
    
    return ContentService
      .createTextOutput(JSON.stringify({
        result: 'error',
        message: 'データの処理中にエラーが発生しました: ' + error.toString(),
        timestamp: new Date().toISOString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * シートのヘッダーを設定する関数
 */
function setupSheetHeaders(sheet) {
  const headers = [
    'タイムスタンプ',
    '発送者会社名',
    '発送者担当者',
    '発送者メール',
    '発送者電話',
    '発送先会社名',
    '発送先担当者',
    '発送先住所',
    '発送先メール',
    '発送先電話',
    '発送日',
    '発送方法',
    '追跡番号',
    '梱包数',
    '商品名',
    '商品コード',
    '数量',
    '単位',
    '重量(kg)',
    '商品価格(円)',
    '備考'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダー行のフォーマット設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4285F4');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // 列幅の自動調整
  sheet.autoResizeColumns(1, headers.length);
  
  console.log('ヘッダー設定完了');
}

/**
 * 最新行のフォーマットを設定する関数
 */
function formatLatestRow(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const range = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn());
    
    // 交互の背景色設定
    if (lastRow % 2 === 0) {
      range.setBackground('#F8F9FA');
    }
    
    // 境界線の設定
    range.setBorder(true, true, true, true, true, true);
    
    // 発送日の列（11列目）の日付フォーマット
    sheet.getRange(lastRow, 11).setNumberFormat('yyyy/mm/dd');
    
    // 数量、重量、価格の列の数値フォーマット
    sheet.getRange(lastRow, 14).setNumberFormat('#,##0'); // 梱包数
    sheet.getRange(lastRow, 17).setNumberFormat('#,##0'); // 数量
    sheet.getRange(lastRow, 19).setNumberFormat('#,##0.0'); // 重量
    sheet.getRange(lastRow, 20).setNumberFormat('#,##0'); // 価格
  }
}

/**
 * 通知メールを送信する関数
 */
function sendNotificationEmail(data) {
  try {
    const subject = `【出荷証明書】新しい出荷が登録されました - ${data.productName}`;
    
    const body = `
出荷証明書が新しく登録されました。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

【発送者情報】
会社名: ${data.senderCompany}
担当者: ${data.senderName}
メール: ${data.senderEmail}
電話番号: ${data.senderPhone}

【発送先情報】
会社名: ${data.receiverCompany}
担当者: ${data.receiverName}
住所: ${data.receiverAddress}
メール: ${data.receiverEmail}
電話番号: ${data.receiverPhone}

【発送情報】
発送日: ${data.shippingDate}
発送方法: ${data.shippingMethod}
追跡番号: ${data.trackingNumber}
梱包数: ${data.packageCount}

【商品情報】
商品名: ${data.productName}
商品コード: ${data.productCode}
数量: ${data.quantity} ${data.unit}
重量: ${data.weight} kg
価格: ${data.value} 円

【備考】
${data.notes}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

登録日時: ${data.timestamp}

このメールは自動送信されています。
    `;
    
    MailApp.sendEmail(NOTIFICATION_EMAIL, subject, body);
    console.log('通知メール送信完了:', NOTIFICATION_EMAIL);
    
  } catch (error) {
    console.error('メール送信エラー:', error);
  }
}

/**
 * テスト用関数 - Apps Scriptエディターから実行可能
 */
function testFunction() {
  const testData = {
    timestamp: new Date().toLocaleString('ja-JP'),
    senderCompany: 'テスト発送会社',
    senderName: 'テスト太郎',
    senderEmail: 'test@example.com',
    senderPhone: '03-1234-5678',
    receiverCompany: 'テスト受取会社',
    receiverName: '受取花子',
    receiverAddress: '東京都渋谷区テスト1-1-1',
    receiverEmail: 'receive@example.com',
    receiverPhone: '03-8765-4321',
    shippingDate: '2024-01-15',
    shippingMethod: '宅急便',
    trackingNumber: '1234-5678-9012',
    packageCount: '2',
    productName: 'テスト商品',
    productCode: 'TEST-001',
    quantity: '10',
    unit: '個',
    weight: '5.5',
    value: '50000',
    notes: 'テスト用の出荷証明書です。'
  };
  
  try {
    const mockEvent = {
      postData: {
        contents: JSON.stringify(testData)
      }
    };
    
    const result = doPost(mockEvent);
    console.log('テスト実行結果:', result.getContent());
  } catch (error) {
    console.error('テスト実行エラー:', error);
  }
}

/**
 * 既存データを分析する関数
 */
function analyzeData() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      console.log('シートが見つかりません:', SHEET_NAME);
      return;
    }
    
    const lastRow = sheet.getLastRow();
    console.log('総データ数:', lastRow - 1); // ヘッダー行を除く
    
    if (lastRow > 1) {
      // 発送方法の集計
      const shippingMethods = sheet.getRange(2, 12, lastRow - 1, 1).getValues().flat();
      const methodCount = {};
      shippingMethods.forEach(method => {
        if (method) {
          methodCount[method] = (methodCount[method] || 0) + 1;
        }
      });
      console.log('発送方法別集計:', methodCount);
      
      // 直近7日間のデータ数
      const today = new Date();
      const weekAgo = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);
      const timestamps = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
      const recentCount = timestamps.filter(timestamp => {
        const date = new Date(timestamp);
        return date >= weekAgo;
      }).length;
      console.log('直近7日間の出荷件数:', recentCount);
    }
    
  } catch (error) {
    console.error('データ分析エラー:', error);
  }
}