/**
 * 出荷証明書作成依頼WebフォームからのデータをGoogle Sheetsに転記するGoogle Apps Script
 * 
 * セットアップ手順:
 * 1. Google Sheetsで「拡張機能」→「Apps Script」を選択
 * 2. このコードをCode.gsにコピー＆ペースト
 * 3. 「デプロイ」→「新しいデプロイ」でウェブアプリとして公開
 * 4. アクセス権限を「全員」に設定
 * 5. 取得したURLをHTMLファイルのGOOGLE_SCRIPT_URLに設定
 */

// 設定項目
const SPREADSHEET_ID = '1tw3L-PQpr2D4o9GMISCQkfEMfR2aqr8aQtYcsqGptYY'; // 中央データベースのスプレッドシートID
const SHEET_NAME = 'シート1'; // シート名
const NOTIFICATION_EMAIL = 'YOUR_NOTIFICATION_EMAIL_HERE'; // 通知メールアドレス（オプション）

// 注意: 実際のGoogle Apps Scriptでは、既存のDB_ID、DB_NAME定数を使用してください

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
    
    if (!sheet) {
      throw new Error(`シート "${SHEET_NAME}" が見つかりません`);
    }
    
    // 管理番号を生成
    const controlId = generateControlId(sheet);
    
    // データを行として追加（中央データベースの構造に合わせる）
    const rowData = buildRowData(controlId, data);
    
    sheet.appendRow(rowData);
    
    // 最新行のフォーマットを設定
    formatLatestRow(sheet);
    
    // 通知メールを送信（設定されている場合）
    if (NOTIFICATION_EMAIL && NOTIFICATION_EMAIL !== 'YOUR_NOTIFICATION_EMAIL_HERE') {
      sendNotificationEmail(data, controlId);
    }
    
    console.log('データ追加成功:', controlId);
    
    return ContentService
      .createTextOutput(JSON.stringify({
        result: 'success',
        message: '出荷証明書作成依頼が正常に登録されました',
        controlId: controlId,
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
 * 管理番号を生成する関数
 */
function generateControlId(sheet) {
  const lastRow = sheet.getLastRow();
  let lastIdNumber = 0;
  
  // 最後の行から管理番号を取得
  if (lastRow > 1) {
    const lastCell = sheet.getRange(lastRow, 2); // B列（管理番号列）
    const lastId = String(lastCell.getValue() || '');
    const parts = lastId.split('-');
    if (parts.length >= 2) {
      lastIdNumber = parseInt(parts[1], 10) || 0;
    }
  }
  
  const nextIdNumber = lastIdNumber + 1;
  const today = new Date();
  const yyyy = today.getFullYear();
  const mm = ('0' + (today.getMonth() + 1)).slice(-2);
  const dd = ('0' + today.getDate()).slice(-2);
  const seq = ('000' + nextIdNumber).slice(-3);
  
  return `${yyyy}${mm}${dd}-${seq}-1`;
}

/**
 * 中央データベース構造に合わせてデータを構築する関数
 */
function buildRowData(controlId, data) {
  const rowData = [
    false,                                    // A: チェックボックス
    controlId,                               // B: 管理番号
    new Date(),                              // C: 申請日時
    '申請中',                                // D: ステータス
    1,                                       // E: 版数
    data.companyName || '',                  // F: 会社名
    data.contactPerson || '',                // G: 申請者名
    data.phoneNumber || '',                  // H: 申請者TEL
    data.faxNumber || '',                    // I: FAX
    data.addressee || '',                    // J: 宛名
    data.honorific || '様',                  // K: 敬称
    data.constructionName || '',             // L: 工事名
    data.constructionAddress || '',          // M: 工事住所
    data.creationDate || '',                 // N: 作成日
  ];

  // 業者情報（最大3業者）
  const contractors = data.contractors || [];
  for (let i = 0; i < 3; i++) {
    if (contractors[i]) {
      rowData.push(contractors[i].type || '');     // 業者分類
      rowData.push(contractors[i].name || '');     // 業者名
    } else {
      rowData.push('', '');
    }
  }

  // 商品情報（最大7商品）
  const products = data.products || [];
  for (let i = 0; i < 7; i++) {
    if (products[i]) {
      rowData.push(products[i].productName || '');   // 商品名
      rowData.push(products[i].quantity || '');      // 数量
      rowData.push(products[i].lotNumber || '');     // ロットNo.
      rowData.push(products[i].shipmentDate || '');  // 出荷年月日
    } else {
      rowData.push('', '', '', '');
    }
  }

  // 必要書類
  const documents = data.documents || [];
  rowData.push(documents.join(', '));              // 書類リスト

  // 送信先情報
  rowData.push(data.destEmailAddress || '');       // 客先メールアドレス
  rowData.push(data.destCompanyName || '');        // 送信先会社名
  rowData.push(data.destContactPerson || '');      // 送信先担当者
  rowData.push(data.destPhoneNumber || '');        // 送信先電話番号

  // その他の項目（空で初期化）
  const additionalFields = 20; // 必要に応じて調整
  for (let i = 0; i < additionalFields; i++) {
    rowData.push('');
  }

  // 最終更新日時
  rowData.push(new Date());

  return rowData;
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
    
    // 日付フォーマット（申請日時、作成日）
    sheet.getRange(lastRow, 3).setNumberFormat('yyyy/mm/dd hh:mm:ss'); // 申請日時
    if (sheet.getRange(lastRow, 14).getValue()) {
      sheet.getRange(lastRow, 14).setNumberFormat('yyyy/mm/dd'); // 作成日
    }
    
    // 出荷年月日の列（商品情報部分）
    for (let i = 0; i < 7; i++) {
      const dateCol = 24 + (i * 4); // 出荷年月日の列位置
      if (sheet.getRange(lastRow, dateCol).getValue()) {
        sheet.getRange(lastRow, dateCol).setNumberFormat('yyyy/mm/dd');
      }
    }
  }
}

/**
 * 通知メールを送信する関数
 */
function sendNotificationEmail(data, controlId) {
  try {
    const subject = `【出荷証明書依頼】新しい依頼が登録されました - 管理番号: ${controlId}`;
    
    // 商品情報を文字列化
    const productList = data.products.map((product, index) => 
      `${index + 1}. ${product.productName} (数量: ${product.quantity}, ロット: ${product.lotNumber}, 出荷日: ${product.shipmentDate})`
    ).join('\n');

    // 業者情報を文字列化
    const contractorList = data.contractors.map((contractor, index) => 
      `${index + 1}. ${contractor.type}: ${contractor.name}`
    ).join('\n');

    // 必要書類を文字列化
    const documentList = data.documents.join(', ');
    
    const body = `
出荷証明書作成依頼が新しく登録されました。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

【管理番号】
${controlId}

【発信元情報】
会社名: ${data.companyName}
担当者: ${data.contactPerson}
電話番号: ${data.phoneNumber}
FAX番号: ${data.faxNumber}
メール: ${data.emailAddress}

【基本情報】
宛名: ${data.addressee} ${data.honorific}
工事名: ${data.constructionName}
工事住所: ${data.constructionAddress}
作成日: ${data.creationDate}

【業者情報】
${contractorList}

【商品情報】
${productList}

【必要書類】
${documentList}

【送信先情報】
会社名: ${data.destCompanyName}
担当者: ${data.destContactPerson}
電話番号: ${data.destPhoneNumber}
メール: ${data.destEmailAddress}

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
    receiptNumber: '20250130-123456-789',
    companyName: 'テスト会社',
    contactPerson: 'テスト太郎',
    phoneNumber: '03-1234-5678',
    faxNumber: '03-1234-5679',
    emailAddress: 'test@test.com',
    addressee: 'テスト株式会社',
    honorific: '御中',
    constructionName: 'テスト工事',
    constructionAddress: '東京都テスト区テスト1-1-1',
    creationDate: '2025-01-30',
    destCompanyName: '送信先会社',
    destContactPerson: '送信先太郎',
    destPhoneNumber: '03-8765-4321',
    destEmailAddress: 'dest@test.com',
    contractors: [
      { type: '施工業者', name: 'テスト施工' },
      { type: '塗装業者', name: 'テスト塗装' }
    ],
    products: [
      { shipmentDate: '2025-01-30', productName: 'テスト商品1', lotNumber: 'LOT001', quantity: '10' },
      { shipmentDate: '2025-01-31', productName: 'テスト商品2', lotNumber: 'LOT002', quantity: '5' }
    ],
    documents: ['出荷証明書', '成分表・試験成績書']
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
      // ステータス別集計
      const statuses = sheet.getRange(2, 4, lastRow - 1, 1).getValues().flat();
      const statusCount = {};
      statuses.forEach(status => {
        if (status) {
          statusCount[status] = (statusCount[status] || 0) + 1;
        }
      });
      console.log('ステータス別集計:', statusCount);
      
      // 会社別集計
      const companies = sheet.getRange(2, 6, lastRow - 1, 1).getValues().flat();
      const companyCount = {};
      companies.forEach(company => {
        if (company) {
          companyCount[company] = (companyCount[company] || 0) + 1;
        }
      });
      console.log('会社別集計:', companyCount);
      
      // 直近7日間のデータ数
      const today = new Date();
      const weekAgo = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);
      const timestamps = sheet.getRange(2, 3, lastRow - 1, 1).getValues().flat();
      const recentCount = timestamps.filter(timestamp => {
        const date = new Date(timestamp);
        return date >= weekAgo;
      }).length;
      console.log('直近7日間の依頼件数:', recentCount);
    }
    
  } catch (error) {
    console.error('データ分析エラー:', error);
  }
}