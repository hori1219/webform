/**
 * WebフォームからのPOSTリクエストを処理するdoPost関数
 * 新環境用（環境移行版）
 */

// 設定項目（新しいスプレッドシートID）
const DB_ID = '1ouyWnRo70vg7GPjj3YNIEHm6ROrXU8GiQjOOHFsV-B_me3DXZLKFfxdK';
const DB_NAME = 'シート1'; // シート名を確認してください

/**
 * POST リクエストを処理する関数
 * Webフォームからのデータを中央データベースに追加
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  const lockTimeout = 10000; // 10秒のタイムアウト
  
  if (!lock.tryLock(lockTimeout)) {
    console.error('ロック取得失敗');
    return ContentService
      .createTextOutput(JSON.stringify({
        result: 'error',
        message: 'システムが混雑しています。時間をおいて再試行してください。',
        timestamp: new Date().toISOString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  try {
    // POSTデータを解析
    const data = JSON.parse(e.postData.contents);
    console.log('受信データ:', data);
    
    // スプレッドシートを開く
    const spreadsheet = SpreadsheetApp.openById(DB_ID);
    const sheet = spreadsheet.getSheetByName(DB_NAME);
    
    if (!sheet) {
      throw new Error(`シート "${DB_NAME}" が見つかりません`);
    }
    
    // 管理番号を生成
    const controlId = generateControlIdForForm(sheet);
    
    // データを行として追加（中央データベースの構造に合わせる）
    const rowData = buildRowDataForForm(controlId, data);
    
    sheet.appendRow(rowData);
    
    // 最新行のフォーマットを設定
    formatLatestRowForForm(sheet);
    
    // 通知メールを送信
    sendNotificationEmailForForm(data, controlId);
    
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
  } finally {
    lock.releaseLock();
  }
}

/**
 * フォーム用の管理番号を生成する関数
 */
function generateControlIdForForm(sheet) {
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
 * フォーム用のデータ構造を構築する関数
 */
function buildRowDataForForm(controlId, data) {
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
  const documentFlags = {
    '成績書': documents.includes('成績書'),
    'SDS': documents.includes('SDS'),
    '出荷証明書': documents.includes('出荷証明書'),
    'カタログ': documents.includes('カタログ'),
    '塗料': documents.includes('塗料'),
    '希釈剤': documents.includes('希釈剤'),
    '硬化剤': documents.includes('硬化剤')
  };
  
  // 書類フラグを追加
  Object.values(documentFlags).forEach(flag => rowData.push(flag));
  
  // メールアドレス（送信先）
  rowData.push(data.destEmailAddress || '');       // 客先メールアドレス

  // To/Cc/Bcc（初期値は空）
  for (let i = 0; i < 15; i++) { // To×5 + Cc×5 + Bcc×5
    rowData.push('');
  }

  // 作成日
  rowData.push(data.creationDate || '');

  // その他の項目（スプレッドシートの構造に合わせて調整）
  const additionalFields = 20; // 必要に応じて調整
  for (let i = 0; i < additionalFields; i++) {
    rowData.push('');
  }

  return rowData;
}

/**
 * フォーム用の最新行フォーマット設定
 */
function formatLatestRowForForm(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const range = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn());
    
    // 交互の背景色設定
    if (lastRow % 2 === 0) {
      range.setBackground('#F8F9FA');
    }
    
    // 境界線の設定
    range.setBorder(true, true, true, true, true, true);
    
    // 日付フォーマット（申請日時）
    sheet.getRange(lastRow, 3).setNumberFormat('yyyy/mm/dd hh:mm:ss'); 
  }
}

/**
 * フォーム用の通知メール送信
 */
function sendNotificationEmailForForm(data, controlId) {
  try {
    const subject = `【出荷証明書依頼】新しい依頼が登録されました - 管理番号: ${controlId}`;
    
    // 商品情報を文字列化
    const productList = (data.products || []).map((product, index) => 
      `${index + 1}. ${product.productName} (数量: ${product.quantity}, ロット: ${product.lotNumber}, 出荷日: ${product.shipmentDate})`
    ).join('\n');

    // 業者情報を文字列化
    const contractorList = (data.contractors || []).map((contractor, index) => 
      `${index + 1}. ${contractor.type}: ${contractor.name}`
    ).join('\n');

    // 必要書類を文字列化
    const documentList = (data.documents || []).join(', ');
    
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
    
    // 通知先メールアドレス（実際のメールアドレスに変更してください）
    const notificationEmail = 'notification@example.com'; // ★ここを変更してください
    MailApp.sendEmail(notificationEmail, subject, body);
    console.log('通知メール送信完了:', notificationEmail);
    
  } catch (error) {
    console.error('メール送信エラー:', error);
  }
}

/**
 * テスト用：GET リクエストでテストページを表示
 */
function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'ok',
      message: '出荷証明書作成依頼システムが正常に動作しています',
      timestamp: new Date().toISOString()
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * テスト用関数 - Apps Scriptエディタから実行可能
 */
function testFunction() {
  const testData = {
    timestamp: new Date().toLocaleString('ja-JP'),
    receiptNumber: '20250131-123456-789',
    companyName: 'テスト会社',
    contactPerson: 'テスト太郎',
    phoneNumber: '03-1234-5678',
    faxNumber: '03-1234-5679',
    emailAddress: 'test@test.com',
    addressee: 'テスト株式会社',
    honorific: '御中',
    constructionName: 'テスト工事',
    constructionAddress: '東京都テスト区テスト1-1-1',
    creationDate: '2025-01-31',
    destCompanyName: '送信先会社',
    destContactPerson: '送信先太郎',
    destPhoneNumber: '03-8765-4321',
    destEmailAddress: 'dest@test.com',
    contractors: [
      { type: '施工業者', name: 'テスト施工' },
      { type: '塗装業者', name: 'テスト塗装' }
    ],
    products: [
      { productName: 'テスト商品1', quantity: '10', lotNumber: 'LOT001', shipmentDate: '2025-01-31' },
      { productName: 'テスト商品2', quantity: '20', lotNumber: 'LOT002', shipmentDate: '2025-01-31' }
    ],
    documents: ['成績書', 'SDS', '出荷証明書']
  };
  
  // doPostをシミュレート
  const e = {
    postData: {
      contents: JSON.stringify(testData)
    }
  };
  
  const result = doPost(e);
  console.log('テスト結果:', result.getContent());
}