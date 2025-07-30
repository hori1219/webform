function doPost(e) {
  try {
    // POST データを解析
    const data = JSON.parse(e.postData.contents);
    
    // スプレッドシートにアクセス
    const spreadsheetId = '1Ern5sRPYnwaxJB7xfQUQWsQi6I4ZHjsE9qDrHx6cLko'; // 実際のスプレッドシートIDに変更
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getActiveSheet();
    
    // データを行にマッピング（M列に工事住所が追加されたため、M列以降を調整）
    const row = [];
    
    // A-BT列に対応
    row[0] = '';  // A列: 発行チェック（空）
    row[1] = '';  // B列: 管理番号（空）
    row[2] = data.timestamp || new Date().toLocaleString('ja-JP'); // C列: 申請日時
    row[3] = '申請済';  // D列: ステータス
    row[4] = '1';  // E列: 版数
    row[5] = data.companyName || '';  // F列: 会社名
    row[6] = data.contactPerson || '';  // G列: 申請者名
    row[7] = data.phoneNumber || '';  // H列: 申請者TEL
    row[8] = data.emailAddress || '';  // I列: 依頼元メールアドレス
    row[9] = data.addressee || '';  // J列: 宛名
    row[10] = data.honorific || '';  // K列: 敬称
    row[11] = data.constructionName || '';  // L列: 工事名
    row[12] = data.constructionAddress || '';  // M列: 工事住所（新規追加）
    row[13] = data.creationDate || '';  // N列: 作成日
    
    // 業者情報（O-V列）
    for (let i = 0; i < 4; i++) {
      if (i < data.contractors.length) {
        row[14 + i * 2] = data.contractors[i].type || '';  // 業者分類
        row[15 + i * 2] = data.contractors[i].name || '';  // 業者名
      } else {
        row[14 + i * 2] = '';  // 業者分類
        row[15 + i * 2] = '';  // 業者名
      }
    }
    
    // 商品情報（W-AX列）- 最大7商品
    for (let i = 0; i < 7; i++) {
      if (i < data.products.length) {
        row[22 + i * 4] = data.products[i].productName || '';  // 商品名
        row[23 + i * 4] = data.products[i].quantity || '';    // 数量
        row[24 + i * 4] = data.products[i].lotNumber || '';   // ロットNo.
        row[25 + i * 4] = data.products[i].shippingDate || ''; // 出荷年月日
      } else {
        row[22 + i * 4] = '';  // 商品名
        row[23 + i * 4] = '';  // 数量
        row[24 + i * 4] = '';  // ロットNo.
        row[25 + i * 4] = '';  // 出荷年月日
      }
    }
    
    // 必要書類（AY-BC列）
    const docTypes = [
      '成分表・試験成績書',  // AY列
      'SDS',                 // AZ列
      '検査表',              // BA列
      'カタログ',            // BB列
      'ホルムアルデヒド証明書' // BC列
    ];
    
    for (let i = 0; i < docTypes.length; i++) {
      if (data.documents && data.documents.includes(docTypes[i])) {
        row[50 + i] = '○';
      } else {
        row[50 + i] = '';
      }
    }
    
    // 送付先情報（BD-BH列）
    row[55] = data.destCompanyName || '';     // BD列: 送付先会社名
    row[56] = data.destContactPerson || '';   // BE列: 送付先担当者名
    row[57] = data.destPhoneNumber || '';     // BF列: 送付先TEL
    row[58] = data.destEmailAddress || '';    // BG列: 送付先メールアドレス
    row[59] = '';  // BH列: 送付先メールアドレス(Cc)
    row[60] = '';  // BI列: 送付先メールアドレス(Bcc)
    row[61] = '';  // BJ列: 依頼書PDFリンク
    row[62] = '';  // BK列: 出荷証明書PDFリンク
    row[63] = '';  // BL列: 客先メールアドレス
    row[64] = new Date().toLocaleString('ja-JP');  // BM列: 最終更新日時
    row[65] = '';  // BN列: 社内通知先メール
    row[66] = '';  // BO列: テンプレ版
    row[67] = '';  // BP列: 証明書PDF fileId
    row[68] = '';  // BQ列: 証明書PDF作成日時
    row[69] = '';  // BR列: データハッシュ
    row[70] = '';  // BS列: 発行抑止
    row[71] = '';  // BT列: 送付日時（外部）
    row[72] = '';  // BU列: 送付結果
    
    // シートにデータを追加
    sheet.appendRow(row);
    
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'success',
        message: 'データが正常に保存されました'
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    console.error('エラー:', error);
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error',
        message: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput('Hello World');
}

// テスト用関数
function testFormSubmission() {
  const testData = {
    timestamp: new Date().toLocaleString('ja-JP'),
    companyName: 'テスト株式会社',
    contactPerson: '山田太郎',
    phoneNumber: '03-1234-5678',
    emailAddress: 'test@example.com',
    addressee: 'サンプル建設',
    honorific: '御中',
    constructionName: 'テスト工事',
    constructionAddress: '東京都千代田区1-1-1',
    creationDate: '2024-01-15',
    contractors: [
      { type: '元請', name: 'ABC建設' },
      { type: '下請', name: 'XYZ工業' }
    ],
    products: [
      { 
        productName: 'ローバル１ｋｇ',
        quantity: '10',
        lotNumber: 'LOT123',
        shippingDate: '2024-01-20'
      }
    ],
    documents: ['成分表・試験成績書', 'ＳＤＳ'],
    destCompanyName: '送信先会社',
    destContactPerson: '鈴木花子',
    destPhoneNumber: '06-5678-9012',
    destEmailAddress: 'dest@example.com'
  };
  
  const mockEvent = {
    postData: {
      contents: JSON.stringify(testData)
    }
  };
  
  const result = doPost(mockEvent);
  console.log(result.getContent());
}