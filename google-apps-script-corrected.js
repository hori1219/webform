function doPost(e) {
  try {
    // POST データを解析
    const data = JSON.parse(e.postData.contents);
    
    // スプレッドシートにアクセス
    const spreadsheetId = '1xfFlHJihYyhJ-CKP3Aj5veN9c9lanolsTj4kyvR_9R0'; // 実際のスプレッドシートID
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getActiveSheet();
    
    // 管理番号を生成
    const managementNumber = generateManagementNumber(sheet);
    
    // データを行にマッピング（M列に工事住所が追加されたため、M列以降を調整）
    const row = [];
    
    // A-BT列に対応
    row[0] = '';  // A列: 発行チェック（空）
    row[1] = managementNumber;  // B列: 管理番号
    row[2] = data.timestamp || new Date().toLocaleString('ja-JP'); // C列: 申請日時
    row[3] = '依頼受付';  // D列: ステータス
    row[4] = '01';  // E列: 版数
    row[5] = data.companyName || '';  // F列: 会社名
    row[6] = data.contactPerson || '';  // G列: 申請者名
    row[7] = "'" + (data.phoneNumber || '');  // H列: 申請者TEL（テキスト形式）
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
        row[25 + i * 4] = data.products[i].shipmentDate || ''; // 出荷年月日
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
      'ＳＤＳ',                 // AZ列
      '検査表(ロットが必要です)',              // BA列
      'カタログ',            // BB列
      'ﾎﾙﾑｱﾙﾃﾞﾋﾄﾞ(F☆☆☆☆)証明書' // BC列
    ];
    
    for (let i = 0; i < docTypes.length; i++) {
      if (data.documents && data.documents.includes(docTypes[i])) {
        row[50 + i] = 1;  // 数値の1でフラグ
      } else {
        row[50 + i] = 0;  // 数値の0
      }
    }
    
    // 送付先情報（BD-BH列）
    row[55] = data.destCompanyName || '';     // BD列: 送付先会社名
    row[56] = data.destContactPerson || '';   // BE列: 送付先担当者名
    row[57] = "'" + (data.destPhoneNumber || '');     // BF列: 送付先TEL（テキスト形式）
    row[58] = data.destEmailAddress || '';    // BG列: 送付先メールアドレス
    row[59] = '';  // BH列: 送付先メールアドレス(Cc)
    row[60] = '';  // BI列: 送付先メールアドレス(Bcc)
    row[61] = data.remarks || '';  // BJ列: 備考
    row[62] = '';  // BK列: 依頼書PDFリンク
    row[63] = '';  // BL列: 出荷証明書PDFリンク
    row[64] = '';  // BM列: 客先メールアドレス
    row[65] = new Date().toLocaleString('ja-JP');  // BN列: 最終更新日時
    row[66] = '';  // BO列: 社内通知先メール
    row[67] = '';  // BP列: テンプレ版
    row[68] = '';  // BQ列: 証明書PDF fileId
    row[69] = '';  // BR列: 証明書PDF作成日時
    row[70] = '';  // BS列: データハッシュ
    row[71] = '';  // BT列: 発行抑止
    row[72] = '';  // BU列: 送付日時（外部）
    row[73] = '';  // BV列: 送付結果
    
    // シートにデータを追加
    sheet.appendRow(row);
    
    // 依頼書PDF生成とメール送信
    try {
      const pdfResult = generateRequestPDF(data, managementNumber);
      const emailResult = sendRequestEmail(data, pdfResult.pdfBlob, managementNumber);
      
      // PDFリンクをスプレッドシートに更新
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow, 62).setValue(pdfResult.pdfUrl); // BK列: 依頼書PDFリンク
      
    } catch (pdfError) {
      console.error('PDF生成・メール送信エラー:', pdfError);
      // PDF生成エラーでもデータ保存は成功とする
    }
    
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'success',
        message: 'データが正常に保存され、依頼書PDFが送信されました',
        managementNumber: managementNumber
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

// 管理番号を生成する関数（LockService対応）
function generateManagementNumber(sheet) {
  const lock = LockService.getScriptLock();
  
  try {
    // 10秒間ロックを試行
    lock.waitLock(10000);
    
    // 最後の行を取得
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      // データが1行目のヘッダーのみの場合、初回番号を返す
      return '000001-01';
    }
    
    // B列（管理番号列）の最後の値を取得
    const lastManagementNumber = sheet.getRange(lastRow, 2).getValue();
    
    if (!lastManagementNumber) {
      // 管理番号が空の場合、初回番号を返す
      return '000001-01';
    }
    
    // 管理番号を分解（例: "000001-01" → ["000001", "01"]）
    const parts = lastManagementNumber.toString().split('-');
    if (parts.length !== 2) {
      // フォーマットが正しくない場合、初回番号を返す
      return '000001-01';
    }
    
    const serialNumber = parseInt(parts[0]);
    
    // 次の連番を生成
    const nextSerialNumber = serialNumber + 1;
    const formattedSerialNumber = nextSerialNumber.toString().padStart(6, '0');
    
    return `${formattedSerialNumber}-01`;
    
  } catch (error) {
    console.error('管理番号生成エラー:', error);
    // エラーの場合は現在時刻ベースの番号を生成
    const now = new Date();
    const timeBasedNumber = (now.getTime() % 1000000).toString().padStart(6, '0');
    return `${timeBasedNumber}-01`;
  } finally {
    // 必ずロックを解放
    lock.releaseLock();
  }
}

// 依頼書PDF生成機能（Spreadsheet方式）
function generateRequestPDF(data, managementNumber) {
  try {
    // 現在のスプレッドシートを取得
    const spreadsheetId = '1xfFlHJihYyhJ-CKP3Aj5veN9c9lanolsTj4kyvR_9R0';
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    
    // テンプレートシートを取得（事前に作成が必要）
    let templateSheet;
    try {
      templateSheet = spreadsheet.getSheetByName('依頼書テンプレート');
    } catch (e) {
      // テンプレートシートが存在しない場合は作成
      templateSheet = createTemplateSheet(spreadsheet);
    }
    
    // データをテンプレートに転記
    fillTemplateData(templateSheet, data, managementNumber);
    
    // 保存先フォルダID
    const folderId = '1RJpSMtCHBUKqRL4kTisqEVs5YFLzlsk-';
    const folder = DriveApp.getFolderById(folderId);
    const fileName = `依頼書_${managementNumber}_${data.companyName}`;
    
    // PDF生成のための印刷設定
    const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?` +
                `format=pdf&gid=${templateSheet.getSheetId()}` +
                `&range=A1:I40` +
                `&size=A4` +
                `&portrait=true` +
                `&fitw=true` +
                `&top_margin=0.5` +
                `&bottom_margin=0.5` +
                `&left_margin=0.5` +
                `&right_margin=0.5`;
    
    // PDF生成
    const response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      }
    });
    
    const pdfBlob = response.getBlob();
    pdfBlob.setName(`${fileName}.pdf`);
    
    // PDFをフォルダに保存
    const pdfFile = folder.createFile(pdfBlob);
    
    // テンプレートシートのデータをクリア
    clearTemplateData(templateSheet);
    
    return {
      pdfBlob: pdfBlob,
      pdfUrl: pdfFile.getUrl(),
      pdfId: pdfFile.getId()
    };
    
  } catch (error) {
    console.error('PDF生成エラー:', error);
    throw error;
  }
}

// メール送信機能
function sendRequestEmail(data, pdfBlob, managementNumber) {
  try {
    const subject = `【依頼書送付】出荷証明書作成依頼書_${managementNumber}`;
    
    const emailBody = `
${data.addressee} ${data.honorific}

いつもお世話になっております。

出荷証明書作成依頼書をお送りいたします。

■ 管理番号：${managementNumber}
■ 工事名：${data.constructionName}
■ 申請者：${data.companyName} ${data.contactPerson}様

添付の依頼書をご確認の上、出荷証明書の作成をお願いいたします。

ご質問等ございましたら、下記までご連絡ください。

---
${data.companyName}
${data.contactPerson}
TEL: ${data.phoneNumber}
Email: ${data.emailAddress}
---

※このメールは自動送信されています。
`;

    // メール送信
    GmailApp.sendEmail(
      data.destEmailAddress,
      subject,
      emailBody,
      {
        attachments: [pdfBlob],
        name: data.companyName // 送信者名
      }
    );
    
    return {
      status: 'success',
      sentTo: data.destEmailAddress,
      sentAt: new Date().toLocaleString('ja-JP')
    };
    
  } catch (error) {
    console.error('メール送信エラー:', error);
    throw error;
  }
}

// テンプレートシート作成関数
function createTemplateSheet(spreadsheet) {
  const templateSheet = spreadsheet.insertSheet('依頼書テンプレート');
  
  // シート設定
  templateSheet.setColumnWidths(1, 9, 100); // A-I列の幅を100pxに設定
  templateSheet.setRowHeights(1, 40, 25);   // 行高を25pxに設定
  
  // タイトル設定
  templateSheet.getRange('A1:I1').merge();
  templateSheet.getRange('A1').setValue('出荷証明書作成依頼書');
  templateSheet.getRange('A1').setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
  
  // セクション見出しと枠組みを設定
  setupTemplateLayout(templateSheet);
  
  return templateSheet;
}

// テンプレートレイアウト設定
function setupTemplateLayout(sheet) {
  // 受付番号・依頼日（3行目）
  sheet.getRange('A3').setValue('受付番号：');
  sheet.getRange('G3').setValue('ご依頼日：');
  
  // 発信元セクション（5-8行目）
  sheet.getRange('A5').setValue('【発信元】').setFontWeight('bold');
  sheet.getRange('A6').setValue('御社名：');
  sheet.getRange('A7').setValue('ご担当者名：');
  sheet.getRange('D7').setValue('様');
  sheet.getRange('F7').setValue('FAX番号：');
  sheet.getRange('A8').setValue('お電話番号：');
  sheet.getRange('F8').setValue('メールアドレス：');
  
  // 基本情報セクション（10-12行目）
  sheet.getRange('A10').setValue('宛名：');
  sheet.getRange('A11').setValue('工事名：');
  sheet.getRange('A12').setValue('工事住所：');
  sheet.getRange('H12').setValue('作成日：');
  
  // 商品テーブルヘッダー（19行目）
  sheet.getRange('A19').setValue('出荷年月日').setFontWeight('bold');
  sheet.getRange('C19').setValue('商品名').setFontWeight('bold');
  sheet.getRange('F19').setValue('数量').setFontWeight('bold');
  sheet.getRange('G19').setValue('ロットNo').setFontWeight('bold');
  
  // 必要書類セクション（28-33行目）
  sheet.getRange('A28').setValue('【必要書類】必要書類にチェックを入れてください').setFontWeight('bold');
  sheet.getRange('A29').setValue('□ 成分表・試験成績書');
  sheet.getRange('A30').setValue('□ ＳＤＳ');
  sheet.getRange('A31').setValue('□ 検査表(ロットが必要です)');
  sheet.getRange('A32').setValue('□ カタログ');
  sheet.getRange('A33').setValue('□ ﾎﾙﾑｱﾙﾃﾞﾋﾄﾞ(F☆☆☆☆)証明書');
  
  // 送信先セクション（28-32行目、右側）
  sheet.getRange('F28').setValue('【送信先】').setFontWeight('bold');
  sheet.getRange('F29').setValue('会社名：');
  sheet.getRange('F30').setValue('ご担当者名：');
  sheet.getRange('I30').setValue('様');
  sheet.getRange('F31').setValue('お電話番号：');
  sheet.getRange('F32').setValue('メールアドレス：');
  
  // 備考欄（35-37行目）
  sheet.getRange('A35').setValue('【備考欄】').setFontWeight('bold');
  sheet.getRange('A36:I37').merge();
  
  // 罫線設定
  const range = sheet.getRange('A1:I40');
  range.setBorder(true, true, true, true, true, true);
}

// データ転記関数
function fillTemplateData(sheet, data, managementNumber) {
  // 基本情報
  sheet.getRange('B3').setValue(managementNumber);
  sheet.getRange('H3').setValue(data.timestamp || new Date().toLocaleString('ja-JP'));
  
  // 発信元情報
  sheet.getRange('B6:D6').merge().setValue(data.companyName || '');
  sheet.getRange('B7:C7').merge().setValue(data.contactPerson || '');
  sheet.getRange('B8:D8').merge().setValue(data.phoneNumber || '');
  sheet.getRange('G8:I8').merge().setValue(data.emailAddress || '');
  
  // 基本情報
  sheet.getRange('B10:C10').merge().setValue(data.addressee || '');
  sheet.getRange('D10').setValue(data.honorific || '');
  sheet.getRange('B11:G11').merge().setValue(data.constructionName || '');
  sheet.getRange('B12:G12').merge().setValue(data.constructionAddress || '');
  sheet.getRange('I12').setValue(data.creationDate || '');
  
  // 業者情報
  if (data.contractors && data.contractors.length > 0) {
    for (let i = 0; i < Math.min(data.contractors.length, 4); i++) {
      const row = 14 + i;
      sheet.getRange(`A${row}:B${row}`).merge().setValue(data.contractors[i].type || '');
      sheet.getRange(`C${row}:F${row}`).merge().setValue(data.contractors[i].name || '');
    }
  }
  
  // 商品情報
  if (data.products && data.products.length > 0) {
    for (let i = 0; i < Math.min(data.products.length, 7); i++) {
      const row = 20 + i;
      sheet.getRange(`A${row}:B${row}`).merge().setValue(data.products[i].shipmentDate || '');
      sheet.getRange(`C${row}:E${row}`).merge().setValue(data.products[i].productName || '');
      sheet.getRange(`F${row}`).setValue(data.products[i].quantity || '');
      sheet.getRange(`G${row}:I${row}`).merge().setValue(data.products[i].lotNumber || '');
    }
  }
  
  // 必要書類のチェックマーク
  const docChecks = {
    'A29': data.documents && data.documents.includes('成分表・試験成績書'),
    'A30': data.documents && data.documents.includes('ＳＤＳ'),
    'A31': data.documents && data.documents.includes('検査表(ロットが必要です)'),
    'A32': data.documents && data.documents.includes('カタログ'),
    'A33': data.documents && data.documents.includes('ﾎﾙﾑｱﾙﾃﾞﾋﾄﾞ(F☆☆☆☆)証明書')
  };
  
  Object.keys(docChecks).forEach(cell => {
    const currentValue = sheet.getRange(cell).getValue();
    const newValue = docChecks[cell] ? currentValue.replace('□', '☑') : currentValue;
    sheet.getRange(cell).setValue(newValue);
  });
  
  // 送信先情報
  sheet.getRange('G29:I29').merge().setValue(data.destCompanyName || '');
  sheet.getRange('G30:H30').merge().setValue(data.destContactPerson || '');
  sheet.getRange('G31:I31').merge().setValue(data.destPhoneNumber || '');
  sheet.getRange('G32:I32').merge().setValue(data.destEmailAddress || '');
  
  // 備考
  sheet.getRange('A36:I37').setValue(data.remarks || '');
}

// テンプレートデータクリア関数
function clearTemplateData(sheet) {
  // データ部分のみクリア（レイアウトは保持）
  const clearRanges = [
    'B3', 'H3', // 受付番号、依頼日
    'B6:D6', 'B7:C7', 'B8:D8', 'G8:I8', // 発信元
    'B10:C10', 'D10', 'B11:G11', 'B12:G12', 'I12', // 基本情報
    'A14:F17', // 業者情報
    'A20:I26', // 商品情報
    'G29:I32', // 送信先
    'A36:I37'  // 備考
  ];
  
  clearRanges.forEach(range => {
    sheet.getRange(range).clearContent();
  });
  
  // 必要書類のチェックマークをリセット
  const docCells = ['A29', 'A30', 'A31', 'A32', 'A33'];
  docCells.forEach((cell, index) => {
    const labels = [
      '□ 成分表・試験成績書',
      '□ ＳＤＳ', 
      '□ 検査表(ロットが必要です)',
      '□ カタログ',
      '□ ﾎﾙﾑｱﾙﾃﾞﾋﾄﾞ(F☆☆☆☆)証明書'
    ];
    sheet.getRange(cell).setValue(labels[index]);
  });
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