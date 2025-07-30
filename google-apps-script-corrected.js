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

// 依頼書PDF生成機能
function generateRequestPDF(data, managementNumber) {
  try {
    // テンプレートドキュメントID（作成後に設定）
    const templateDocId = 'YOUR_TEMPLATE_DOC_ID'; // 後で実際のIDに変更
    
    // 保存先フォルダID
    const folderId = '1RJpSMtCHBUKqRL4kTisqEVs5YFLzlsk-';
    
    // テンプレートを複製
    const templateDoc = DriveApp.getFileById(templateDocId);
    const folder = DriveApp.getFolderById(folderId);
    const fileName = `依頼書_${managementNumber}_${data.companyName}`;
    const copiedDoc = templateDoc.makeCopy(fileName, folder);
    
    // 複製した文書を開く
    const doc = DocumentApp.openById(copiedDoc.getId());
    const body = doc.getBody();
    
    // プレースホルダーを実際の値に置換
    const replacements = {
      '{{受付番号}}': managementNumber,
      '{{依頼日}}': data.timestamp || new Date().toLocaleString('ja-JP'),
      '{{会社名}}': data.companyName || '',
      '{{担当者名}}': data.contactPerson || '',
      '{{電話番号}}': data.phoneNumber || '',
      '{{FAX番号}}': '', // 固定値または空
      '{{メールアドレス}}': data.emailAddress || '',
      '{{宛名}}': data.addressee || '',
      '{{敬称}}': data.honorific || '',
      '{{工事名}}': data.constructionName || '',
      '{{工事住所}}': data.constructionAddress || '',
      '{{作成日}}': data.creationDate || '',
      '{{送信先会社名}}': data.destCompanyName || '',
      '{{送信先担当者名}}': data.destContactPerson || '',
      '{{送信先電話番号}}': data.destPhoneNumber || '',
      '{{送信先メールアドレス}}': data.destEmailAddress || '',
      '{{備考}}': data.remarks || ''
    };
    
    // 基本的な置換を実行
    Object.keys(replacements).forEach(placeholder => {
      body.replaceText(placeholder, replacements[placeholder]);
    });
    
    // 業者情報の置換（最大4社）
    for (let i = 1; i <= 4; i++) {
      const contractor = data.contractors && data.contractors[i-1];
      body.replaceText(`{{業者分類${i}}}`, contractor ? contractor.type : '');
      body.replaceText(`{{業者名${i}}}`, contractor ? contractor.name : '');
    }
    
    // 必要書類のチェックマークを設定
    const docChecks = {
      '{{成分表試験成績書}}': data.documents && data.documents.includes('成分表・試験成績書') ? '✓' : '',
      '{{SDS}}': data.documents && data.documents.includes('ＳＤＳ') ? '✓' : '',
      '{{検査表}}': data.documents && data.documents.includes('検査表(ロットが必要です)') ? '✓' : '',
      '{{カタログ}}': data.documents && data.documents.includes('カタログ') ? '✓' : '',
      '{{ホルムアルデヒド証明書}}': data.documents && data.documents.includes('ﾎﾙﾑｱﾙﾃﾞﾋﾄﾞ(F☆☆☆☆)証明書') ? '✓' : ''
    };
    
    Object.keys(docChecks).forEach(placeholder => {
      body.replaceText(placeholder, docChecks[placeholder]);
    });
    
    // 商品テーブルの生成（簡単な文字列として挿入）
    let productTable = '';
    if (data.products && data.products.length > 0) {
      data.products.forEach(product => {
        productTable += `${product.shipmentDate || ''} | ${product.productName || ''} | ${product.quantity || ''} | ${product.lotNumber || ''}\n`;
      });
    }
    body.replaceText('{{商品テーブル}}', productTable);
    
    // 文書を保存
    doc.saveAndClose();
    
    // PDFとして出力
    const pdfBlob = DriveApp.getFileById(copiedDoc.getId()).getAs('application/pdf');
    pdfBlob.setName(`${fileName}.pdf`);
    
    // PDFをフォルダに保存
    const pdfFile = folder.createFile(pdfBlob);
    
    // 元のGoogle Doc文書は削除（PDFのみ保持）
    DriveApp.getFileById(copiedDoc.getId()).setTrashed(true);
    
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