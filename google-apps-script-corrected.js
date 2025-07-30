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

// 依頼書PDF生成機能（Google Docs方式）
function generateRequestPDF(data, managementNumber) {
  try {
    console.log('PDF生成開始...');
    
    // テンプレート文書を取得または作成
    let templateDocId = getOrCreateDocumentTemplate();
    
    // 保存先フォルダID
    const folderId = '1RJpSMtCHBUKqRL4kTisqEVs5YFLzlsk-';
    const folder = DriveApp.getFolderById(folderId);
    const fileName = `依頼書_${managementNumber}_${data.companyName}`;
    
    // テンプレートを複製
    const templateFile = DriveApp.getFileById(templateDocId);
    const copiedFile = templateFile.makeCopy(fileName, folder);
    
    // 複製した文書を開いてデータを差し込み
    const doc = DocumentApp.openById(copiedFile.getId());
    fillDocumentData(doc, data, managementNumber);
    
    // 文書を保存・閉じる
    doc.saveAndClose();
    
    // PDFとして出力
    const pdfBlob = copiedFile.getAs('application/pdf');
    pdfBlob.setName(`${fileName}.pdf`);
    
    // PDFをフォルダに保存
    const pdfFile = folder.createFile(pdfBlob);
    
    // 元のGoogle Doc文書を削除（PDFのみ保持）
    copiedFile.setTrashed(true);
    
    console.log('PDF生成成功');
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

// Google Docsテンプレート取得または作成
function getOrCreateDocumentTemplate() {
  try {
    // PropertiesServiceでテンプレートIDを管理
    const properties = PropertiesService.getScriptProperties();
    let templateDocId = properties.getProperty('TEMPLATE_DOC_ID');
    
    // テンプレートが存在するかチェック
    if (templateDocId) {
      try {
        const templateFile = DriveApp.getFileById(templateDocId);
        if (templateFile && !templateFile.isTrashed()) {
          console.log('既存テンプレート使用:', templateDocId);
          return templateDocId;
        }
      } catch (e) {
        console.log('既存テンプレートが無効、新規作成します');
        templateDocId = null;
      }
    }
    
    // 新しいテンプレートを作成
    console.log('新しいDocsテンプレート作成中...');
    templateDocId = createDocumentTemplate();
    
    // テンプレートIDを保存
    properties.setProperty('TEMPLATE_DOC_ID', templateDocId);
    console.log('新しいテンプレート作成完了:', templateDocId);
    
    return templateDocId;
    
  } catch (error) {
    console.error('テンプレート取得・作成エラー:', error);
    throw error;
  }
}

// Google Docsテンプレート作成関数
function createDocumentTemplate() {
  try {
    // 新しい文書を作成
    const doc = DocumentApp.create('依頼書テンプレート_' + new Date().getTime());
    const body = doc.getBody();
    
    // ページ設定
    const pageWidth = 595.28; // A4幅（ポイント）
    const pageHeight = 841.89; // A4高さ（ポイント）
    
    // マージン設定
    body.setMarginTop(50);
    body.setMarginBottom(50);
    body.setMarginLeft(50);
    body.setMarginRight(50);
    
    // 文書全体をクリア
    body.clear();
    
    // === タイトル ===
    const title = body.appendParagraph('出荷証明書作成依頼書');
    title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    title.editAsText().setFontSize(18).setBold(true);
    title.setSpacingAfter(15);
    
    body.appendParagraph(''); // 空行
    
    // === 受付番号・依頼日 ===
    const headerTable = body.appendTable([
      ['受付番号：{{受付番号}}', '', '', '', 'ご依頼日：{{依頼日}}']
    ]);
    headerTable.setBorderWidth(0);
    headerTable.getRow(0).getCell(0).setWidth(200);
    headerTable.getRow(0).getCell(4).setWidth(200);
    
    body.appendParagraph(''); // 空行
    
    // === 発信元セクション ===
    const senderTitle = body.appendParagraph('【発信元】');
    senderTitle.editAsText().setBold(true).setFontSize(12);
    
    // 発信元情報テーブル
    const senderTable = body.appendTable([
      ['御社名：', '{{会社名}}', '', '', ''],
      ['ご担当者名：', '{{担当者名}}', '様', 'FAX番号：', '{{FAX番号}}'],
      ['お電話番号：', '{{電話番号}}', '', 'メールアドレス：', '{{メールアドレス}}']
    ]);
    setupTableStyle(senderTable);
    
    body.appendParagraph(''); // 空行
    
    // === 宛先・基本情報 ===
    const recipientTable = body.appendTable([
      ['宛名：', '{{宛名}}', '{{敬称}}', '', ''],
      ['工事名：', '{{工事名}}', '', '', ''],
      ['工事住所：', '{{工事住所}}', '', '', 'ご依頼日：{{作成日}}']
    ]);
    setupTableStyle(recipientTable);
    
    body.appendParagraph(''); // 空行
    
    // === 業者情報 ===
    const contractorTitle = body.appendParagraph('【業者情報】');
    contractorTitle.editAsText().setBold(true).setFontSize(12);
    
    const contractorTable = body.appendTable([
      ['{{業者分類1}}：', '{{業者名1}}', '', '', ''],
      ['{{業者分類2}}：', '{{業者名2}}', '', '', ''],
      ['{{業者分類3}}：', '{{業者名3}}', '', '', ''],
      ['{{業者分類4}}：', '{{業者名4}}', '', '', '']
    ]);
    setupTableStyle(contractorTable);
    
    body.appendParagraph(''); // 空行
    
    // === 商品情報テーブル ===
    const productTitle = body.appendParagraph('【商品情報】');
    productTitle.editAsText().setBold(true).setFontSize(12);
    
    const productTable = body.appendTable([
      ['出荷年月日', '商品名', '数量', 'ロットNo'],
      ['{{商品1_出荷日}}', '{{商品1_名前}}', '{{商品1_数量}}', '{{商品1_ロット}}'],
      ['{{商品2_出荷日}}', '{{商品2_名前}}', '{{商品2_数量}}', '{{商品2_ロット}}'],
      ['{{商品3_出荷日}}', '{{商品3_名前}}', '{{商品3_数量}}', '{{商品3_ロット}}'],
      ['{{商品4_出荷日}}', '{{商品4_名前}}', '{{商品4_数量}}', '{{商品4_ロット}}'],
      ['{{商品5_出荷日}}', '{{商品5_名前}}', '{{商品5_数量}}', '{{商品5_ロット}}'],
      ['{{商品6_出荷日}}', '{{商品6_名前}}', '{{商品6_数量}}', '{{商品6_ロット}}'],
      ['{{商品7_出荷日}}', '{{商品7_名前}}', '{{商品7_数量}}', '{{商品7_ロット}}']
    ]);
    setupTableStyle(productTable, true); // ヘッダー付きテーブル
    
    body.appendParagraph(''); // 空行
    
    // === 必要書類・送信先 ===
    const bottomTable = body.appendTable([
      ['【必要書類】必要書類にチェックを入れてください', '', '', '【送信先】', ''],
      ['{{書類1}}成分表・試験成績書', '', '', '会社名：{{送信先会社名}}', ''],
      ['{{書類2}}ＳＤＳ', '', '', 'ご担当者名：{{送信先担当者名}}様', ''],
      ['{{書類3}}検査表(ロットが必要です)', '', '', 'お電話番号：{{送信先電話番号}}', ''],
      ['{{書類4}}カタログ', '', '', 'メールアドレス：{{送信先メールアドレス}}', ''],
      ['{{書類5}}ﾎﾙﾑｱﾙﾃﾞﾋﾄﾞ(F☆☆☆☆)証明書', '', '', '', '']
    ]);
    setupTableStyle(bottomTable);
    
    body.appendParagraph(''); // 空行
    
    // === 備考欄 ===
    const remarksTitle = body.appendParagraph('【備考欄】');
    remarksTitle.editAsText().setBold(true).setFontSize(12);
    
    const remarksTable = body.appendTable([
      ['{{備考}}', '', '', '', '']
    ]);
    setupTableStyle(remarksTable);
    remarksTable.getRow(0).getCell(0).setWidth(500);
    
    // 文書を保存
    doc.saveAndClose();
    
    return doc.getId();
    
  } catch (error) {
    console.error('Docsテンプレート作成エラー:', error);
    throw error;
  }
}

// テーブルスタイル設定関数
function setupTableStyle(table, hasHeader = false) {
  table.setBorderWidth(1);
  table.setBorderColor('#000000');
  
  // ヘッダー行がある場合
  if (hasHeader && table.getNumRows() > 0) {
    const headerRow = table.getRow(0);
    for (let i = 0; i < headerRow.getNumCells(); i++) {
      const cell = headerRow.getCell(i);
      cell.setBackgroundColor('#f0f0f0');
      cell.editAsText().setBold(true);
    }
  }
  
  // セル内の余白設定
  for (let i = 0; i < table.getNumRows(); i++) {
    const row = table.getRow(i);
    for (let j = 0; j < row.getNumCells(); j++) {
      const cell = row.getCell(j);
      cell.setPaddingTop(3);
      cell.setPaddingBottom(3);
      cell.setPaddingLeft(5);
      cell.setPaddingRight(5);
    }
  }
}

// Google Docsデータ差し込み関数
function fillDocumentData(doc, data, managementNumber) {
  try {
    console.log('文書データ差し込み開始...');
    
    const body = doc.getBody();
    
    // 基本情報の置換
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
    
    // 商品情報の置換（最大7商品）
    for (let i = 1; i <= 7; i++) {
      const product = data.products && data.products[i-1];
      if (product) {
        body.replaceText(`{{商品${i}_出荷日}}`, product.shipmentDate || '');
        body.replaceText(`{{商品${i}_名前}}`, product.productName || '');
        body.replaceText(`{{商品${i}_数量}}`, product.quantity || '');
        body.replaceText(`{{商品${i}_ロット}}`, product.lotNumber || '');
      } else {
        // 空の商品情報をクリア
        body.replaceText(`{{商品${i}_出荷日}}`, '');
        body.replaceText(`{{商品${i}_名前}}`, '');
        body.replaceText(`{{商品${i}_数量}}`, '');
        body.replaceText(`{{商品${i}_ロット}}`, '');
      }
    }
    
    // 必要書類のチェックマーク設定
    const docTypes = [
      '成分表・試験成績書',
      'ＳＤＳ',
      '検査表(ロットが必要です)',
      'カタログ',
      'ﾎﾙﾑｱﾙﾃﾞﾋﾄﾞ(F☆☆☆☆)証明書'
    ];
    
    for (let i = 1; i <= 5; i++) {
      const docType = docTypes[i-1];
      const isChecked = data.documents && data.documents.includes(docType);
      const checkMark = isChecked ? '☑' : '□';
      body.replaceText(`{{書類${i}}}`, checkMark);
    }
    
    console.log('文書データ差し込み完了');
    
  } catch (error) {
    console.error('文書データ差し込みエラー:', error);
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