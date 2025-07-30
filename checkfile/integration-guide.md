# Google Apps Script統合ガイド

## 🔗 URL更新完了
- **新しいWebアプリURL**: `https://script.google.com/macros/s/AKfycbyz9LBPf7CpZsNHygIN-lCw9MrJsqDcsjkA_7x-2TxojzpjXiKsJlDYd0ZSpzGtU5v75Q/exec`
- **index.html**: 更新済み ✅

## 📋 Google Apps Script統合手順

### 1. 既存コードの保持
既存のCode.gsには以下の重要な関数があります：
- `submitCertificateRequest()` - 出荷証明依頼処理
- `processPendingCertificates()` - 申請中を処理  
- `generateControlIdOptimized_()` - 管理番号採番
- `writeByHeader_()` - ヘッダー名での書き込み

### 2. 新規追加するコード
既存のCode.gsの**最後に**以下を追加してください：

```javascript
// =============== 新規追加：Webフォーム連携 ===============
/**
 * POST リクエストを処理する関数
 * Webフォームからのデータをスプレッドシートに追加
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    console.log('Webフォーム受信データ:', data);
    
    // 既存の定数を使用
    const spreadsheet = SpreadsheetApp.openById(DB_ID);
    let sheet = spreadsheet.getSheetByName(DB_NAME);
    
    if (!sheet) {
      throw new Error(`シート "${DB_NAME}" が見つかりません`);
    }
    
    // 既存の関数を活用して管理番号採番
    const controlId = generateControlIdOptimized_(sheet);
    
    // Webフォーム用データ構築
    const rowData = buildWebFormRowData_(controlId, data);
    
    sheet.appendRow(rowData);
    formatLatestRowWeb_(sheet);
    
    console.log('Webフォームデータ追加成功:', controlId);
    
    return ContentService
      .createTextOutput(JSON.stringify({
        result: 'success',
        message: '出荷証明書作成依頼が正常に登録されました',
        controlId: controlId,
        timestamp: new Date().toISOString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    console.error('Webフォームエラー:', error);
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
 * Webフォーム専用のデータ構築関数
 */
function buildWebFormRowData_(controlId, data) {
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

  // 必要書類・送信先情報など
  const documents = data.documents || [];
  rowData.push(documents.join(', '));              // 書類リスト
  rowData.push(data.destEmailAddress || '');       // 客先メールアドレス
  
  // 残りの列を空で埋める（既存構造に合わせる）
  while (rowData.length < 50) { // 適宜調整
    rowData.push('');
  }
  
  // 最終更新日時
  rowData.push(new Date());

  return rowData;
}

/**
 * Webフォーム用フォーマット関数
 */
function formatLatestRowWeb_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const range = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn());
    
    // 交互の背景色設定
    if (lastRow % 2 === 0) {
      range.setBackground('#F8F9FA');
    }
    
    // 境界線の設定
    range.setBorder(true, true, true, true, true, true);
    
    // 日付フォーマット
    sheet.getRange(lastRow, 3).setNumberFormat('yyyy/mm/dd hh:mm:ss'); // 申請日時
  }
}
```

### 3. 注意事項
- **既存コードは削除しない**
- **定数名を既存に合わせる** (`DB_ID`, `DB_NAME`を使用)
- **関数名の重複を避ける** (Web用に別名を使用)

### 4. テスト方法
1. コード追加後、保存
2. Webフォームから送信テスト
3. スプレッドシートに正しく追記されるか確認
4. ログでエラーがないか確認

## 🚨 トラブルシューティング
- **権限エラー**: Apps Scriptの実行権限を確認
- **データ形式エラー**: コンソールログでデータ構造を確認
- **列数不一致**: rowDataの要素数をスプレッドシートの列数に合わせる