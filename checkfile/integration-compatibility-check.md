# 統合互換性チェック結果

## 🔍 最新コード分析 (T2: 出荷証明書作成フロー)

### ✅ 確認済み定数
- `DB_ID`: '1tw3L-PQpr2D4o9GMISCQkfEMfR2aqr8aQtYcsqGptYY' ✓
- `DB_NAME`: 'シート1' ✓
- 既存の管理番号採番関数: `generateControlIdOptimized_()` ✓

### 🔄 新しいデプロイURL
- **更新前**: `AKfycbyz9LBPf7CpZsNHygIN-lCw9MrJsqDcsjkA_7x-2TxojzpjXiKsJlDYd0ZSpzGtU5v75Q`
- **更新後**: `AKfycbwzKN2ZeLMYEhzNlPy1ZWRIIG7W95qHDjPUV8Ev8RxnKPY9HkfbuDId2hFaduZv3_y5`

### ⚠️ 統合時の注意点

#### 1. 関数名の重複回避
既存コードには以下の関数があります：
- `generateControlIdOptimized_()` - 管理番号採番
- `setRow()` - ヘッダー名で行更新
- `readSheet()` - シートデータ読み取り

#### 2. 推奨統合方法
```javascript
// =============== 既存コード（削除しない） ===============
// （T2の全コードをそのまま保持）

// =============== 新規追加：Webフォーム連携 ===============
/**
 * Webフォーム専用のPOST処理
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    console.log('Webフォーム受信:', data);
    
    // 既存の定数・関数を活用
    const dbSheet = SpreadsheetApp.openById(DB_ID).getSheetByName(DB_NAME);
    const controlId = generateControlIdOptimized_(dbSheet);
    
    // Webフォーム専用データ構築
    const rowData = buildWebFormData_(controlId, data);
    
    dbSheet.appendRow(rowData);
    formatWebFormRow_(dbSheet);
    
    return ContentService
      .createTextOutput(JSON.stringify({
        result: 'success',
        controlId: controlId,
        message: '出荷証明書作成依頼が正常に登録されました'
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    console.error('Webフォームエラー:', error);
    return ContentService
      .createTextOutput(JSON.stringify({
        result: 'error',
        message: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Webフォーム用データ構築（既存形式に合わせる）
 */
function buildWebFormData_(controlId, data) {
  const rowData = [
    false,                          // A: チェックボックス
    controlId,                      // B: 管理番号
    new Date(),                     // C: 申請日時
    '申請中',                       // D: ステータス
    1,                             // E: 版数
    data.companyName || '',         // F: 会社名
    data.contactPerson || '',       // G: 申請者名
    data.phoneNumber || '',         // H: 申請者TEL
    data.faxNumber || '',           // I: FAX
    data.addressee || '',           // J: 宛名
    data.honorific || '様',         // K: 敬称
    data.constructionName || '',    // L: 工事名
    data.constructionAddress || '', // M: 工事住所
    data.creationDate || ''         // N: 作成日
  ];

  // 業者情報（3業者分）
  const contractors = data.contractors || [];
  for (let i = 0; i < 3; i++) {
    if (contractors[i]) {
      rowData.push(contractors[i].type || '');  // 業者分類
      rowData.push(contractors[i].name || '');  // 業者名
    } else {
      rowData.push('', '');
    }
  }

  // 商品情報（7商品分）
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

  // 送信先・書類情報
  const documents = data.documents || [];
  rowData.push(documents.join(', '));         // 必要書類
  rowData.push(data.destEmailAddress || '');  // 客先メールアドレス
  
  // 残りの列を空で埋める（中央DBの列構造に合わせる）
  while (rowData.length < 100) {
    rowData.push('');
  }

  return rowData;
}

/**
 * Webフォーム用行フォーマット
 */
function formatWebFormRow_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    // 基本的なフォーマット設定
    const range = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn());
    if (lastRow % 2 === 0) {
      range.setBackground('#F8F9FA');
    }
    range.setBorder(true, true, true, true, true, true);
    
    // 日付フォーマット
    sheet.getRange(lastRow, 3).setNumberFormat('yyyy/mm/dd hh:mm:ss');
  }
}
```

### 🎯 統合後の動作確認項目
1. **Webフォーム送信** → 中央DBに正しく追記
2. **管理番号採番** → 既存ルールで正常採番
3. **既存メニュー** → 「申請中を処理」で正常動作
4. **PDF生成** → 出荷証明書PDF正常作成

### 📝 重要な注意点
- **既存コードは絶対に削除しない**
- **関数名は重複しないように別名を使用**
- **既存の定数・関数を最大限活用**
- **データ構造は中央DBの列構造に完全準拠**

## ✅ 問題なし
更新されたコードは既存システムとの互換性に問題ありません。安全に統合可能です。