/**
 * 🧪 Phase3完全版 Google Apps Script
 * 
 * 全項目対応:
 * - 発信元＋基本情報: F28～O28 (10項目)
 * - 業者情報: P28～U28 (3業者×2列=6項目)
 * - 商品情報: V28～AO28 (7商品×4列=28項目)
 * - その他: AP28～AR28 (3項目)
 * 
 * 合計: F28～AR28 (47項目)
 */

// 設定
const TEST_SPREADSHEET_ID = '1xfFlHJihYyhJ-CKP3Aj5veN9c9lanolsTj4kyvR_9R0';
const TEST_SHEET_NAME = 'シート1';
const TEST_TARGET_ROW = 28;

// 完全列マッピング（1-based）
const FULL_COLUMN_MAP = {
    // 発信元 (F～J)
    companyName: 6,          // F: 会社名
    contactPerson: 7,        // G: 担当者名
    phoneNumber: 8,          // H: 電話番号
    faxNumber: 9,            // I: FAX
    emailAddress: 10,        // J: メールアドレス
    
    // 基本情報 (K～O)
    addressee: 11,           // K: 宛名
    honorific: 12,           // L: 敬称
    constructionName: 13,    // M: 工事名
    constructionAddress: 14, // N: 工事住所
    creationDate: 15,        // O: 作成日
    
    // 業者情報 (P～U: 3業者×2列)
    contractor1Type: 16,     // P: 業者分類1
    contractor1Name: 17,     // Q: 業者名1
    contractor2Type: 18,     // R: 業者分類2
    contractor2Name: 19,     // S: 業者名2
    contractor3Type: 20,     // T: 業者分類3
    contractor3Name: 21,     // U: 業者名3
    
    // 商品情報 (V～AO: 7商品×4列)
    // 商品1: V～Y
    product1Name: 22,        // V: 商品名1
    product1Quantity: 23,    // W: 数量1
    product1Lot: 24,         // X: ロット1
    product1Date: 25,        // Y: 出荷日1
    // 商品2: Z～AC
    product2Name: 26,        // Z: 商品名2
    product2Quantity: 27,    // AA: 数量2
    product2Lot: 28,         // AB: ロット2
    product2Date: 29,        // AC: 出荷日2
    // 商品3: AD～AG
    product3Name: 30,        // AD: 商品名3
    product3Quantity: 31,    // AE: 数量3
    product3Lot: 32,         // AF: ロット3
    product3Date: 33,        // AG: 出荷日3
    // 商品4: AH～AK
    product4Name: 34,        // AH: 商品名4
    product4Quantity: 35,    // AI: 数量4
    product4Lot: 36,         // AJ: ロット4
    product4Date: 37,        // AK: 出荷日4
    // 商品5: AL～AO
    product5Name: 38,        // AL: 商品名5
    product5Quantity: 39,    // AM: 数量5
    product5Lot: 40,         // AN: ロット5
    product5Date: 41,        // AO: 出荷日5
    // 商品6: AP～AS
    product6Name: 42,        // AP: 商品名6
    product6Quantity: 43,    // AQ: 数量6
    product6Lot: 44,         // AR: ロット6
    product6Date: 45,        // AS: 出荷日6
    // 商品7: AT～AW
    product7Name: 46,        // AT: 商品名7
    product7Quantity: 47,    // AU: 数量7
    product7Lot: 48,         // AV: ロット7
    product7Date: 49,        // AW: 出荷日7
    
    // その他情報 (AX～AZ)
    documents: 50,           // AX: 必要書類
    destEmailAddress: 51,    // AY: 送信先メール
    timestamp: 52            // AZ: 処理日時
};

/**
 * GET リクエスト処理
 */
function doGet(e) {
    return ContentService
        .createTextOutput('🧪 Phase3完全版API - 全項目統合転記システム')
        .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * POST リクエスト処理（Phase3完全版）
 */
function doPost(e) {
    try {
        console.log('=== 🧪 Phase3完全テスト開始 ===');
        
        // 1. データ受信
        if (!e || !e.postData || !e.postData.contents) {
            throw new Error('POSTデータが空です');
        }
        
        const data = JSON.parse(e.postData.contents);
        console.log('Phase3受信データ:', JSON.stringify(data, null, 2));
        
        // 2. 必須項目チェック
        const requiredFields = ['companyName', 'contactPerson', 'phoneNumber', 'emailAddress', 'addressee', 'constructionName', 'destEmailAddress'];
        const missingFields = requiredFields.filter(field => !data[field] || !data[field].trim());
        
        if (missingFields.length > 0) {
            throw new Error(`必須項目が未入力です: ${missingFields.join(', ')}`);
        }
        
        // 3. 商品必須チェック
        if (!data.products || data.products.length === 0) {
            throw new Error('商品情報が1つも入力されていません');
        }
        
        // 4. スプレッドシート接続
        console.log('スプレッドシート接続中...');
        const spreadsheet = SpreadsheetApp.openById(TEST_SPREADSHEET_ID);
        const sheet = spreadsheet.getSheetByName(TEST_SHEET_NAME);
        
        if (!sheet) {
            throw new Error(`シート "${TEST_SHEET_NAME}" が見つかりません`);
        }
        
        console.log('スプレッドシート接続成功');
        
        // 5. データ転記実行
        console.log('=== 全項目データ転記開始 ===');
        const timestamp = new Date().toLocaleString('ja-JP');
        const results = {};
        
        // 基本情報転記
        const basicFields = ['companyName', 'contactPerson', 'phoneNumber', 'faxNumber', 'emailAddress', 'addressee', 'honorific', 'constructionName', 'constructionAddress', 'creationDate'];
        basicFields.forEach(field => {
            if (FULL_COLUMN_MAP[field]) {
                let value = data[field] || '';
                if (field === 'honorific' && !value) value = '様';
                
                const writeValue = value ? `${value} (${timestamp})` : '';
                const col = FULL_COLUMN_MAP[field];
                
                console.log(`転記: ${field} → ${getColumnLetter(col)}${TEST_TARGET_ROW} = "${writeValue}"`);
                sheet.getRange(TEST_TARGET_ROW, col).setValue(writeValue);
                
                results[field] = { column: getColumnLetter(col), value: writeValue };
            }
        });
        
        // 業者情報転記（最大3業者）
        console.log('=== 業者情報転記 ===');
        const contractors = data.contractors || [];
        for (let i = 0; i < 3; i++) {
            const typeField = `contractor${i+1}Type`;
            const nameField = `contractor${i+1}Name`;
            
            if (FULL_COLUMN_MAP[typeField] && FULL_COLUMN_MAP[nameField]) {
                const contractor = contractors[i];
                const type = contractor ? contractor.type || '' : '';
                const name = contractor ? contractor.name || '' : '';
                
                const typeValue = type ? `${type} (${timestamp})` : '';
                const nameValue = name ? `${name} (${timestamp})` : '';
                
                const typeCol = FULL_COLUMN_MAP[typeField];
                const nameCol = FULL_COLUMN_MAP[nameField];
                
                console.log(`業者${i+1}: ${getColumnLetter(typeCol)}${TEST_TARGET_ROW}="${typeValue}", ${getColumnLetter(nameCol)}${TEST_TARGET_ROW}="${nameValue}"`);
                
                sheet.getRange(TEST_TARGET_ROW, typeCol).setValue(typeValue);
                sheet.getRange(TEST_TARGET_ROW, nameCol).setValue(nameValue);
                
                results[typeField] = { column: getColumnLetter(typeCol), value: typeValue };
                results[nameField] = { column: getColumnLetter(nameCol), value: nameValue };
            }
        }
        
        // 商品情報転記（最大7商品）
        console.log('=== 商品情報転記 ===');
        const products = data.products || [];
        for (let i = 0; i < 7; i++) {
            const nameField = `product${i+1}Name`;
            const quantityField = `product${i+1}Quantity`;
            const lotField = `product${i+1}Lot`;
            const dateField = `product${i+1}Date`;
            
            if (FULL_COLUMN_MAP[nameField]) {
                const product = products[i];
                const productName = product ? product.productName || '' : '';
                const quantity = product ? product.quantity || '' : '';
                const lotNumber = product ? product.lotNumber || '' : '';
                const shipmentDate = product ? product.shipmentDate || '' : '';
                
                const nameValue = productName ? `${productName} (${timestamp})` : '';
                const quantityValue = quantity ? `${quantity} (${timestamp})` : '';
                const lotValue = lotNumber ? `${lotNumber} (${timestamp})` : '';
                const dateValue = shipmentDate ? `${shipmentDate} (${timestamp})` : '';
                
                const nameCol = FULL_COLUMN_MAP[nameField];
                const quantityCol = FULL_COLUMN_MAP[quantityField];
                const lotCol = FULL_COLUMN_MAP[lotField];
                const dateCol = FULL_COLUMN_MAP[dateField];
                
                console.log(`商品${i+1}: ${getColumnLetter(nameCol)}="${nameValue}"`);
                
                sheet.getRange(TEST_TARGET_ROW, nameCol).setValue(nameValue);
                sheet.getRange(TEST_TARGET_ROW, quantityCol).setValue(quantityValue);
                sheet.getRange(TEST_TARGET_ROW, lotCol).setValue(lotValue);  
                sheet.getRange(TEST_TARGET_ROW, dateCol).setValue(dateValue);
                
                results[nameField] = { column: getColumnLetter(nameCol), value: nameValue };
                results[quantityField] = { column: getColumnLetter(quantityCol), value: quantityValue };
                results[lotField] = { column: getColumnLetter(lotCol), value: lotValue };
                results[dateField] = { column: getColumnLetter(dateCol), value: dateValue };
            }
        }
        
        // その他情報転記
        console.log('=== その他情報転記 ===');
        const documents = (data.documents || []).join(', ');
        const documentsValue = documents ? `${documents} (${timestamp})` : '';
        const destEmailValue = data.destEmailAddress ? `${data.destEmailAddress} (${timestamp})` : '';
        const timestampValue = `${timestamp}`;
        
        if (FULL_COLUMN_MAP.documents) {
            sheet.getRange(TEST_TARGET_ROW, FULL_COLUMN_MAP.documents).setValue(documentsValue);
            results.documents = { column: getColumnLetter(FULL_COLUMN_MAP.documents), value: documentsValue };
        }
        
        if (FULL_COLUMN_MAP.destEmailAddress) {
            sheet.getRange(TEST_TARGET_ROW, FULL_COLUMN_MAP.destEmailAddress).setValue(destEmailValue);
            results.destEmailAddress = { column: getColumnLetter(FULL_COLUMN_MAP.destEmailAddress), value: destEmailValue };
        }
        
        if (FULL_COLUMN_MAP.timestamp) {
            sheet.getRange(TEST_TARGET_ROW, FULL_COLUMN_MAP.timestamp).setValue(timestampValue);
            results.timestamp = { column: getColumnLetter(FULL_COLUMN_MAP.timestamp), value: timestampValue };
        }
        
        console.log('=== Phase3完全テスト成功 ===');
        console.log(`転記項目数: ${Object.keys(results).length}`);
        
        return ContentService
            .createTextOutput(JSON.stringify({
                result: 'success',
                message: 'Phase3完全テスト完了：全項目を転記しました',
                transferredItems: Object.keys(results).length,
                targetRow: TEST_TARGET_ROW,
                columnRange: `F${TEST_TARGET_ROW}:AZ${TEST_TARGET_ROW}`,
                contractors: contractors.length,
                products: products.length,
                timestamp: timestamp,
                details: results
            }))
            .setMimeType(ContentService.MimeType.JSON);
            
    } catch (error) {
        console.error('=== 🚨 Phase3完全テストエラー ===');
        console.error('エラー:', error.message);
        console.error('スタック:', error.stack);
        
        return ContentService
            .createTextOutput(JSON.stringify({
                result: 'error',
                message: `Phase3完全テストエラー: ${error.message}`,
                timestamp: new Date().toISOString()
            }))
            .setMimeType(ContentService.MimeType.JSON);
    }
}

/**
 * 列番号を列文字に変換
 */
function getColumnLetter(columnNumber) {
    let temp, letter = '';
    while (columnNumber > 0) {
        temp = (columnNumber - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        columnNumber = (columnNumber - temp - 1) / 26;
    }
    return letter;
}

/**
 * Phase3完全版手動テスト
 */
function manualTestPhase3Complete() {
    try {
        console.log('=== Phase3完全版手動テスト開始 ===');
        
        const testData = {
            // 発信元
            companyName: 'テスト会社Phase3',
            contactPerson: 'テスト太郎',
            phoneNumber: '03-1234-5678',
            faxNumber: '03-1234-5679',
            emailAddress: 'test@example.com',
            
            // 基本情報
            addressee: 'テスト株式会社',
            honorific: '御中',
            constructionName: 'テスト工事Phase3',
            constructionAddress: '東京都テスト区テスト町1-2-3',
            creationDate: '2025-01-30',
            
            // 業者情報
            contractors: [
                { type: '施工業者', name: 'テスト施工株式会社' },
                { type: '塗装業者', name: 'テスト塗装工業' },
                { type: '納品業者', name: 'テスト納品商事' }
            ],
            
            // 商品情報
            products: [
                { productName: 'テスト商品1', quantity: '10', lotNumber: 'LOT001', shipmentDate: '2025-01-30' },
                { productName: 'テスト商品2', quantity: '5', lotNumber: 'LOT002', shipmentDate: '2025-01-31' },
                { productName: 'テスト商品3', quantity: '20', lotNumber: 'LOT003', shipmentDate: '2025-02-01' }
            ],
            
            // その他
            documents: ['出荷証明書', '成分表・試験成績書'],
            destEmailAddress: 'dest@example.com',
            
            timestamp: new Date().toISOString(),
            testMode: 'phase3-complete'
        };
        
        const mockEvent = {
            postData: {
                contents: JSON.stringify(testData)
            }
        };
        
        const result = doPost(mockEvent);
        console.log('Phase3完全版テスト結果:', JSON.parse(result.getContent()));
        
        console.log('=== Phase3完全版手動テスト完了 ===');
        
    } catch (error) {
        console.error('Phase3完全版手動テストエラー:', error);
    }
}

/**
 * 全列マッピング確認
 */
function checkFullColumnMapping() {
    try {
        console.log('=== 全列マッピング確認 ===');
        
        const spreadsheet = SpreadsheetApp.openById(TEST_SPREADSHEET_ID);
        const sheet = spreadsheet.getSheetByName(TEST_SHEET_NAME);
        
        console.log(`対象行: ${TEST_TARGET_ROW}`);
        console.log('全列マッピング:');
        
        Object.entries(FULL_COLUMN_MAP).forEach(([field, col]) => {
            const colLetter = getColumnLetter(col);
            const currentValue = sheet.getRange(TEST_TARGET_ROW, col).getValue();
            console.log(`${field}: ${colLetter}${TEST_TARGET_ROW} (列${col}) = "${currentValue}"`);
        });
        
        console.log(`総転記項目数: ${Object.keys(FULL_COLUMN_MAP).length}`);
        console.log(`転記範囲: F${TEST_TARGET_ROW}:${getColumnLetter(Math.max(...Object.values(FULL_COLUMN_MAP)))}${TEST_TARGET_ROW}`);
        
        console.log('=== 全列マッピング確認完了 ===');
        
    } catch (error) {
        console.error('全列マッピング確認エラー:', error);
    }
}

/**
 * Phase3完全版データクリア
 */
function clearPhase3CompleteData() {
    try {
        console.log('=== Phase3完全版データクリア開始 ===');
        
        const spreadsheet = SpreadsheetApp.openById(TEST_SPREADSHEET_ID);
        const sheet = spreadsheet.getSheetByName(TEST_SHEET_NAME);
        
        Object.values(FULL_COLUMN_MAP).forEach(col => {
            sheet.getRange(TEST_TARGET_ROW, col).setValue('');
        });
        
        console.log(`${Object.keys(FULL_COLUMN_MAP).length}項目をクリアしました`);
        console.log('=== Phase3完全版データクリア完了 ===');
        
    } catch (error) {
        console.error('Phase3完全版データクリアエラー:', error);
    }
}