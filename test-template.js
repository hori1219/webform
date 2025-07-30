// テンプレートテスト用関数
function testTemplateGeneration() {
  console.log('=== テンプレートテスト開始 ===');
  
  try {
    // テストデータ
    const testData = {
      timestamp: new Date().toLocaleString('ja-JP'),
      companyName: 'テスト株式会社',
      contactPerson: '山田太郎',
      phoneNumber: '03-1234-5678',
      emailAddress: 'test@example.com',
      addressee: 'サンプル建設',
      honorific: '御中',
      constructionName: 'テスト工事プロジェクト',
      constructionAddress: '東京都千代田区丸の内1-1-1',
      creationDate: '2025-01-30',
      contractors: [
        { type: '元請', name: 'ABC建設株式会社' },
        { type: '下請', name: 'XYZ工業株式会社' },
        { type: '資材', name: 'DEF商事' }
      ],
      products: [
        { 
          productName: 'ローバル１ｋｇ缶',
          quantity: '10缶',
          lotNumber: 'LOT2025-001',
          shipmentDate: '2025-02-01'
        },
        { 
          productName: 'シーラー500ml',
          quantity: '5本',
          lotNumber: 'LOT2025-002',
          shipmentDate: '2025-02-01'
        }
      ],
      documents: ['成分表・試験成績書', 'ＳＤＳ', 'カタログ'],
      destCompanyName: '受信先建設株式会社',
      destContactPerson: '鈴木花子',
      destPhoneNumber: '06-5678-9012',
      destEmailAddress: 'test-dest@example.com',
      remarks: 'テスト用の依頼書です。レイアウト確認のため。'
    };
    
    // 管理番号を生成（簡易版）
    const managementNumber = 'TEST-001-01';
    
    // PDF生成テスト
    console.log('PDF生成テスト開始...');
    const pdfResult = generateRequestPDF(testData, managementNumber);
    
    console.log('PDF生成成功！');
    console.log('PDF URL:', pdfResult.pdfUrl);
    console.log('PDF ID:', pdfResult.pdfId);
    
    // テンプレートURL表示
    const properties = PropertiesService.getScriptProperties();
    const templateDocId = properties.getProperty('TEMPLATE_DOC_ID');
    if (templateDocId) {
      console.log('使用テンプレートURL:', `https://docs.google.com/document/d/${templateDocId}/edit`);
    }
    
    console.log('=== テスト完了 ===');
    return {
      status: 'success',
      pdfUrl: pdfResult.pdfUrl,
      templateUrl: `https://docs.google.com/document/d/${templateDocId}/edit`
    };
    
  } catch (error) {
    console.error('テストエラー:', error);
    console.log('エラー詳細:', error.stack);
    return {
      status: 'error',
      message: error.toString()
    };
  }
}

// 簡易テスト（PDFファイル名のみ生成）
function quickTest() {
  console.log('=== 簡易テスト ===');
  
  try {
    // テンプレートIDを確認
    const properties = PropertiesService.getScriptProperties();
    const templateDocId = properties.getProperty('TEMPLATE_DOC_ID');
    
    if (templateDocId) {
      console.log('テンプレートID:', templateDocId);
      console.log('テンプレートURL:', `https://docs.google.com/document/d/${templateDocId}/edit`);
      
      // テンプレートファイルの存在確認
      const file = DriveApp.getFileById(templateDocId);
      console.log('テンプレート名:', file.getName());
      console.log('テンプレート状態:', file.isTrashed() ? '削除済み' : '有効');
      
      return 'テンプレート確認完了';
    } else {
      console.log('テンプレートが見つかりません');
      return 'テンプレート未作成';
    }
  } catch (error) {
    console.error('簡易テストエラー:', error);
    return 'エラー発生';
  }
}