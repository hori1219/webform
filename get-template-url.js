// テンプレートURLを取得する関数
function getTemplateUrl() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const templateDocId = properties.getProperty('TEMPLATE_DOC_ID');
    
    if (templateDocId) {
      const templateUrl = `https://docs.google.com/document/d/${templateDocId}/edit`;
      console.log('テンプレートURL:', templateUrl);
      console.log('テンプレートID:', templateDocId);
      return templateUrl;
    } else {
      console.log('テンプレートがまだ作成されていません。testFormSubmission()を実行してください。');
      return null;
    }
  } catch (error) {
    console.error('エラー:', error);
    return null;
  }
}

// 新しいテンプレートを強制作成する関数
function createNewTemplate() {
  try {
    // 既存のテンプレートIDを削除
    const properties = PropertiesService.getScriptProperties();
    properties.deleteProperty('TEMPLATE_DOC_ID');
    
    // 新しいテンプレートを作成
    const templateDocId = createDocumentTemplate();
    properties.setProperty('TEMPLATE_DOC_ID', templateDocId);
    
    const templateUrl = `https://docs.google.com/document/d/${templateDocId}/edit`;
    console.log('新しいテンプレートが作成されました:');
    console.log('URL:', templateUrl);
    console.log('ID:', templateDocId);
    
    return templateUrl;
  } catch (error) {
    console.error('テンプレート作成エラー:', error);
    return null;
  }
}

// 全テンプレート情報を表示
function showAllTemplateInfo() {
  console.log('=== テンプレート情報 ===');
  
  const properties = PropertiesService.getScriptProperties();
  const templateDocId = properties.getProperty('TEMPLATE_DOC_ID');
  
  if (templateDocId) {
    console.log('テンプレートID:', templateDocId);
    console.log('テンプレートURL:', `https://docs.google.com/document/d/${templateDocId}/edit`);
    
    try {
      const file = DriveApp.getFileById(templateDocId);
      console.log('テンプレート名:', file.getName());
      console.log('作成日時:', file.getDateCreated());
      console.log('最終更新:', file.getLastUpdated());
      console.log('削除状態:', file.isTrashed() ? '削除済み' : '有効');
    } catch (e) {
      console.log('テンプレートファイルにアクセスできません:', e.message);
    }
  } else {
    console.log('テンプレートが見つかりません');
  }
  
  console.log('==================');
}