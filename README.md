# 🚚 出荷証明書Webシステム

出荷証明書をWebフォームで入力し、Google Sheetsに自動転記するシステムです。

*created by Claude Sonnet 4*

## 🌟 システム概要

このシステムは出荷証明書の管理を効率化するWebアプリケーションです：

- 📝 **直感的なWebフォーム** - 発送者・発送先・商品情報を簡単入力
- 📊 **Google Sheets自動転記** - 入力データを自動でスプレッドシートに保存
- 💾 **CSV出力機能** - ローカルでのデータ保存も可能
- 📱 **レスポンシブデザイン** - PC・タブレット・スマートフォン対応
- ✅ **リアルタイム検証** - 入力ミスを即座にチェック

## 📂 プロジェクト構成

```
webform/
├── src/
│   └── index.html          # メインのWebフォーム
├── scripts/
│   └── gas-shipping-form.js # Google Apps Script コード
├── docs/
└── README.md               # このファイル
```

## 🚀 セットアップ手順

### 1. Webフォームの準備

1. `src/index.html` をWebサーバーにアップロード、またはローカルで開く
2. ファイルをブラウザで開いて動作確認

### 2. Google Sheets + Apps Script の設定

#### Step 1: Google Sheetsの準備
1. [Google Sheets](https://sheets.google.com/) で新しいスプレッドシートを作成
2. スプレッドシートのURLから **ID** をコピー
   ```
   https://docs.google.com/spreadsheets/d/[SPREADSHEET_ID]/edit
   ```

#### Step 2: Google Apps Script の設定
1. 作成したスプレッドシートで「**拡張機能**」→「**Apps Script**」を選択
2. `scripts/gas-shipping-form.js` の内容をコピー＆ペースト
3. **設定項目を編集**：
   ```javascript
   const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE'; // ← Step 1でコピーしたID
   const SHEET_NAME = '出荷証明書'; // ← 使用するシート名
   const NOTIFICATION_EMAIL = 'your-email@example.com'; // ← 通知メール（オプション）
   ```

#### Step 3: ウェブアプリとして公開
1. Apps Script エディターで「**デプロイ**」→「**新しいデプロイ**」
2. 種類を「**ウェブアプリ**」に設定
3. 実行者を「**自分**」、アクセス権限を「**全員**」に設定
4. 「**デプロイ**」をクリックして **ウェブアプリURL** を取得

#### Step 4: フォームとの連携
1. `src/index.html` を開く
2. 以下の行を修正：
   ```javascript
   const GOOGLE_SCRIPT_URL = 'YOUR_GOOGLE_APPS_SCRIPT_URL_HERE';
   ```
   ↓
   ```javascript
   const GOOGLE_SCRIPT_URL = 'https://script.google.com/macros/s/.../exec';
   ```

### 3. テスト実行
1. Webフォームでテストデータを入力・送信
2. Google Sheetsにデータが追加されることを確認

## 📋 フォーム項目

### 発送者情報
- 会社名 *（必須）*
- 担当者名 *（必須）*
- メールアドレス *（必須）*
- 電話番号

### 発送先情報
- 会社名 *（必須）*
- 担当者名 *（必須）*
- 住所 *（必須）*
- メールアドレス
- 電話番号

### 発送情報
- 発送日 *（必須）*
- 発送方法 *（必須）*
- 追跡番号
- 梱包数 *（必須）*

### 商品情報
- 商品名 *（必須）*
- 商品コード
- 数量 *（必須）*
- 単位
- 重量 (kg)
- 商品価格 (円)
- 備考・特記事項

## ⚙️ カスタマイズ

### フォーム項目の追加・変更
1. `src/index.html` のフォーム部分を編集
2. `scripts/gas-shipping-form.js` のデータ処理部分も対応修正

### デザインの変更
- CSS部分を編集してカラーテーマやレイアウトを変更可能

### 通知機能
- Google Apps Script で自動メール通知を設定可能
- Slack や Chatwork 等への通知も実装可能

## 🔧 トラブルシューティング

### データが送信されない場合
1. **Google Script URL** が正しく設定されているか確認
2. Apps Script の **実行権限** が適切に設定されているか確認
3. ブラウザの開発者ツールでエラーメッセージを確認

### スプレッドシートにデータが追加されない場合
1. **SPREADSHEET_ID** が正しいか確認
2. **シート名** が存在するか確認
3. Apps Script の実行ログを確認

## 📊 データ分析機能

Google Apps Script には以下の分析機能も含まれています：
- 発送方法別の集計
- 期間別の出荷件数分析
- `analyzeData()` 関数をApps Scriptエディターで実行

## 💡 活用例

- **物流会社**: 日々の出荷業務の効率化
- **製造業**: 製品出荷の記録管理
- **EC事業者**: 注文商品の発送管理
- **卸売業**: 取引先への商品配送管理

---

**出荷業務の効率化を実現しましょう！** 🚚✨