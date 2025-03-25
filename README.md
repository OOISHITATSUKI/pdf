# PDF to Excel 変換ツール

## 概要
PDFファイルをExcelに変換するWebアプリケーション。
請求書、確定申告書、決算書などの帳票をレイアウトそのままでExcelに変換できます。

## 機能
- PDFからExcelへの変換
- OCRによる画像型PDFの処理
- 確定申告書・決算書の専用処理
- Google Drive連携（有料プラン）
- 変換履歴の保存（有料プラン）

## 必要なAPI設定

### 1. Google Cloud Vision API
1. Google Cloud Platformでプロジェクトを作成
2. Vision APIを有効化
3. サービスアカウントキーを作成
4. キーファイルを安全な場所に保存
5. 環境変数の設定:
   ```bash
   export GOOGLE_APPLICATION_CREDENTIALS="/path/to/service-account-key.json"
   ```

### 2. Google Drive API
1. Google Cloud ConsoleでDrive APIを有効化
2. OAuth 2.0クライアントIDを作成
3. 認証情報をダウンロード
4. 環境変数の設定:
   ```bash
   export GOOGLE_DRIVE_CLIENT_ID="your-client-id"
   export GOOGLE_DRIVE_CLIENT_SECRET="your-client-secret"
   ```

### 3. Stripe API（決済用）
1. Stripeアカウントを作成
2. APIキーを取得
3. 環境変数の設定:
   ```bash
   export STRIPE_SECRET_KEY="your-stripe-secret-key"
   export STRIPE_PUBLISHABLE_KEY="your-stripe-publishable-key"
   ```

### 4. PostgreSQL設定
1. データベースを作成
2. 環境変数の設定:
   ```bash
   export DATABASE_URL="postgresql://user:password@localhost:5432/dbname"
   ```

### 5. Redis設定（セッション管理用）
1. Redisサーバーをセットアップ
2. 環境変数の設定:
   ```bash
   export REDIS_URL="redis://localhost:6379/0"
   ```

## インストール手順

1. 依存関係のインストール:
   ```bash
   pip install -r requirements.txt
   ```

2. Tesseractのインストール:
   - macOS:
     ```bash
     brew install tesseract
     brew install tesseract-lang
     ```
   - Ubuntu:
     ```bash
     sudo apt-get install tesseract-ocr
     sudo apt-get install tesseract-ocr-jpn
     ```

3. 環境変数の設定:
   ```bash
   source .env  # 環境変数ファイルを作成して使用
   ```

4. アプリケーションの起動:
   ```bash
   streamlit run app.py
   ```

## 開発環境
- Python 3.9+
- PostgreSQL 13+
- Redis 6+

## ライセンス
MIT License

## 対応PDFの種類

- テキストが含まれるPDF
- 表形式のデータが含まれるPDF
- ファイルサイズ16MB以下

## 技術スタック

- Frontend: HTML, CSS, JavaScript
- Backend: Python, Flask
- PDF処理: pdfplumber
- Excel処理: pandas, openpyxl

## ローカルでの実行方法

1. リポジトリをクローン
```bash
git clone [リポジトリURL]
cd pdf-to-excel-converter
```

2. 仮想環境を作成し、依存関係をインストール
```bash
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

3. アプリケーションを実行
```bash
python app.py
```

4. ブラウザで http://localhost:5001 にアクセス

## デプロイ

このアプリケーションはNetlifyでホストされています。
デプロイは自動的に行われ、mainブランチへのプッシュ時に更新されます。 