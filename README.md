# PDF to Excel 変換ツール

PDFファイルをExcel形式に変換できるWebアプリケーションです。すべての処理はブラウザ内で行われます。

## 機能

- PDFファイルのアップロード（ドラッグ＆ドロップ対応）
- テーブルデータのプレビュー表示
- Excel形式への変換
- 複数ページPDFの対応

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