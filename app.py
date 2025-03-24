from flask import Flask, render_template, request, send_file, jsonify
import pdfplumber
import pandas as pd
import os
from werkzeug.utils import secure_filename
import tempfile

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max-limit

# アップロードフォルダが存在しない場合は作成
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'pdf'}

# アップロードされたファイルの一時保存ディレクトリ
UPLOAD_FOLDER = 'temp'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def convert_pdf_to_excel(pdf_path):
    all_tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            if tables:
                all_tables.extend(tables)
    
    if not all_tables:
        return None
    
    # 最初のテーブルをDataFrameに変換
    df = pd.DataFrame(all_tables[0])
    return df

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/preview', methods=['POST'])
def preview_pdf():
    if 'file' not in request.files:
        return jsonify({'error': 'ファイルがアップロードされていません'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'ファイルが選択されていません'}), 400
    
    if not file.filename.endswith('.pdf'):
        return jsonify({'error': 'PDFファイルのみ対応しています'}), 400

    try:
        # 一時ファイルとして保存
        temp_path = os.path.join(UPLOAD_FOLDER, secure_filename(file.filename))
        file.save(temp_path)
        
        # PDFからテーブルを抽出
        with pdfplumber.open(temp_path) as pdf:
            # 最初のページからテーブルを抽出
            first_page = pdf.pages[0]
            tables = first_page.extract_tables()
            
            if not tables:
                return jsonify({'error': 'テーブルが見つかりませんでした'}), 400
            
            # 最初のテーブルを使用
            table = tables[0]
            
            # ヘッダーとデータを分離
            headers = table[0]
            data = table[1:6]  # 最初の5行のみを表示
            
            # 一時ファイルを削除
            os.remove(temp_path)
            
            return jsonify({
                'headers': headers,
                'data': data
            })
            
    except Exception as e:
        if os.path.exists(temp_path):
            os.remove(temp_path)
        return jsonify({'error': str(e)}), 500

@app.route('/convert', methods=['POST'])
def convert_pdf():
    if 'file' not in request.files:
        return 'ファイルがアップロードされていません', 400
    
    file = request.files['file']
    if file.filename == '':
        return 'ファイルが選択されていません', 400
    
    if not file.filename.endswith('.pdf'):
        return 'PDFファイルのみ対応しています', 400

    try:
        # 一時ファイルとして保存
        temp_path = os.path.join(UPLOAD_FOLDER, secure_filename(file.filename))
        file.save(temp_path)
        
        # PDFからテーブルを抽出
        with pdfplumber.open(temp_path) as pdf:
            # 最初のページからテーブルを抽出
            first_page = pdf.pages[0]
            tables = first_page.extract_tables()
            
            if not tables:
                return 'テーブルが見つかりませんでした', 400
            
            # 最初のテーブルを使用
            table = tables[0]
            
            # DataFrameに変換
            df = pd.DataFrame(table[1:], columns=table[0])
            
            # Excelファイルとして保存
            excel_path = os.path.join(UPLOAD_FOLDER, os.path.splitext(file.filename)[0] + '.xlsx')
            df.to_excel(excel_path, index=False)
            
            # 一時ファイルを削除
            os.remove(temp_path)
            
            # Excelファイルを送信
            return send_file(
                excel_path,
                as_attachment=True,
                download_name=os.path.splitext(file.filename)[0] + '.xlsx'
            )
            
    except Exception as e:
        if os.path.exists(temp_path):
            os.remove(temp_path)
        return str(e), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5001, debug=True) 