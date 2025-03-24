from flask import Flask, request, jsonify, send_file
import pdfplumber
import pandas as pd
import os
import tempfile
from werkzeug.utils import secure_filename
import logging

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB制限
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()  # システムの一時ディレクトリを使用

# ロギングの設定
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def extract_table_from_pdf(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # 最初のページから表を抽出
            first_page = pdf.pages[0]
            table = first_page.extract_table()
            
            if table is None:
                return None
            
            # ヘッダー行を取得
            headers = table[0]
            
            # データ行を取得
            data = table[1:]
            
            # DataFrameを作成
            df = pd.DataFrame(data, columns=headers)
            return df
    except Exception as e:
        logger.error(f"PDFからの表抽出中にエラーが発生しました: {str(e)}")
        return None

@app.route('/')
def index():
    return app.send_static_file('index.html')

@app.route('/preview', methods=['POST'])
def preview():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'ファイルがアップロードされていません'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'ファイルが選択されていません'}), 400
        
        if not file.filename.endswith('.pdf'):
            return jsonify({'error': 'PDFファイルのみ対応しています'}), 400
        
        # ファイルを一時的に保存
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # 表を抽出
        df = extract_table_from_pdf(filepath)
        
        # 一時ファイルを削除
        os.remove(filepath)
        
        if df is None:
            return jsonify({'error': '表が見つかりませんでした'}), 400
        
        # データをJSONに変換
        data = {
            'headers': df.columns.tolist(),
            'rows': df.values.tolist()
        }
        
        return jsonify(data)
    except Exception as e:
        logger.error(f"プレビュー生成中にエラーが発生しました: {str(e)}")
        return jsonify({'error': 'プレビューの生成に失敗しました'}), 500

@app.route('/convert', methods=['POST'])
def convert():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'ファイルがアップロードされていません'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'ファイルが選択されていません'}), 400
        
        if not file.filename.endswith('.pdf'):
            return jsonify({'error': 'PDFファイルのみ対応しています'}), 400
        
        # ファイルを一時的に保存
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # 表を抽出
        df = extract_table_from_pdf(filepath)
        
        # 一時ファイルを削除
        os.remove(filepath)
        
        if df is None:
            return jsonify({'error': '表が見つかりませんでした'}), 400
        
        # Excelファイルを生成
        excel_filename = os.path.splitext(filename)[0] + '.xlsx'
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)
        df.to_excel(excel_path, index=False)
        
        # ファイルを送信
        response = send_file(
            excel_path,
            as_attachment=True,
            download_name=excel_filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
        # 一時ファイルを削除
        os.remove(excel_path)
        
        return response
    except Exception as e:
        logger.error(f"変換中にエラーが発生しました: {str(e)}")
        return jsonify({'error': '変換に失敗しました'}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5001) 