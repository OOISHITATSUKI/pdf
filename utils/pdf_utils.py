import os
import subprocess
import logging
import streamlit as st

def check_poppler_installation():
    try:
        # PATHの確認
        path = os.environ.get('PATH', '')
        st.write(f"現在のPATH: {path}")
        
        # popplerのバージョン確認
        result = subprocess.run(['pdftoppm', '-v'], capture_output=True, text=True)
        st.write(f"Popplerバージョン情報: {result.stderr}")
        return True
    except Exception as e:
        st.error(f"Popplerの確認中にエラーが発生しました: {str(e)}")
        return False

def generate_pdf_preview(pdf_path, output_path):
    if not check_poppler_installation():
        st.error("Popplerが正しくインストールされていないか、PATHが通っていません。")
        return False
    
    try:
        # PDFの最初のページをJPEGに変換
        cmd = ['pdftoppm', '-jpeg', '-f', '1', '-l', '1', pdf_path, output_path]
        st.write(f"実行するコマンド: {' '.join(cmd)}")
        
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode != 0:
            st.error(f"プレビュー生成エラー: {result.stderr}")
            return False
        return True
    except Exception as e:
        st.error(f"プレビュー生成中にエラーが発生しました: {str(e)}")
        return False 