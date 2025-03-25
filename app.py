import streamlit as st
import pdfplumber
import pandas as pd
import numpy as np
from PIL import Image
import tempfile
import os
import re
from datetime import datetime
from openpyxl.utils import get_column_letter
import uuid
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import Column, String, DateTime, Enum, JSON, ForeignKey, Text
from sqlalchemy.sql import func
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
import io
import json
import sqlite3
from dataclasses import dataclass
from typing import Optional

# ページ設定（必ず最初に実行）
st.set_page_config(
    page_title="PDF to Excel 変換ツール",
    page_icon="📄",
    layout="wide"
)

# データベースの設定
DB_PATH = "pdf_converter.db"

@dataclass
class ConversionHistory:
    id: int
    user_id: str
    document_type: str
    document_date: str
    conversion_date: datetime
    file_name: str
    status: str

def init_db():
    """データベースの初期化"""
    try:
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        
        # ユーザーテーブル
        c.execute('''
            CREATE TABLE IF NOT EXISTS users (
                id TEXT PRIMARY KEY,
                email TEXT UNIQUE,
                plan_type TEXT DEFAULT 'free',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # 変換履歴テーブル
        c.execute('''
            CREATE TABLE IF NOT EXISTS conversion_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id TEXT,
                document_type TEXT,
                document_date TEXT,
                conversion_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                file_name TEXT,
                status TEXT,
                FOREIGN KEY (user_id) REFERENCES users (id)
            )
        ''')
        
        # 変換カウントテーブル
        c.execute('''
            CREATE TABLE IF NOT EXISTS conversion_count (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id TEXT,
                count_date TEXT,
                count INTEGER DEFAULT 0,
                FOREIGN KEY (user_id) REFERENCES users (id),
                UNIQUE(user_id, count_date)
            )
        ''')
        
        # 初期ユーザーの作成（セッションユーザー用）
        c.execute('''
            INSERT OR IGNORE INTO users (id, plan_type)
            VALUES (?, 'free')
        ''', (str(datetime.now().timestamp()),))
        
        conn.commit()
    except Exception as e:
        st.error(f"データベースの初期化中にエラーが発生しました: {str(e)}")
    finally:
        conn.close()

def save_conversion_history(user_id: str, document_type: str, document_date: str, file_name: str, status: str):
    """変換履歴を保存"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''
        INSERT INTO conversion_history 
        (user_id, document_type, document_date, file_name, status)
        VALUES (?, ?, ?, ?, ?)
    ''', (user_id, document_type, document_date, file_name, status))
    conn.commit()
    conn.close()

def get_daily_conversion_count(user_id: str) -> int:
    """ユーザーの本日の変換回数を取得"""
    try:
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        today = datetime.now().strftime('%Y-%m-%d')
        
        # ユーザーと日付の組み合わせがなければ作成
        c.execute('''
            INSERT OR IGNORE INTO conversion_count (user_id, count_date, count)
            VALUES (?, ?, 0)
        ''', (user_id, today))
        
        # カウントを取得
        c.execute('''
            SELECT count FROM conversion_count
            WHERE user_id = ? AND count_date = ?
        ''', (user_id, today))
        
        result = c.fetchone()
        conn.commit()
        return result[0] if result else 0
    except Exception as e:
        st.error(f"変換回数の取得中にエラーが発生しました: {str(e)}")
        return 0
    finally:
        conn.close()

def increment_conversion_count(user_id: str) -> bool:
    """変換回数をインクリメント"""
    try:
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        today = datetime.now().strftime('%Y-%m-%d')
        
        c.execute('''
            UPDATE conversion_count
            SET count = count + 1
            WHERE user_id = ? AND count_date = ?
        ''', (user_id, today))
        
        if c.rowcount == 0:
            c.execute('''
                INSERT INTO conversion_count (user_id, count_date, count)
                VALUES (?, ?, 1)
            ''', (user_id, today))
        
        conn.commit()
        return True
    except Exception as e:
        st.error(f"変換回数の更新中にエラーが発生しました: {str(e)}")
        return False
    finally:
        conn.close()

def check_conversion_limit(user_id: str) -> bool:
    """変換制限をチェック"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    # ユーザーの種類を確認
    c.execute('SELECT plan_type FROM users WHERE id = ?', (user_id,))
    result = c.fetchone()
    plan_type = result[0] if result else 'free'
    
    # 本日の変換回数を取得
    daily_count = get_daily_conversion_count(user_id)
    
    # プランごとの制限チェック
    if plan_type == 'premium':
        return True  # 無制限
    elif plan_type == 'basic':
        return daily_count < 5  # 1日5回まで
    else:  # free
        return daily_count < 3  # 1日3回まで

# セッション状態の初期化
if 'user_id' not in st.session_state:
    st.session_state.user_id = str(datetime.now().timestamp())

# データベースの初期化を実行
init_db()

def create_hero_section():
    """ヒーローセクションを作成"""
    st.title("PDF to Excel 変換ツール")
    st.write("PDFファイルをかんたんにExcelに変換できます。")
    st.write("請求書、決算書、納品書など、帳票をレイアウトそのままで変換可能。")
    st.write("ブラウザ上で完結し、安心・安全にご利用いただけます。")

def create_login_section():
    """ログインセクションを作成"""
    with st.sidebar:
        st.subheader("ログイン")
        email = st.text_input("メールアドレス")
        password = st.text_input("パスワード", type="password")
        if st.button("ログイン"):
            # ログイン処理（実装予定）
            pass
        
        st.markdown("---")
        st.subheader("新規登録")
        if st.button("アカウントを作成"):
            # 新規登録処理（実装予定）
            pass

def create_preview_section(uploaded_file):
    """プレビューセクションを作成"""
    st.subheader("プレビュー")
    if uploaded_file is not None:
        preview_image = create_preview(uploaded_file)
        if preview_image is not None:
            st.image(preview_image, use_container_width=True)

def create_upload_section():
    """アップロードセクションを作成"""
    st.subheader("ファイルをアップロード")
    
    # 残り変換回数の表示
    daily_count = get_daily_conversion_count(st.session_state.user_id)
    remaining = 3 - daily_count  # 基本は3回
    st.markdown(f"📊 本日の残り変換回数：{remaining}/3回")
    
    # ドキュメントタイプの選択
    document_type = st.selectbox(
        "ドキュメントの種類を選択",
        ["請求書", "見積書", "納品書", "確定申告書", "その他"]
    )
    
    # 日付入力
    document_date = st.date_input(
        "書類の日付",
        value=None,
        help="YYYY/MM/DD形式で入力してください"
    )
    
    # ファイルアップロード
    uploaded_file = st.file_uploader(
        "クリックまたはドラッグ&ドロップでPDFファイルを選択",
        type=['pdf'],
        help="ファイルサイズの制限: 200MB"
    )
    
    st.info("💡 無料プランでは1ページ目のみ変換されます。全ページ変換は有料プランでご利用いただけます。")
    
    if uploaded_file is not None:
        if st.button("Excelに変換する"):
            if check_conversion_limit(st.session_state.user_id):
                try:
                    excel_data = process_pdf(uploaded_file, document_type, document_date)
                    # 変換履歴を保存
                    save_conversion_history(
                        st.session_state.user_id,
                        document_type,
                        document_date.strftime('%Y-%m-%d') if document_date else None,
                        uploaded_file.name,
                        "success"
                    )
                    st.download_button(
                        label="Excelファイルをダウンロード",
                        data=excel_data,
                        file_name="converted.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"処理中にエラーが発生しました: {str(e)}")
                    # エラー履歴を保存
                    save_conversion_history(
                        st.session_state.user_id,
                        document_type,
                        document_date.strftime('%Y-%m-%d') if document_date else None,
                        uploaded_file.name,
                        f"error: {str(e)}"
                    )
            else:
                st.error("本日の変換回数制限に達しました。プランをアップグレードすると、より多くの変換が可能です。")
    
    return uploaded_file

def process_pdf(uploaded_file, document_type=None, document_date=None):
    """PDFの処理を行う関数"""
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_pdf:
            temp_pdf.write(uploaded_file.getvalue())
            pdf_path = temp_pdf.name

        with pdfplumber.open(pdf_path) as pdf:
            # 1ページ目のみ処理（無料プラン）
            page = pdf.pages[0]
            
            # テーブルの抽出
            tables = page.extract_tables()
            if not tables:
                raise ValueError("テーブルが見つかりませんでした")

            # Excelファイルの作成
            wb = Workbook()
            ws = wb.active
            
            # スタイルの定義
            header_font = Font(bold=True)
            border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            
            # テーブルデータの書き込み
            for i, row in enumerate(tables[0], 1):
                for j, cell in enumerate(row, 1):
                    if cell is not None:
                        cell_value = str(cell).strip()
                        ws_cell = ws.cell(row=i, column=j, value=cell_value)
                        ws_cell.border = border
                        if i == 1:
                            ws_cell.font = header_font
                        if cell_value.replace(',', '').replace('.', '').isdigit():
                            ws_cell.alignment = Alignment(horizontal='right')

            # 列幅の自動調整
            for column in ws.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 2)
                ws.column_dimensions[column_letter].width = adjusted_width

            # 一時ファイルとして保存
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_excel:
                wb.save(temp_excel.name)
                with open(temp_excel.name, 'rb') as f:
                    excel_data = f.read()

            # 一時ファイルの削除
            os.unlink(pdf_path)
            os.unlink(temp_excel.name)

            return excel_data

    except Exception as e:
        raise Exception(f"PDFの処理中にエラーが発生しました: {str(e)}")

def create_preview(uploaded_file):
    """PDFのプレビューを生成する関数"""
    try:
        if uploaded_file is not None:
            with pdfplumber.open(io.BytesIO(uploaded_file.getvalue())) as pdf:
                first_page = pdf.pages[0]
                img = first_page.to_image()
                img_byte_arr = io.BytesIO()
                img.save(img_byte_arr, format='PNG')
                img_byte_arr = img_byte_arr.getvalue()
                return img_byte_arr
        return None
    except Exception as e:
        st.error(f"プレビューの生成中にエラーが発生しました: {str(e)}")
        return None

def create_document_type_buttons():
    """ドキュメントタイプ選択ボタンを作成"""
    st.write("ドキュメントの種類を選択")
    
    # セッション状態の初期化
    if 'selected_document_type' not in st.session_state:
        st.session_state.selected_document_type = None
    
    # ラジオボタンをボタン風にするスタイル
    button_style = """
        <style>
        /* ラジオボタンを非表示 */
        div[data-testid="stRadio"] > div > div > label > div:first-child {
            display: none;
        }
        
        /* ラベルをボタン風にスタイリング */
        div[data-testid="stRadio"] label {
            width: 100%;
            min-height: 60px;
            margin: 8px 0;
            padding: 10px 15px;
            border: 2px solid #e0e0e0;
            border-radius: 8px;
            background: linear-gradient(145deg, #ffffff 0%, #f5f5f5 100%);
            font-size: 16px;
            color: #333;
            transition: all 0.3s ease;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            display: flex;
            align-items: center;
            justify-content: center;
            cursor: pointer;
        }
        
        /* ホバー時のスタイル */
        div[data-testid="stRadio"] label:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0,0,0,0.15);
            border-color: #2196F3;
            background: linear-gradient(145deg, #f5f5f5 0%, #e3f2fd 100%);
        }
        
        /* 選択時のスタイル */
        div[data-testid="stRadio"] label[data-checked="true"] {
            border-color: #2196F3 !important;
            background: linear-gradient(145deg, #e3f2fd 0%, #bbdefb 100%) !important;
            color: #1565C0 !important;
            box-shadow: 0 4px 8px rgba(33,150,243,0.2) !important;
        }
        
        /* 2カラムレイアウト */
        div[data-testid="stRadio"] > div {
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 10px;
        }
        </style>
    """
    st.markdown(button_style, unsafe_allow_html=True)
    
    # ドキュメントタイプの定義
    document_types = {
        "見積書": "estimate",
        "請求書": "invoice",
        "納品書": "delivery",
        "領収書": "receipt",
        "決算書": "financial",
        "確定申告書": "tax_return",
        "その他": "other"
    }
    
    # ラジオボタンで選択
    selected_label = st.radio(
        "書類の種類",
        options=list(document_types.keys()),
        key="doc_type_radio",
        label_visibility="collapsed",  # ラベルを非表示
        horizontal=True,  # 水平配置
        index=None if st.session_state.selected_document_type is None else 
              list(document_types.values()).index(st.session_state.selected_document_type)
    )
    
    # 選択状態の更新
    if selected_label is not None:
        st.session_state.selected_document_type = document_types[selected_label]
    else:
        st.warning("書類の種類を選択してください")
        return None
    
    return st.session_state.selected_document_type

def create_footer():
    """フッターを作成"""
    st.markdown("---")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("[利用規約](/terms)")
    with col2:
        st.markdown("[プライバシーポリシー](/privacy)")
    with col3:
        st.markdown("[お問い合わせ](/contact)")

def main():
    """メイン関数"""
    create_hero_section()
    create_login_section()
    
    # 2カラムレイアウト
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.subheader("ファイルをアップロード")
        
        # 残り変換回数の表示
        daily_count = get_daily_conversion_count(st.session_state.user_id)
        remaining = 3 - daily_count  # 基本は3回
        st.markdown(f"📊 本日の残り変換回数：{remaining}/3回")
        
        # ドキュメントタイプの選択（ボタン形式）
        document_type = create_document_type_buttons()
        
        # 日付入力
        document_date = st.date_input(
            "書類の日付",
            value=None,
            help="YYYY/MM/DD形式で入力してください"
        )
        
        # ファイルアップロード
        uploaded_file = st.file_uploader(
            "クリックまたはドラッグ&ドロップでPDFファイルを選択",
            type=['pdf'],
            help="ファイルサイズの制限: 200MB"
        )
        
        st.info("💡 無料プランでは1ページ目のみ変換されます。全ページ変換は有料プランでご利用いただけます。")
        
        if uploaded_file is not None and document_type is not None:
            if st.button("Excelに変換する"):
                if check_conversion_limit(st.session_state.user_id):
                    try:
                        excel_data = process_pdf(uploaded_file, document_type, document_date)
                        # 変換履歴を保存
                        save_conversion_history(
                            st.session_state.user_id,
                            document_type,
                            document_date.strftime('%Y-%m-%d') if document_date else None,
                            uploaded_file.name,
                            "success"
                        )
                        st.download_button(
                            label="Excelファイルをダウンロード",
                            data=excel_data,
                            file_name="converted.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    except Exception as e:
                        st.error(f"処理中にエラーが発生しました: {str(e)}")
                        # エラー履歴を保存
                        save_conversion_history(
                            st.session_state.user_id,
                            document_type,
                            document_date.strftime('%Y-%m-%d') if document_date else None,
                            uploaded_file.name,
                            f"error: {str(e)}"
                        )
                else:
                    st.error("本日の変換回数制限に達しました。プランをアップグレードすると、より多くの変換が可能です。")
    
    with col2:
        st.subheader("プレビュー")
        if uploaded_file is not None:
            preview_image = create_preview(uploaded_file)
            if preview_image is not None:
                st.image(preview_image, use_container_width=True)
    
    create_footer()

if __name__ == "__main__":
    main() 