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
            st.image(preview_image, use_column_width=True)

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

def main():
    """メイン関数"""
    create_hero_section()
    create_login_section()
    
    # 2カラムレイアウト
    col1, col2 = st.columns([1, 1])
    
    with col1:
        uploaded_file = create_upload_section()
    
    with col2:
        create_preview_section(uploaded_file)

if __name__ == "__main__":
    main() 