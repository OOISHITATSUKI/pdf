import streamlit as st
import pdfplumber
from pdf2image import convert_from_bytes
import pandas as pd
import numpy as np
from PIL import Image
import tempfile
import os
import re
from datetime import datetime, date, timedelta
from openpyxl.utils import get_column_letter
import uuid
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import Column, String, DateTime, Enum, JSON, ForeignKey, Text
from sqlalchemy.sql import func
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
import io
import json
import sqlite3
from dataclasses import dataclass
from typing import Optional, Dict, Set
from openpyxl.cell.cell import MergedCell
import hashlib
import base64
import requests
from dotenv import load_dotenv

# 環境変数の読み込み
load_dotenv()

# Popplerのパス設定
if os.path.exists('/usr/local/bin/pdftoppm'):
    os.environ['PATH'] = f"/usr/local/bin:{os.environ['PATH']}"

# ページ設定（必ず最初に実行）
st.set_page_config(
    page_title="PDF to Excel 変換ツール",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# セッション状態の初期化
if 'rerun_count' not in st.session_state:
    st.session_state.rerun_count = 0
if 'last_rerun_time' not in st.session_state:
    st.session_state.last_rerun_time = datetime.now()
if 'conversion_success' not in st.session_state:
    st.session_state.conversion_success = False
if 'processing_pdf' not in st.session_state:
    st.session_state.processing_pdf = False

# データベースの設定
DB_PATH = "pdf_converter.db"

# プラン定義
PLAN_LIMITS = {
    "free_guest": 3,        # 未ログインユーザー
    "free_registered": 5,   # 登録済み無料ユーザー
    "premium_basic": 1000,  # $5プラン
    "premium_pro": float('inf')  # $20プラン
}

PLAN_NAMES = {
    "free_guest": "無料プラン（未登録）",
    "free_registered": "無料プラン（登録済）",
    "premium_basic": "ベーシックプラン（$5）",
    "premium_pro": "プロフェッショナルプラン（$20）"
}

class ConversionTracker:
    def __init__(self):
        self.conn = sqlite3.connect('conversion_tracker.db')
        self.setup_database()
    
    def setup_database(self):
        """データベースのセットアップ"""
        with self.conn:
            self.conn.execute("""
                CREATE TABLE IF NOT EXISTS conversion_counts (
                    user_id TEXT,
                    ip_address TEXT,
                    browser_id TEXT,
                    conversion_date DATE,
                    count INTEGER,
                    PRIMARY KEY (user_id, ip_address, browser_id, conversion_date)
                )
            """)
            
            self.conn.execute("""
                CREATE TABLE IF NOT EXISTS user_plans (
                    user_id TEXT PRIMARY KEY,
                    plan_type TEXT,
                    updated_at TIMESTAMP
                )
            """)
    
    def get_unique_identifier(self, user_id: Optional[str] = None) -> str:
        """ユーザー識別子の生成（IPアドレス + ブラウザID + ユーザーID）"""
        ip = st.session_state.get('client_ip', 'unknown')
        browser_id = st.session_state.get('browser_id')
        if not browser_id:
            browser_id = str(uuid.uuid4())
            st.session_state['browser_id'] = browser_id
        
        identifier = f"{ip}:{browser_id}"
        if user_id:
            identifier += f":{user_id}"
        
        return hashlib.sha256(identifier.encode()).hexdigest()
    
    def get_daily_count(self, user_id: Optional[str] = None) -> int:
        """日次変換回数の取得"""
        identifier = self.get_unique_identifier(user_id)
        today = date.today()
        
        with self.conn:
            cursor = self.conn.execute("""
                SELECT SUM(count) FROM conversion_counts
                WHERE (user_id = ? OR ip_address = ? OR browser_id = ?)
                AND conversion_date = ?
            """, (identifier, identifier, identifier, today))
            
            count = cursor.fetchone()[0]
            return count if count is not None else 0
    
    def increment_count(self, user_id: Optional[str] = None) -> bool:
        """変換回数のインクリメント"""
        identifier = self.get_unique_identifier(user_id)
        today = date.today()
        current_count = self.get_daily_count(user_id)
        plan_limit = self.get_plan_limit(user_id)
        
        if current_count >= plan_limit:
            return False
        
        with self.conn:
            self.conn.execute("""
                INSERT INTO conversion_counts (user_id, ip_address, browser_id, conversion_date, count)
                VALUES (?, ?, ?, ?, 1)
                ON CONFLICT (user_id, ip_address, browser_id, conversion_date)
                DO UPDATE SET count = count + 1
            """, (identifier, identifier, identifier, today))
        
        return True
    
    def get_plan_limit(self, user_id: Optional[str] = None) -> int:
        """ユーザープランの制限値を取得"""
        if not user_id:
            return PLAN_LIMITS["free_guest"]
        
        with self.conn:
            cursor = self.conn.execute("""
                SELECT plan_type FROM user_plans WHERE user_id = ?
            """, (user_id,))
            result = cursor.fetchone()
            
            if result:
                return PLAN_LIMITS.get(result[0], PLAN_LIMITS["free_guest"])
            return PLAN_LIMITS["free_registered"]
    
    def update_plan(self, user_id: str, plan_type: str):
        """ユーザープランの更新"""
        with self.conn:
            self.conn.execute("""
                INSERT INTO user_plans (user_id, plan_type, updated_at)
                VALUES (?, ?, CURRENT_TIMESTAMP)
                ON CONFLICT (user_id) DO UPDATE SET
                    plan_type = excluded.plan_type,
                    updated_at = CURRENT_TIMESTAMP
            """, (user_id, plan_type))
    
    def adjust_count_after_registration(self, user_id: str):
        """登録後の変換回数調整（+2回）"""
        identifier = self.get_unique_identifier(user_id)
        today = date.today()
        
        with self.conn:
            # 既存の回数を取得
            cursor = self.conn.execute("""
                SELECT count FROM conversion_counts
                WHERE (user_id = ? OR ip_address = ? OR browser_id = ?)
                AND conversion_date = ?
            """, (identifier, identifier, identifier, today))
            
            current_count = cursor.fetchone()
            if current_count:
                # 既存レコードの更新（最大5回まで）
                new_count = min(current_count[0] + 2, PLAN_LIMITS["free_registered"])
                self.conn.execute("""
                    UPDATE conversion_counts
                    SET count = ?
                    WHERE (user_id = ? OR ip_address = ? OR browser_id = ?)
                    AND conversion_date = ?
                """, (new_count, identifier, identifier, identifier, today))

# グローバルなトラッカーインスタンス
tracker = ConversionTracker()

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

def check_and_increment_conversion_count(user_id: Optional[str] = None) -> bool:
    """変換回数をチェックしてインクリメント"""
    try:
        # 現在の変換回数と制限を取得
        daily_count = tracker.get_daily_count(user_id)
        limit = tracker.get_plan_limit(user_id)
        
        # 制限チェック
        if daily_count >= limit:
            st.error("本日の変換回数制限に達しました。プランをアップグレードするか、明日以降に再度お試しください。")
            return False
            
        # カウントをインクリメント
        if tracker.increment_count(user_id):
            st.session_state.conversion_success = True
            st.success("変換が完了しました！")
            
            # 画面の更新（最大1回まで）
            current_time = datetime.now()
            if (current_time - st.session_state.last_rerun_time).total_seconds() > 1:
                st.session_state.rerun_count += 1
                if st.session_state.rerun_count <= 1:
                    st.session_state.last_rerun_time = current_time
                    st.experimental_rerun()
            return True
            
        st.error("変換回数の更新に失敗しました。")
        return False
        
    except Exception as e:
        st.error(f"変換回数の処理中にエラーが発生しました: {str(e)}")
        return False

def get_user_plan(user_id):
    """ユーザーのプランを取得する関数"""
    try:
        if user_id is None:
            return "free_guest"
        
        # セッションからプラン情報を取得
        user_plan = st.session_state.get('user_plan', 'free_registered')
        
        # プレミアムユーザーの判定（仮の実装）
        premium_users = st.session_state.get('premium_users', set())
        if user_id in premium_users:
            return "premium"
        
        return user_plan
    except Exception as e:
        st.error(f"プラン情報の取得中にエラーが発生しました: {str(e)}")
        return "free_guest"  # エラー時は最も制限の厳しいプランを返す

def get_plan_limits(plan_type):
    """プランごとの制限を取得"""
    limits = {
        "premium": float('inf'),  # 無制限
        "free_registered": 5,     # ログインユーザー
        "free_guest": 3          # 未ログインユーザー
    }
    return limits.get(plan_type, 3)  # デフォルトは3回

def get_conversion_limit(user_id=None):
    """ユーザーの変換制限を取得"""
    plan = get_user_plan(user_id)
    return get_plan_limits(plan)

def process_pdf_with_ocr(image_bytes, document_type):
    """Google Cloud Vision APIを使用してOCR処理を実行"""
    try:
        # 画像をbase64エンコード
        image_base64 = base64.b64encode(image_bytes).decode('utf-8')
        
        # APIキーの取得
        api_key = os.getenv('GOOGLE_VISION_API_KEY')
        if not api_key:
            st.error("Google Cloud Vision APIの設定が見つかりません。")
            return None
            
        # APIエンドポイントの設定
        url = f"https://vision.googleapis.com/v1/images:annotate?key={api_key}"
        headers = {"Content-Type": "application/json"}
        
        # リクエストデータの作成
        data = {
            "requests": [{
                "image": {"content": image_base64},
                "features": [{"type": "DOCUMENT_TEXT_DETECTION"}]
            }]
        }
        
        # APIリクエストの送信
        response = requests.post(url, headers=headers, json=data)
        
        # レスポンスの確認
        if response.status_code != 200:
            st.error(f"OCR処理に失敗しました。")
            return None
            
        # レスポンスの解析
        result = response.json()
        if 'responses' in result and result['responses']:
            text_annotations = result['responses'][0].get('textAnnotations', [])
            if text_annotations:
                return text_annotations[0].get('description', '')
                
        return None
        
    except Exception as e:
        st.error(f"OCR処理中にエラーが発生しました: {str(e)}")
        return None

def process_pdf(uploaded_file, document_type=None, document_date=None):
    """PDFを処理してExcelに変換する関数"""
    try:
        if st.session_state.processing_pdf:
            return None, None
        st.session_state.processing_pdf = True

        # 変換回数制限のチェック
        user_id = st.session_state.get('user_id')
        if not check_and_increment_conversion_count(user_id):
            st.session_state.processing_pdf = False
            return None, None

        # PDFを画像に変換
        pdf_bytes = uploaded_file.getvalue()
        try:
            images = convert_from_bytes(pdf_bytes, first_page=1, last_page=1)
        except Exception as e:
            st.error("PDFの読み込みに失敗しました。Popplerがインストールされているか確認してください。")
            st.session_state.processing_pdf = False
            return None, None
        
        if not images:
            st.error("PDFの読み込みに失敗しました。")
            st.session_state.processing_pdf = False
            return None, None

        # 画像をバイトデータに変換
        img_byte_arr = io.BytesIO()
        images[0].save(img_byte_arr, format='PNG')
        img_bytes = img_byte_arr.getvalue()

        # まずpdfplumberでテキスト抽出を試みる
        try:
            with pdfplumber.open(uploaded_file) as pdf:
                page = pdf.pages[0]
                text_content = page.extract_text()
                if text_content and len(text_content.strip()) > 0:
                    return create_excel_file(text_content, document_type, document_date)
        except Exception:
            pass  # pdfplumberでの抽出に失敗した場合、OCRを試みる

        # テキスト抽出に失敗した場合、OCRを使用
        text_content = process_pdf_with_ocr(img_bytes, document_type)
        if not text_content:
            st.error("PDFからテキストを抽出できませんでした。")
            st.session_state.processing_pdf = False
            return None, None

        return create_excel_file(text_content, document_type, document_date)

    except Exception as e:
        st.error(f"PDFの処理中にエラーが発生しました: {str(e)}")
        return None, None
    finally:
        st.session_state.processing_pdf = False

def create_excel_file(text_content, document_type, document_date=None):
    """Excelファイルを作成して保存"""
    try:
        # ドキュメントタイプの日本語名を取得
        doc_type_names = {
            "estimate": "見積書",
            "invoice": "請求書",
            "delivery": "納品書",
            "receipt": "領収書",
            "financial": "決算書",
            "tax_return": "確定申告書",
            "other": "その他"
        }
        doc_type_ja = doc_type_names.get(document_type, "その他")
        
        # 日付の処理
        if document_date:
            date_str = document_date.strftime("%Y%m%d")
        else:
            date_str = datetime.now().strftime("%Y%m%d")
        
        # Excelワークブックの作成
        wb = Workbook()
        ws = wb.active
        ws.title = f"{doc_type_ja}_{date_str}"
        
        # フォントとスタイルの設定
        default_font = Font(name='Yu Gothic', size=11)
        header_font = Font(name='Yu Gothic', size=12, bold=True)
        title_font = Font(name='Yu Gothic', size=14, bold=True)
        
        # 基本的なセルスタイル
        normal_alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        center_alignment = Alignment(horizontal='center', vertical='center')
        
        # 罫線スタイル
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # テキストを行に分割
        lines = text_content.split('\n')
        
        # 行の高さと列幅の初期設定
        ws.row_dimensions[1].height = 30
        for col in range(1, 10):
            ws.column_dimensions[get_column_letter(col)].width = 15
        
        # データの書き込みと書式設定
        for i, line in enumerate(lines, 1):
            ws.cell(row=i, column=1, value=line)
            cell = ws[f"A{i}"]
            cell.font = default_font
            cell.alignment = normal_alignment
            cell.border = thin_border
            
            # 金額と思われる部分は右寄せに
            if any(char in line for char in ['¥', '円', '税']):
                cell.alignment = Alignment(horizontal='right', vertical='center')
        
        # タイトル行の特別な書式設定
        if len(lines) > 0:
            title_cell = ws['A1']
            title_cell.font = title_font
            title_cell.alignment = center_alignment
            title_cell.fill = PatternFill(start_color="E3F2FD", end_color="E3F2FD", fill_type="solid")
        
        # メモリ上にExcelファイルを保存
        excel_buffer = io.BytesIO()
        wb.save(excel_buffer)
        excel_buffer.seek(0)
        
        # ファイル名の生成
        file_name = f"{doc_type_ja}_{date_str}.xlsx"
        
        return excel_buffer, file_name
    except Exception as e:
        st.error(f"Excelファイルの作成中にエラーが発生しました: {str(e)}")
        return None, None

def get_document_type_label(doc_type):
    """ドキュメントタイプのコードから表示用ラベルを取得"""
    type_map = {
        "estimate": "見積書",
        "invoice": "請求書",
        "delivery": "納品書",
        "receipt": "領収書",
        "financial": "決算書",
        "tax_return": "確定申告書",
        "other": "その他"
    }
    return type_map.get(doc_type, "不明な書類")

def display_conversion_count():
    """変換回数を表示"""
    if 'conversion_count' not in st.session_state:
        st.session_state.conversion_count = 0
    
    if st.session_state.logged_in:
        if st.session_state.get('is_premium', False):
            st.info("本日の変換回数：無制限（有料プラン）")
        else:
            st.info(f"本日の変換回数：{st.session_state.conversion_count} / 5回（無料プラン・登録済）")
    else:
        st.info(f"本日の変換回数：{st.session_state.conversion_count} / 3回（未登録）")

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
            background: #ffffff;
            font-size: 16px;
            color: #333;
            transition: all 0.2s ease;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            display: flex;
            align-items: center;
            justify-content: center;
            cursor: pointer;
            user-select: none;
        }
        
        /* ホバー時のスタイル */
        div[data-testid="stRadio"] label:hover {
            border-color: #2196F3;
            box-shadow: 0 4px 8px rgba(33,150,243,0.2);
        }
        
        /* 選択時のスタイル */
        div[data-testid="stRadio"] label[data-checked="true"] {
            border-color: #2196F3;
            background: #e3f2fd;
            color: #1565C0;
            box-shadow: 0 4px 8px rgba(33,150,243,0.2);
        }
        
        /* 2カラムレイアウト */
        div[data-testid="stRadio"] > div {
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 10px;
        }

        /* 変換ボタンのスタイル */
        .stButton > button {
            width: 100%;
            padding: 15px 30px;
            background: linear-gradient(145deg, #2196F3 0%, #1976D2 100%);
            color: white;
            border: none;
            border-radius: 25px;
            font-size: 18px;
            font-weight: bold;
            transition: all 0.3s ease;
            box-shadow: 0 4px 6px rgba(33,150,243,0.2);
            margin-top: 20px;
        }

        .stButton > button:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 12px rgba(33,150,243,0.3);
            background: linear-gradient(145deg, #1976D2 0%, #1565C0 100%);
        }

        /* エラーメッセージのスタイル */
        .stAlert {
            background: #ffebee;
            color: #c62828;
            padding: 1rem;
            border-radius: 8px;
            border-left: 4px solid #c62828;
            margin: 1rem 0;
        }

        /* 成功メッセージのスタイル */
        .stSuccess {
            background: #e8f5e9;
            color: #2e7d32;
            padding: 1rem;
            border-radius: 8px;
            border-left: 4px solid #2e7d32;
            margin: 1rem 0;
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

def create_preview(uploaded_file):
    """PDFのプレビューを生成"""
    try:
        # PDFを画像に変換
        pdf_bytes = uploaded_file.getvalue()
        try:
            images = convert_from_bytes(pdf_bytes, first_page=1, last_page=1)
            if not images:
                st.error("PDFの読み込みに失敗しました。")
                return None
            
            # 画像をバイトデータに変換
            img_byte_arr = io.BytesIO()
            images[0].save(img_byte_arr, format='PNG')
            img_bytes = img_byte_arr.getvalue()
            
            # プレビューを表示
            st.image(img_bytes, caption="PDFプレビュー", use_column_width=True)
            return True
            
        except Exception as e:
            st.error(f"プレビューの生成中にエラーが発生しました: {str(e)}")
            if "poppler" in str(e).lower():
                st.info("💡 Popplerのインストールが必要です。以下のコマンドでインストールできます：\n```\nbrew install poppler\n```")
            return None
            
    except Exception as e:
        st.error(f"プレビューの生成中にエラーが発生しました: {str(e)}")
        return None

def create_hero_section():
    """ヒーローセクションを作成"""
    st.title("📄 PDF to Excel 変換ツール")
    st.markdown("""
    PDFファイルを簡単にExcelファイルに変換できます。
    以下の機能を提供しています：
    
    - 📝 請求書、納品書、見積書などのPDFをExcelに変換
    - 🔍 OCR機能による文字認識
    - 📊 表形式のデータを自動でExcelに整形
    - 🎨 見やすいレイアウトで出力
    """)

def create_login_section():
    """ログインセクションを作成"""
    if 'logged_in' not in st.session_state:
        st.session_state.logged_in = False
    
    if not st.session_state.logged_in:
        st.markdown("### 🔐 ログイン")
        with st.form("login_form"):
            username = st.text_input("ユーザー名")
            password = st.text_input("パスワード", type="password")
            submitted = st.form_submit_button("ログイン")
            
            if submitted:
                if username and password:
                    # ここにログイン認証のロジックを実装
                    st.session_state.logged_in = True
                    st.session_state.username = username
                    st.success("ログイン成功！")
                    st.rerun()
                else:
                    st.error("ユーザー名とパスワードを入力してください。")
        
        st.markdown("---")
        st.markdown("### 📝 新規登録")
        if st.button("アカウントを作成", type="primary"):
            st.session_state.show_register = True
            st.rerun()
    else:
        st.markdown("### 👤 ログイン済み")
        st.markdown(f"ようこそ、{st.session_state.username}さん！")
        if st.button("ログアウト"):
            st.session_state.logged_in = False
            st.session_state.username = None
            st.rerun()

def create_register_section():
    """新規登録セクションを作成"""
    st.markdown("### 📝 新規アカウント登録")
    with st.form("register_form"):
        new_username = st.text_input("ユーザー名")
        new_password = st.text_input("パスワード", type="password")
        confirm_password = st.text_input("パスワード（確認）", type="password")
        submitted = st.form_submit_button("登録")
        
        if submitted:
            if new_username and new_password and confirm_password:
                if new_password == confirm_password:
                    # ここに新規登録のロジックを実装
                    st.success("アカウントが作成されました！")
                    st.session_state.show_register = False
                    st.rerun()
                else:
                    st.error("パスワードが一致しません。")
            else:
                st.error("すべての項目を入力してください。")
    
    if st.button("ログイン画面に戻る"):
        st.session_state.show_register = False
        st.rerun()

def main():
    """メイン関数"""
    create_hero_section()
    
    # 新規登録画面の表示制御
    if 'show_register' not in st.session_state:
        st.session_state.show_register = False
    
    if st.session_state.show_register:
        create_register_section()
    else:
        create_login_section()
    
    # 変換回数の表示（最上部）
    display_conversion_count()
    
    # 2カラムレイアウト
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.subheader("ファイルをアップロード")
        
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
                if not st.session_state.processing_pdf:  # 処理中でない場合のみ実行
                    if check_and_increment_conversion_count(st.session_state.get('user_id')):
                        try:
                            excel_data, file_name = process_pdf(uploaded_file, document_type, document_date)
                            # 変換履歴を保存
                            save_conversion_history(
                                st.session_state.get('user_id'),
                                document_type,
                                document_date.strftime('%Y-%m-%d') if document_date else None,
                                file_name,
                                "success"
                            )
                            # 変換回数の表示を更新
                            display_conversion_count()
                            
                            st.download_button(
                                label="Excelファイルをダウンロード",
                                data=excel_data,
                                file_name=file_name,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        except Exception as e:
                            st.error(f"処理中にエラーが発生しました: {str(e)}")
                            # エラー履歴を保存
                            save_conversion_history(
                                st.session_state.get('user_id'),
                                document_type,
                                document_date.strftime('%Y-%m-%d') if document_date else None,
                                file_name,
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

# メイン処理部分
if st.session_state.conversion_success:
    st.session_state.conversion_success = False
    st.session_state.rerun_count = 0

if __name__ == "__main__":
    main() 