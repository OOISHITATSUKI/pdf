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

def increment_conversion_count(user_id: str) -> bool:
    """変換回数のインクリメント（バックエンド側）"""
    try:
        # 再度制限チェック
        if not check_conversion_limit(user_id):
            return False
            
        return tracker.increment_count(user_id)
    except Exception as e:
        st.error(f"変換回数の更新中にエラーが発生しました: {str(e)}")
        return False

def check_conversion_limit(user_id: Optional[str] = None) -> bool:
    """変換制限のチェック（バックエンド側）"""
    try:
        daily_count = tracker.get_daily_count(user_id)
        limit = tracker.get_plan_limit(user_id)
        
        # デバッグ情報の出力
        st.write(f"現在の変換回数: {daily_count}/{limit}")
        
        return daily_count < limit
    except Exception as e:
        st.error(f"変換制限のチェック中にエラーが発生しました: {str(e)}")
        return False

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
            if not st.session_state.processing_pdf:  # 処理中でない場合のみ実行
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
            st.error("APIキーが設定されていません。")
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
            st.error(f"APIリクエストが失敗しました: {response.status_code}")
            return None

        # レスポンスの解析
        result = response.json()
        
        # テキストの抽出
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
            return None
        st.session_state.processing_pdf = True

        # 変換回数制限のチェック
        user_id = st.session_state.get('user_id')
        if not check_conversion_limit(user_id):
            st.error("本日の変換回数制限に達しました。プランをアップグレードするか、明日以降に再度お試しください。")
            return None

        # PDFを画像に変換
        pdf_bytes = uploaded_file.getvalue()
        try:
            images = convert_from_bytes(pdf_bytes, first_page=1, last_page=1)
        except Exception as e:
            st.error("PDFの読み込みに失敗しました。Popplerがインストールされているか確認してください。")
            return None
        
        if not images:
            st.error("PDFの読み込みに失敗しました。")
            return None

        # 画像をバイトデータに変換
        img_byte_arr = io.BytesIO()
        images[0].save(img_byte_arr, format='PNG')
        img_bytes = img_byte_arr.getvalue()

        # ユーザープランに応じてOCR処理を選択
        user_plan = get_user_plan(user_id)
        if user_plan in ['premium_basic', 'premium_pro']:
            # 有料ユーザーはGoogle Cloud Vision APIを使用
            text_content = process_pdf_with_ocr(img_bytes, document_type)
        else:
            # 無料ユーザーはpdfplumberを使用
            try:
                with pdfplumber.open(uploaded_file) as pdf:
                    page = pdf.pages[0]
                    text_content = page.extract_text()
            except Exception as e:
                st.error("PDFのテキスト抽出に失敗しました。画像のみのPDFや、スキャンされたPDFの場合は有料プランでのOCR処理をお試しください。")
                return None

        if not text_content:
            st.error("このPDFは読み取れませんでした。画像のみのPDFや、スキャンされたPDFの場合は有料プランでのOCR処理をお試しください。")
            return None

        # Excelファイルの作成
        wb = Workbook()
        ws = wb.active

        # シート名の設定
        sheet_name = f"{get_document_type_label(document_type)}_{document_date.strftime('%Y-%m-%d') if document_date else 'unknown_date'}"
        ws.title = sheet_name[:31]

        # テキストをExcelに書き込み
        for i, line in enumerate(text_content.split('\n'), 1):
            ws.cell(row=i, column=1, value=line)

        # 列幅の自動調整
        for column_cells in ws.columns:
            max_length = 0
            column = column_cells[0].column
            for cell in column_cells:
                if cell.value:
                    try:
                        if not isinstance(cell, MergedCell):
                            length = len(str(cell.value))
                            max_length = max(max_length, length)
                    except:
                        pass
            adjusted_width = max(max_length + 2, 8) * 1.2
            ws.column_dimensions[get_column_letter(column)].width = adjusted_width

        # 一時ファイルとして保存
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_excel:
            wb.save(temp_excel.name)
            with open(temp_excel.name, 'rb') as f:
                excel_data = f.read()

        # 一時ファイルの削除
        os.unlink(temp_excel.name)

        # 変換成功時にカウントをインクリメント
        if increment_conversion_count(user_id):
            st.session_state.conversion_success = True
            st.success("変換が完了しました！")
            current_time = datetime.now()
            if (current_time - st.session_state.last_rerun_time).total_seconds() > 1:
                st.session_state.rerun_count += 1
                if st.session_state.rerun_count <= 1:  # 最大1回まで
                    st.session_state.last_rerun_time = current_time
                    st.experimental_rerun()
        else:
            st.error("変換回数の更新に失敗しました。")

        return excel_data

    except Exception as e:
        st.error(f"PDFの処理中にエラーが発生しました: {str(e)}")
        return None
    finally:
        st.session_state.processing_pdf = False

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
    """変換回数の表示（フロントエンド側）"""
    try:
        user_id = st.session_state.get('user_id')
        daily_count = tracker.get_daily_count(user_id)
        limit = tracker.get_plan_limit(user_id)
        
        if limit == float('inf'):
            st.markdown("📊 **変換回数制限**: 無制限")
        else:
            remaining = limit - daily_count
            plan_name = PLAN_NAMES.get(
                st.session_state.get('user_plan', 'free_guest'),
                "無料プラン（未登録）"
            )
            
            st.markdown(f"📊 **本日の残り変換回数**: {remaining} / {limit}回 ({plan_name})")
            
            # 警告表示
            if remaining <= 1:
                st.warning("⚠️ 本日の変換回数が残りわずかです。プランをアップグレードすると変換回数が増加します。")
                
                # プラン別の案内
                if not user_id:
                    st.info("💡 アカウント登録で、本日の残り回数が2回増加します！")
                elif st.session_state.get('user_plan') == 'free_registered':
                    st.info("💡 $5プランにアップグレードで、1日1000回まで変換可能になります！")
                elif st.session_state.get('user_plan') == 'premium_basic':
                    st.info("💡 $20プランにアップグレードで、無制限で変換可能になります！")
    
    except Exception as e:
        st.error(f"変換回数の取得中にエラーが発生しました: {str(e)}")
        # エラー時はデフォルト値を表示
        st.markdown("📊 **本日の残り変換回数**: 3 / 3回 (無料プラン)")

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
        }

        .stButton > button:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 12px rgba(33,150,243,0.3);
            background: linear-gradient(145deg, #1976D2 0%, #1565C0 100%);
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
    """PDFのプレビューを生成する関数"""
    try:
        if uploaded_file is not None:
            # PDFをバイトデータとして読み込み
            pdf_bytes = uploaded_file.getvalue()
            
            # PDF2Imageを使用して最初のページを画像に変換
            images = convert_from_bytes(
                pdf_bytes,
                first_page=1,
                last_page=1,
                dpi=150,
                fmt='PNG'
            )
            
            if images:
                # 最初のページの画像をバイトストリームに変換
                img_byte_arr = io.BytesIO()
                images[0].save(img_byte_arr, format='PNG')
                return img_byte_arr.getvalue()
        return None
    except Exception as e:
        st.error(f"プレビューの生成中にエラーが発生しました: {str(e)}")
        return None

def main():
    """メイン関数"""
    create_hero_section()
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
                    if check_conversion_limit(st.session_state.get('user_id')):
                        try:
                            excel_data = process_pdf(uploaded_file, document_type, document_date)
                            # 変換履歴を保存
                            save_conversion_history(
                                st.session_state.get('user_id'),
                                document_type,
                                document_date.strftime('%Y-%m-%d') if document_date else None,
                                uploaded_file.name,
                                "success"
                            )
                            # 変換回数を更新
                            increment_conversion_count(st.session_state.get('user_id'))
                            # 変換回数の表示を更新
                            display_conversion_count()
                            
                            st.download_button(
                                label="Excelファイルをダウンロード",
                                data=excel_data,
                                file_name=f"{get_document_type_label(document_type)}_{document_date.strftime('%Y-%m-%d') if document_date else 'unknown_date'}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        except Exception as e:
                            st.error(f"処理中にエラーが発生しました: {str(e)}")
                            # エラー履歴を保存
                            save_conversion_history(
                                st.session_state.get('user_id'),
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

# メイン処理部分
if st.session_state.conversion_success:
    st.session_state.conversion_success = False
    st.session_state.rerun_count = 0

if __name__ == "__main__":
    main() 