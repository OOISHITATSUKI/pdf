import streamlit as st
import pdfplumber
import pandas as pd
import tempfile
import os
from datetime import datetime, timedelta
import stripe
import hashlib
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import sqlite3

# Stripe設定（実際の値に置き換えが必要）
stripe.api_key = st.secrets["stripe"]["api_key"]
PRICE_ID = st.secrets["stripe"]["price_id"]

# ページ設定
st.set_page_config(
    page_title="PDF to Excel 変換ツール｜無料でPDFの表をExcelに変換",
    page_icon="📄",
    layout="wide"
)

# セッション状態の初期化
if 'user_state' not in st.session_state:
    st.session_state.user_state = {
        'is_logged_in': False,
        'is_premium': False,
        'email': None,
        'stored_files': [],
        'conversion_count': 0
    }

# ユーザーデータベース（実際の実装ではデータベースを使用）
if 'users' not in st.session_state:
    st.session_state.users = {}

# データベース初期化
def init_db():
    conn = sqlite3.connect('user_data.db')
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS users (
            email TEXT PRIMARY KEY,
            password TEXT NOT NULL,
            is_premium BOOLEAN DEFAULT FALSE,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            subscription_end_date TIMESTAMP
        )
    ''')
    conn.commit()
    conn.close()

# パスワードのハッシュ化
def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

# ユーザー登録
def register_user(email, password):
    try:
        conn = sqlite3.connect('user_data.db')
        c = conn.cursor()
        
        # メールアドレスの重複チェック
        c.execute('SELECT * FROM users WHERE email = ?', (email,))
        if c.fetchone() is not None:
            conn.close()
            return False, "このメールアドレスは既に登録されています"
        
        # 新規ユーザーの登録
        hashed_password = hash_password(password)
        c.execute('INSERT INTO users (email, password) VALUES (?, ?)',
                 (email, hashed_password))
        conn.commit()
        conn.close()
        return True, "登録が完了しました"
    except Exception as e:
        return False, f"登録中にエラーが発生しました: {str(e)}"

# ログイン認証
def login_user(email, password):
    try:
        conn = sqlite3.connect('user_data.db')
        c = conn.cursor()
        
        # ユーザー検索
        c.execute('SELECT * FROM users WHERE email = ? AND password = ?',
                 (email, hash_password(password)))
        user = c.fetchone()
        conn.close()
        
        if user is not None:
            return True, "ログインしました"
        else:
            return False, "メールアドレスまたはパスワードが正しくありません"
    except Exception as e:
        return False, f"ログイン中にエラーが発生しました: {str(e)}"

# データベースの初期化
init_db()

def create_checkout_session(email):
    """Stripe決済セッションの作成"""
    try:
        checkout_session = stripe.checkout.Session.create(
            payment_method_types=['card'],
            line_items=[{
                'price': PRICE_ID,
                'quantity': 1,
            }],
            mode='subscription',
            success_url=YOUR_DOMAIN + '/success',
            cancel_url=YOUR_DOMAIN + '/cancel',
            customer_email=email,
        )
        return checkout_session.url
    except Exception as e:
        st.error(f"決済セッションの作成に失敗しました: {str(e)}")
        return None

def send_excel_email(email, excel_file):
    """メールでExcelファイルを送信"""
    if st.session_state.user_state['is_premium']:
        try:
            # メール送信のロジック（SMTPサーバーの設定が必要）
            pass
        except Exception as e:
            st.error(f"メール送信に失敗しました: {str(e)}")

def store_conversion(pdf_file, excel_file):
    """変換ファイルの保存（30日間）"""
    if st.session_state.user_state['is_premium']:
        current_time = datetime.now()
        expiry_time = current_time + timedelta(days=30)
        
        file_info = {
            'pdf_name': pdf_file.name,
            'excel_path': excel_file,
            'created_at': current_time,
            'expires_at': expiry_time
        }
        
        st.session_state.user_state['stored_files'].append(file_info)

# 認証UI
def show_auth_ui():
    st.sidebar.markdown("### アカウント管理")
    
    if not st.session_state.user_state['is_logged_in']:
        tab1, tab2 = st.sidebar.tabs(["ログイン", "新規登録"])
        
        with tab1:
            with st.form("login_form"):
                login_email = st.text_input("メールアドレス", key="login_email")
                login_password = st.text_input("パスワード", type="password", key="login_password")
                login_submit = st.form_submit_button("ログイン")
                
                if login_submit:
                    success, message = login_user(login_email, login_password)
                    if success:
                        st.session_state.user_state['is_logged_in'] = True
                        st.session_state.user_state['email'] = login_email
                        st.success(message)
                        st.experimental_rerun()
                    else:
                        st.error(message)
        
        with tab2:
            with st.form("register_form"):
                reg_email = st.text_input("メールアドレス", key="reg_email")
                reg_password = st.text_input("パスワード", type="password", key="reg_password")
                reg_password_confirm = st.text_input("パスワード（確認）", type="password")
                register_submit = st.form_submit_button("新規登録")
                
                if register_submit:
                    if reg_password != reg_password_confirm:
                        st.error("パスワードが一致しません")
                    else:
                        success, message = register_user(reg_email, reg_password)
                        if success:
                            st.success(message)
                            # 自動ログイン
                            st.session_state.user_state['is_logged_in'] = True
                            st.session_state.user_state['email'] = reg_email
                            st.experimental_rerun()
                        else:
                            st.error(message)
    
    else:
        st.sidebar.markdown(f"### ようこそ！")
        st.sidebar.markdown(f"ログイン中: {st.session_state.user_state['email']}")
        
        if not st.session_state.user_state['is_premium']:
            st.sidebar.markdown("### 🌟 プレミアムにアップグレード")
            if st.sidebar.button("プレミアム会員に登録"):
                # Stripe決済ページへのリンク
                checkout_url = create_checkout_session(st.session_state.user_state['email'])
                if checkout_url:
                    st.markdown(f"[決済ページへ進む]({checkout_url})")
        
        if st.sidebar.button("ログアウト"):
            st.session_state.user_state = {
                'is_logged_in': False,
                'is_premium': False,
                'email': None,
                'stored_files': [],
                'conversion_count': 0
            }
            st.experimental_rerun()

# メイン処理部分
def process_files(uploaded_files):
    """ファイル処理のメイン関数"""
    max_files = 10 if st.session_state.user_state['is_premium'] else 3
    
    if len(uploaded_files) > max_files:
        st.error(f"⚠️ 一度に変換できるのは最大{max_files}ファイルまでです")
        return
    
    if not st.session_state.user_state['is_premium'] and st.session_state.user_state['conversion_count'] >= 3:
        st.error("⚠️ 無料プランの変換可能回数を超えました。プレミアムにアップグレードすると無制限で変換できます。")
        return
    
    for uploaded_file in uploaded_files:
        with st.spinner(f'{uploaded_file.name} を処理中...'):
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
                    tmp_file.write(uploaded_file.getvalue())
                    tmp_path = tmp_file.name

                # 広告表示（処理中）
                st.markdown('<div class="ad-container">広告スペース</div>', unsafe_allow_html=True)

                with pdfplumber.open(tmp_path) as pdf:
                    # テーブル認識精度の強化
                    all_tables = []
                    for page in pdf.pages:
                        try:
                            # シンプルな抽出方法を試す
                            table = page.extract_table()
                            if table:
                                all_tables.extend(table)
                            else:
                                # テキストとして抽出を試みる
                                text = page.extract_text()
                                if text:
                                    # テキストを1列のデータとして追加
                                    all_tables.append([text])
                        except Exception as e:
                            st.warning(f"ページの処理中にエラーが発生しました: {str(e)}")
                            continue

                    if all_tables:
                        # データフレームの作成と最適化
                        df = pd.DataFrame(all_tables)
                        # 空の行と列を削除
                        df = df.dropna(how='all').dropna(axis=1, how='all')
                        
                        st.markdown(f"### {uploaded_file.name} のプレビュー")
                        st.dataframe(df, use_container_width=True)
                        
                        # Excelファイル作成
                        excel_file = f'converted_{uploaded_file.name}.xlsx'
                        df.to_excel(excel_file, index=False)
                        
                        # プレミアム機能
                        if st.session_state.user_state['is_premium']:
                            store_conversion(uploaded_file, excel_file)
                            col1, col2 = st.columns(2)
                            with col1:
                                if st.button("📧 メールで受け取る"):
                                    send_excel_email(st.session_state.user_state['email'], excel_file)
                        else:
                            st.session_state.user_state['conversion_count'] += 1

                        with open(excel_file, 'rb') as f:
                            st.download_button(
                                label=f"📥 {uploaded_file.name} をダウンロード",
                                data=f,
                                file_name=excel_file,
                                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                            )
                        os.remove(excel_file)
                    else:
                        st.warning(f"⚠️ {uploaded_file.name} からテーブルデータが見つかりませんでした")

            except Exception as e:
                st.error(f"❌ {uploaded_file.name} の処理中にエラーが発生しました")
            finally:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)

# メインアプリケーション
show_auth_ui()

if not st.session_state.user_state['is_premium']:
    st.markdown("""
    ### 🌟 プレミアム機能 (月額500円)
    - ✨ 無制限の変換回数
    - 📦 一度に10ファイルまで変換可能
    - 📧 変換したファイルをメールで受信
    - 💾 30日間のファイル保存
    - 🚫 広告非表示
    """)

# ファイルアップロード
uploaded_files = st.file_uploader(
    "PDFファイルを選択",
    type=['pdf'],
    accept_multiple_files=True
)

if uploaded_files:
    process_files(uploaded_files)

# 保存されたファイルの表示（プレミアムユーザーのみ）
if st.session_state.user_state['is_premium'] and st.session_state.user_state['stored_files']:
    st.markdown("### 保存されたファイル")
    for file in st.session_state.user_state['stored_files']:
        if datetime.now() < file['expires_at']:
            st.download_button(
                f"📥 {file['pdf_name']}",
                data=open(file['excel_path'], 'rb'),
                file_name=f"converted_{file['pdf_name']}.xlsx"
            )

# テーブル認識精度強化のための関数
def extract_table_with_enhanced_recognition(page):
    """
    複数の抽出方法を試行して最適な結果を返す
    """
    try:
        # 方法1: 標準的な抽出
        table = page.extract_table()
        if table and is_valid_table(table):
            return table

        # 方法2: カスタム設定での抽出
        table = page.extract_table(
            table_settings={
                "vertical_strategy": "text",
                "horizontal_strategy": "text",
                "snap_tolerance": 3,
                "join_tolerance": 3,
                "edge_min_length": 3,
                "min_words_vertical": 3,
            }
        )
        if table and is_valid_table(table):
            return table

        # 方法3: 線による抽出
        table = page.extract_table(
            table_settings={
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
            }
        )
        return table if table and is_valid_table(table) else None

    except Exception:
        return None

def enhance_table_structure(df):
    """
    pandas DataFrameの構造を改善
    """
    # NaN値の処理
    df = df.fillna('')
    
    # 重複列の処理
    df = df.loc[:, ~df.columns.duplicated()]
    
    # 空行の削除
    df = df.dropna(how='all')
    
    # 列名の正規化
    df.columns = [str(col).strip() for col in df.columns]
    
    # データの正規化
    df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
    
    return df

def save_enhanced_excel(df, filename):
    """
    整形されたExcelファイルを保存
    """
    writer = pd.ExcelWriter(filename, engine='openpyxl')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    
    # ワークシートの取得
    worksheet = writer.sheets['Sheet1']
    
    # 列幅の自動調整
    for column in worksheet.columns:
        max_length = 0
        column = [cell for cell in column]
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        worksheet.column_dimensions[column[0].column_letter].width = adjusted_width
    
    writer.save()

# サポート情報
with st.expander("📌 サポート対象PDFについて"):
    st.markdown("""
    ### 対応PDFの種類
    - ✨ 表形式のデータを含むPDF
    - 📄 通常のテキストデータを含むPDF
    - 📊 複合的なコンテンツを含むPDF
    
    ### 注意事項
    - ⚠️ スキャンされたPDFや画像化されたPDFは変換できない場合があります
    - 🔒 パスワード保護されたPDFは処理できません
    """)

# フッター
st.markdown("""
<div class="footer">
    <p>© 2025 PDF to Excel変換ツール</p>
    <p style="font-size: 0.9rem;">プライバシーを重視した無料のオンラインPDF変換サービス</p>
</div>
""", unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True) 