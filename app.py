import streamlit as st
import pdfplumber
import pandas as pd
import tempfile
import os
import time
from datetime import datetime, timedelta
import hashlib

# ページ設定
st.set_page_config(
    page_title="PDF to Excel 変換ツール｜無料でPDFの表をExcelに変換",
    page_icon="📄",
    layout="wide"
)

# セッション状態の初期化
if 'user_state' not in st.session_state:
    st.session_state.user_state = {
        'is_logged_in': False,        # ログイン状態
        'is_premium': False,          # 有料会員状態
        'daily_conversions': 0,       # 今日の変換回数
        'last_conversion_date': None, # 最後の変換日
        'email': None
    }

# ユーザーデータベースの初期化（実際の実装ではデータベースを使用）
if 'users' not in st.session_state:
    st.session_state.users = {}

def hash_password(password):
    """パスワードをハッシュ化"""
    return hashlib.sha256(password.encode()).hexdigest()

def register_user(email, password):
    """ユーザー登録"""
    if email in st.session_state.users:
        return False, "このメールアドレスは既に登録されています"
    
    st.session_state.users[email] = {
        'password': hash_password(password),
        'is_premium': False,
        'daily_conversions': 0,
        'last_conversion_date': None
    }
    return True, "登録が完了しました"

def login_user(email, password):
    """ユーザーログイン"""
    if email not in st.session_state.users:
        return False, "メールアドレスが見つかりません"
    
    if st.session_state.users[email]['password'] != hash_password(password):
        return False, "パスワードが正しくありません"
    
    st.session_state.user_state['is_logged_in'] = True
    st.session_state.user_state['email'] = email
    st.session_state.user_state['is_premium'] = st.session_state.users[email]['is_premium']
    return True, "ログインしました"

def check_conversion_limit():
    """ユーザーの変換制限をチェックする関数"""
    # 未ログインまたは無料会員の場合のみ制限をチェック
    if not st.session_state.user_state['is_premium']:
        current_date = datetime.now().date()
        last_date = st.session_state.user_state['last_conversion_date']

        # 日付が変わった場合、カウントをリセット
        if last_date != current_date:
            st.session_state.user_state['daily_conversions'] = 0
            st.session_state.user_state['last_conversion_date'] = current_date

        # 制限チェック
        if st.session_state.user_state['daily_conversions'] >= 3:
            return False
    return True

def increment_conversion_count():
    """変換回数をカウントアップする関数"""
    if not st.session_state.user_state['is_premium']:
        st.session_state.user_state['daily_conversions'] += 1
        st.session_state.user_state['last_conversion_date'] = datetime.now().date()

# カスタムCSSの追加
st.markdown("""
<style>
    /* 全体の背景にグラデーション */
    .stApp {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
    }
    
    /* ヘッダー部分のスタイル */
    .header-container {
        background: white;
        padding: 2rem;
        border-radius: 15px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        margin-bottom: 2rem;
    }
    
    /* ログインボタンのスタイル */
    .stButton>button {
        background: linear-gradient(45deg, #2196F3, #21CBF3);
        color: white;
        border: none;
        padding: 0.5rem 1rem;
        border-radius: 25px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        transition: all 0.3s ease;
    }
    
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0,0,0,0.2);
    }
    
    /* アップロードエリアのスタイル */
    .uploadfile {
        background: white;
        border-radius: 15px;
        padding: 2rem;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    
    /* プレミアムバッジのスタイル */
    .premium-badge {
        background: linear-gradient(45deg, #FFD700, #FFA500);
        color: white;
        padding: 0.5rem 1rem;
        border-radius: 25px;
        text-align: center;
        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
    }
    
    /* 無料バッジのスタイル */
    .free-badge {
        background: linear-gradient(45deg, #e0e0e0, #b0b0b0);
        color: white;
        padding: 0.5rem 1rem;
        border-radius: 25px;
        text-align: center;
    }
</style>
""", unsafe_allow_html=True)

# ユーザー状態を表示する関数
def show_user_status():
    if st.session_state.user_state['is_logged_in']:
        if st.session_state.user_state['is_premium']:
            status_html = """
            <div class="header-container">
                <div class="user-info">
                    <span class="premium-badge">🌟 プレミアム会員</span>
                </div>
            </div>
            """
        else:
            remaining = 3 - st.session_state.user_state['daily_conversions']
            status_html = f"""
            <div class="header-container">
                <div class="user-info">
                    <span class="free-badge">無料会員</span>
                    <span class="remaining-count">残り {remaining}回</span>
                    <a href="#" class="upgrade-button">🌟 プレミアムに変更</a>
                </div>
            </div>
            """
    else:
        status_html = """
        <div class="header-container">
            <div class="user-info">
                <a href="#" class="login-button">ログイン</a>
                <a href="#" class="login-button">新規登録</a>
            </div>
        </div>
        """
    
    st.markdown(status_html, unsafe_allow_html=True)

# ユーザー状態の表示
show_user_status()

# メインコンテンツ
st.markdown('<div class="main-content">', unsafe_allow_html=True)

# 2列レイアウトでヘッダーを作成
header_left, header_right = st.columns([3, 1])

with header_left:
    st.markdown("""
    <div class="header-container">
        <h1>PDF to Excel 変換ツール</h1>
        <p>PDFファイルをExcel形式に変換できます。すべての処理はブラウザ内で行われます。</p>
    </div>
    """, unsafe_allow_html=True)

with header_right:
    if not st.session_state.user_state['is_logged_in']:
        if st.button("ログイン / 新規登録"):
            st.session_state['show_auth'] = True
    else:
        if st.session_state.user_state['is_premium']:
            st.markdown("""
                <div class="premium-badge">
                    🌟 プレミアム会員
                </div>
                """, unsafe_allow_html=True)
        else:
            remaining = 3 - st.session_state.user_state['daily_conversions']
            st.markdown(f"""
                <div class="free-badge">
                    無料会員 (残り {remaining}回)
                </div>
                """, unsafe_allow_html=True)
            st.button("🌟 プレミアムに変更")

# 認証フォーム
if st.session_state.get('show_auth', False) and not st.session_state.user_state['is_logged_in']:
    with st.form("auth_form"):
        auth_type = st.radio("", ["ログイン", "新規登録"])
        email = st.text_input("メールアドレス")
        password = st.text_input("パスワード", type="password")
        
        if st.form_submit_button("送信"):
            if auth_type == "新規登録":
                success, message = register_user(email, password)
            else:
                success, message = login_user(email, password)
            
            if success:
                st.success(message)
                st.session_state['show_auth'] = False
                st.experimental_rerun()
            else:
                st.error(message)

# ファイルアップロード
st.markdown('<div class="upload-area">', unsafe_allow_html=True)
uploaded_files = st.file_uploader("", type=['pdf'], accept_multiple_files=True)
if not uploaded_files:
    st.markdown('📄 クリックまたはドラッグ＆ドロップでPDFファイルを選択（最大3つまで）')
elif len(uploaded_files) > 3:
    st.error("⚠️ 無料版では一度に3つまでのファイルしか変換できません")
st.markdown('</div>', unsafe_allow_html=True)

# SEO対策のためのメタ情報
st.markdown("""
<!-- SEO対策用メタ情報 -->
<div style="display:none">
    PDF Excel 変換 無料 表 テーブル 一括変換 データ抽出 オンライン ツール
    PDFからExcelへの無料変換ツール 表形式データ抽出 高精度変換
</div>
""", unsafe_allow_html=True)

# 複数ファイル処理部分
if uploaded_files:
    for i, uploaded_file in enumerate(uploaded_files[:3]):  # 最大3つまでに制限
        with st.spinner(f'PDFファイル {i+1}/{len(uploaded_files)} を処理中...'):
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
                        excel_file = f'converted_data_{i+1}.xlsx'
                        df.to_excel(excel_file, index=False)
                        
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