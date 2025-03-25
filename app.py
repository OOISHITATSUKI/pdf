import streamlit as st
import pdfplumber
import pandas as pd
import tempfile
import os
from datetime import datetime, timedelta
import sqlite3
import hashlib
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

# Stripe関連のコードは一時的にコメントアウト
# import stripe
# stripe.api_key = st.secrets["stripe"]["api_key"]
# PRICE_ID = st.secrets["stripe"]["price_id"]

# ページ設定
st.set_page_config(
    page_title="PDF to Excel 変換ツール｜無料でPDFの表をExcelに変換",
    page_icon="📄",
    layout="wide"
)

# セッション状態の初期化
if 'users' not in st.session_state:
    st.session_state.users = {}  # ユーザーデータを一時的に保存

if 'user_state' not in st.session_state:
    st.session_state.user_state = {
        'is_logged_in': False,
        'is_premium': False,
        'email': None,
        'daily_conversions': 0,
        'last_conversion_date': None
    }

# ユーザー登録（セッションベース）
def register_user(email, password):
    if email in st.session_state.users:
        return False, "このメールアドレスは既に登録されています"
    
    st.session_state.users[email] = {
        'password': password,
        'is_premium': False,
        'created_at': datetime.now()
    }
    return True, "登録が完了しました"

# ログイン認証（セッションベース）
def login_user(email, password):
    if email not in st.session_state.users:
        return False, "メールアドレスが見つかりません"
    
    if st.session_state.users[email]['password'] != password:
        return False, "パスワードが正しくありません"
    
    return True, "ログインしました"

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
                        st.session_state.user_state['is_premium'] = st.session_state.users[login_email]['is_premium']
                        st.success(message)
                        st.rerun()
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
                            st.session_state.user_state['is_logged_in'] = True
                            st.session_state.user_state['email'] = reg_email
                            st.rerun()
                        else:
                            st.error(message)
    
    else:
        st.sidebar.markdown(f"### ようこそ！")
        st.sidebar.markdown(f"ログイン中: {st.session_state.user_state['email']}")
        
        if not st.session_state.user_state['is_premium']:
            st.sidebar.markdown("### 🌟 プレミアムにアップグレード")
            if st.sidebar.button("プレミアム会員に登録"):
                st.sidebar.info("準備中です...")
        
        if st.sidebar.button("ログアウト"):
            st.session_state.user_state = {
                'is_logged_in': False,
                'is_premium': False,
                'email': None,
                'daily_conversions': 0,
                'last_conversion_date': None
            }
            st.rerun()

# メインアプリケーションのUI
def main():
    show_auth_ui()
    
    st.title("PDF to Excel 変換ツール")
    st.markdown("PDFファイルをExcel形式に変換できます。すべての処理はブラウザ内で行われます。")
    
    # 利用制限の表示
    if not st.session_state.user_state['is_premium']:
        if st.session_state.user_state['is_logged_in']:
            st.info(f"本日の残り変換回数: {5 - st.session_state.user_state['daily_conversions']}回")
        else:
            st.info(f"本日の残り変換回数: {3 - st.session_state.user_state['daily_conversions']}回")
    
    # ファイルアップロード
    uploaded_file = st.file_uploader("PDFファイルを選択", type=['pdf'])
    
    if uploaded_file:
        if not check_conversion_limit():
            if st.session_state.user_state['is_logged_in']:
                st.error("本日の変換可能回数（5回）を超えました。プレミアムプランへのアップグレードをご検討ください。")
            else:
                st.error("本日の変換可能回数（3回）を超えました。アカウント登録で追加の2回が利用可能になります。")
            return
        
        with st.spinner('変換中...'):
            # 一時ファイルの作成
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_pdf:
                tmp_pdf.write(uploaded_file.getvalue())
                pdf_path = tmp_pdf.name
            
            excel_path = f'converted_{uploaded_file.name}.xlsx'
            
            # 変換実行
            if convert_pdf_to_excel(pdf_path, excel_path):
                st.success("変換が完了しました！")
                
                # ダウンロードボタン
                with open(excel_path, 'rb') as f:
                    st.download_button(
                        label="📥 Excelファイルをダウンロード",
                        data=f,
                        file_name=excel_path,
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )
                
                # 変換回数をカウントアップ
                if not st.session_state.user_state['is_premium']:
                    st.session_state.user_state['daily_conversions'] += 1
            
            # 一時ファイルの削除
            try:
                os.remove(pdf_path)
                os.remove(excel_path)
            except:
                pass

# メイン処理部分
def process_files(uploaded_files):
    """ファイル処理のメイン関数"""
    max_files = 10 if st.session_state.user_state['is_premium'] else 3
    
    if len(uploaded_files) > max_files:
        st.error(f"⚠️ 一度に変換できるのは最大{max_files}ファイルまでです")
        return
    
    if not st.session_state.user_state['is_premium'] and st.session_state.user_state['daily_conversions'] >= 3:
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
                            st.session_state.user_state['daily_conversions'] += 1

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

# PDFの処理関数を修正
def process_pdf(uploaded_file):
    try:
        # 一時ファイルの作成
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_path = tmp_file.name

        tables = []
        try:
            with pdfplumber.open(tmp_path) as pdf:
                for page in pdf.pages:
                    try:
                        # テーブルの抽出を試みる
                        table = page.extract_table()
                        if table:
                            tables.extend(table)
                        else:
                            # テーブルが見つからない場合はテキストとして抽出
                            text = page.extract_text()
                            if text:
                                tables.append([text])
                    except Exception as e:
                        st.warning(f"ページの処理中にエラーが発生しました: {str(e)}")
                        continue
        finally:
            # 一時ファイルの削除を確実に行う
            if os.path.exists(tmp_path):
                os.unlink(tmp_path)

        if tables:
            # データフレームの作成と最適化
            df = pd.DataFrame(tables)
            # 空の行と列を削除
            df = df.dropna(how='all').dropna(axis=1, how='all')
            return df
        else:
            st.warning("PDFからデータを抽出できませんでした")
            return None

    except Exception as e:
        st.error(f"ファイルの処理中にエラーが発生しました: {str(e)}")
        return None

def extract_pdf_content(pdf_path):
    """PDFから詳細な情報を抽出する関数"""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            document_info = {
                'tables': [],
                'titles': [],
                'metadata': {},
                'styles': []
            }
            
            # PDFのメタデータを取得
            document_info['metadata'] = pdf.metadata
            
            for page_num, page in enumerate(pdf.pages, 1):
                # ページの基本情報
                page_info = {
                    'page_number': page_num,
                    'page_size': {'width': page.width, 'height': page.height}
                }
                
                # テキスト要素の抽出（文字サイズ、フォント情報を含む）
                words = page.extract_words(
                    keep_blank_chars=True,
                    x_tolerance=3,
                    y_tolerance=3,
                    extra_attrs=['fontname', 'size']
                )
                
                # タイトルらしき要素の検出（文字サイズが大きい要素）
                for word in words:
                    if word.get('size', 0) > 12:  # サイズの閾値は調整可能
                        document_info['titles'].append({
                            'text': word['text'],
                            'size': word['size'],
                            'font': word.get('fontname', 'unknown'),
                            'page': page_num
                        })
                
                # テーブルの詳細な抽出
                tables = page.find_tables(
                    table_settings={
                        "vertical_strategy": "text",
                        "horizontal_strategy": "text",
                        "snap_tolerance": 3,
                        "join_tolerance": 3,
                        "edge_min_length": 3,
                        "min_words_vertical": 3,
                        "min_words_horizontal": 1,
                        "keep_blank_chars": True,
                        "text_tolerance": 3,
                        "text_x_tolerance": 3,
                        "text_y_tolerance": 3
                    }
                )
                
                for table in tables:
                    table_data = table.extract()
                    if table_data:
                        # セル結合の検出
                        merged_cells = []
                        for i, row in enumerate(table_data):
                            for j, cell in enumerate(row):
                                if cell is not None:
                                    # 横方向の結合を検出
                                    span = 1
                                    while j + span < len(row) and row[j + span] is None:
                                        span += 1
                                    if span > 1:
                                        merged_cells.append({
                                            'type': 'horizontal',
                                            'row': i,
                                            'col': j,
                                            'span': span,
                                            'value': cell
                                        })
                        
                        document_info['tables'].append({
                            'page': page_num,
                            'data': table_data,
                            'merged_cells': merged_cells,
                            'position': {
                                'top': table.bbox[1],
                                'left': table.bbox[0],
                                'bottom': table.bbox[3],
                                'right': table.bbox[2]
                            }
                        })
                
                # スタイル情報の収集
                document_info['styles'].append({
                    'page': page_num,
                    'fonts': list(set(word.get('fontname', 'unknown') for word in words)),
                    'sizes': list(set(word.get('size', 0) for word in words))
                })
            
            return document_info
    
    except Exception as e:
        st.error(f"PDFの解析中にエラーが発生しました: {str(e)}")
        return None

def create_excel_with_formatting(document_info, output_path):
    """抽出した情報をフォーマット付きでExcelに出力"""
    try:
        writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
        workbook = writer.book
        
        # スタイルの定義
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 14,
            'align': 'center'
        })
        
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#D9D9D9',
            'border': 1
        })
        
        # メタデータシートの作成
        metadata_df = pd.DataFrame([document_info['metadata']])
        metadata_df.to_excel(writer, sheet_name='メタデータ', index=False)
        
        # テーブルデータの出力
        for i, table_info in enumerate(document_info['tables']):
            sheet_name = f'テーブル{i+1}'
            df = pd.DataFrame(table_info['data'])
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            worksheet = writer.sheets[sheet_name]
            
            # セル結合の適用
            for merge in table_info['merged_cells']:
                if merge['type'] == 'horizontal':
                    worksheet.merge_range(
                        merge['row'] + 1, merge['col'],
                        merge['row'] + 1, merge['col'] + merge['span'] - 1,
                        merge['value']
                    )
        
        # タイトル情報の出力
        titles_df = pd.DataFrame(document_info['titles'])
        if not titles_df.empty:
            titles_df.to_excel(writer, sheet_name='タイトル一覧', index=False)
        
        writer.close()
        return True
    
    except Exception as e:
        st.error(f"Excel作成中にエラーが発生しました: {str(e)}")
        return False

# メイン処理部分の更新
def process_pdf_file(uploaded_file):
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_path = tmp_file.name
            
            # 詳細な情報を抽出
            document_info = extract_pdf_content(tmp_path)
            
            if document_info:
                # Excel出力用の一時ファイル
                excel_path = f'converted_{uploaded_file.name}.xlsx'
                
                if create_excel_with_formatting(document_info, excel_path):
                    # プレビュー表示
                    st.success("変換が完了しました！")
                    
                    # タイトル情報の表示
                    if document_info['titles']:
                        st.write("検出されたタイトル:")
                        for title in document_info['titles']:
                            st.write(f"- {title['text']} (サイズ: {title['size']})")
                    
                    # テーブル情報の表示
                    for i, table in enumerate(document_info['tables']):
                        st.write(f"テーブル {i+1}:")
                        df = pd.DataFrame(table['data'])
                        st.dataframe(df)
                    
                    # ダウンロードボタン
                    with open(excel_path, 'rb') as f:
                        st.download_button(
                            label="📥 Excelファイルをダウンロード",
                            data=f,
                            file_name=excel_path,
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                        )
                    
                    # 一時ファイルの削除
                    os.remove(excel_path)
            
            # PDF一時ファイルの削除
            os.remove(tmp_path)
            
    except Exception as e:
        st.error(f"処理中にエラーが発生しました: {str(e)}")

def check_conversion_limit():
    """変換制限をチェック"""
    current_date = datetime.now().date()
    
    # 日付が変わった場合はカウントをリセット
    if st.session_state.user_state['last_conversion_date'] != current_date:
        st.session_state.user_state['daily_conversions'] = 0
        st.session_state.user_state['last_conversion_date'] = current_date

    if st.session_state.user_state['is_premium']:
        return True
    elif st.session_state.user_state['is_logged_in']:
        return st.session_state.user_state['daily_conversions'] < 5
    else:
        return st.session_state.user_state['daily_conversions'] < 3

def analyze_table_structure(page):
    """PDFページからテーブル構造を詳細に分析"""
    words = page.extract_words(
        keep_blank_chars=True,
        x_tolerance=3,
        y_tolerance=3,
        extra_attrs=['fontname', 'size']
    )
    
    # テーブルの検出と構造分析
    tables = page.find_tables(
        table_settings={
            "vertical_strategy": "text",
            "horizontal_strategy": "text",
            "snap_tolerance": 3,
            "join_tolerance": 3,
            "edge_min_length": 3,
            "min_words_vertical": 3,
            "min_words_horizontal": 1
        }
    )
    
    table_structures = []
    for table in tables:
        cells = []
        for row_idx, row in enumerate(table.cells):
            for col_idx, cell in enumerate(row):
                if cell is not None:
                    # セルの内容を取得
                    cell_text = cell.extract()
                    if cell_text:
                        # セルの位置情報を保存
                        cells.append({
                            'text': cell_text.strip(),
                            'row': row_idx,
                            'col': col_idx,
                            'bbox': cell.bbox,
                            'merged': False
                        })
        
        # セルの結合を検出
        merged_cells = []
        for i, cell in enumerate(cells):
            if not cell['merged']:
                # 横方向の結合をチェック
                merge_width = 1
                merge_height = 1
                
                # 横方向のチェック
                next_col = cell['col'] + 1
                while next_col < table.shape[1]:
                    if any(c['row'] == cell['row'] and c['col'] == next_col for c in cells):
                        break
                    merge_width += 1
                    next_col += 1
                
                # 縦方向のチェック
                next_row = cell['row'] + 1
                while next_row < table.shape[0]:
                    if any(c['col'] == cell['col'] and c['row'] == next_row for c in cells):
                        break
                    merge_height += 1
                    next_row += 1
                
                if merge_width > 1 or merge_height > 1:
                    merged_cells.append({
                        'start_row': cell['row'],
                        'start_col': cell['col'],
                        'end_row': cell['row'] + merge_height - 1,
                        'end_col': cell['col'] + merge_width - 1,
                        'text': cell['text']
                    })
        
        table_structures.append({
            'cells': cells,
            'merged_cells': merged_cells,
            'shape': table.shape
        })
    
    return table_structures

def create_formatted_excel(table_structures, output_path):
    """フォーマットを保持してExcelファイルを作成"""
    wb = Workbook()
    ws = wb.active
    
    # 基本的なスタイル設定
    default_font = Font(name='Yu Gothic', size=10)
    header_font = Font(name='Yu Gothic', bold=True, size=11)
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    for table_idx, table in enumerate(table_structures):
        # セルの内容を書き込み
        for cell in table['cells']:
            ws.cell(
                row=cell['row'] + 1,
                column=cell['col'] + 1,
                value=cell['text']
            ).font = default_font
        
        # セルの結合を適用
        for merge in table['merged_cells']:
            ws.merge_cells(
                start_row=merge['start_row'] + 1,
                start_column=merge['start_col'] + 1,
                end_row=merge['end_row'] + 1,
                end_column=merge['end_col'] + 1
            )
            merged_cell = ws.cell(
                row=merge['start_row'] + 1,
                column=merge['start_col'] + 1
            )
            merged_cell.value = merge['text']
            merged_cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # 列幅の調整
        for col in range(1, table['shape'][1] + 1):
            ws.column_dimensions[get_column_letter(col)].width = 15
    
    wb.save(output_path)

def process_pdf_file(uploaded_file):
    """PDFファイルを処理"""
    try:
        # 変換制限のチェック
        if not check_conversion_limit():
            if st.session_state.user_state['is_logged_in']:
                st.error("本日の変換可能回数（5回）を超えました。プレミアムプランへのアップグレードをご検討ください。")
            else:
                st.error("本日の変換可能回数（3回）を超えました。アカウント登録で追加の2回が利用可能になります。")
            return

        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_path = tmp_file.name
            
            with pdfplumber.open(tmp_path) as pdf:
                all_tables = []
                for page in pdf.pages:
                    table_structures = analyze_table_structure(page)
                    all_tables.extend(table_structures)
            
            if all_tables:
                excel_path = f'converted_{uploaded_file.name}.xlsx'
                create_formatted_excel(all_tables, excel_path)
                
                st.success("変換が完了しました！")
                
                # プレビューの表示
                for table_idx, table in enumerate(all_tables):
                    st.write(f"テーブル {table_idx + 1}:")
                    preview_data = []
                    for cell in table['cells']:
                        while len(preview_data) <= cell['row']:
                            preview_data.append([''] * table['shape'][1])
                        preview_data[cell['row']][cell['col']] = cell['text']
                    
                    df = pd.DataFrame(preview_data)
                    st.dataframe(df)
                
                # ダウンロードボタン
                with open(excel_path, 'rb') as f:
                    st.download_button(
                        label="📥 Excelファイルをダウンロード",
                        data=f,
                        file_name=excel_path,
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )
                
                os.remove(excel_path)
                
                # 変換回数をカウントアップ
                if not st.session_state.user_state['is_premium']:
                    st.session_state.user_state['daily_conversions'] += 1
            
            os.remove(tmp_path)
            
    except Exception as e:
        st.error(f"処理中にエラーが発生しました: {str(e)}")

def convert_pdf_to_excel(pdf_path, excel_path):
    """PDFをExcelに変換"""
    try:
        wb = Workbook()
        ws = wb.active
        
        with pdfplumber.open(pdf_path) as pdf:
            current_row = 1
            
            for page in pdf.pages:
                # テーブルの抽出
                tables = page.extract_tables()
                
                for table in tables:
                    # 各行の処理
                    for row_idx, row in enumerate(table):
                        # セルの結合状態を確認
                        merged_cells = []
                        current_merge = None
                        
                        for col_idx, cell in enumerate(row):
                            # セルの書き込み
                            if cell is not None:
                                cell = str(cell).strip()
                                ws.cell(row=current_row + row_idx, 
                                      column=col_idx + 1, 
                                      value=cell)
                                
                                # セルのスタイル設定
                                cell_obj = ws.cell(row=current_row + row_idx, 
                                                 column=col_idx + 1)
                                
                                # 基本的なスタイル
                                cell_obj.font = Font(name='Yu Gothic', size=10)
                                cell_obj.border = Border(
                                    left=Side(style='thin'),
                                    right=Side(style='thin'),
                                    top=Side(style='thin'),
                                    bottom=Side(style='thin')
                                )
                                
                                # ヘッダー行のスタイル
                                if row_idx == 0:
                                    cell_obj.font = Font(name='Yu Gothic', 
                                                       size=10, 
                                                       bold=True)
                                    cell_obj.fill = PatternFill(
                                        start_color='F2F2F2',
                                        end_color='F2F2F2',
                                        fill_type='solid'
                                    )
                                
                                # 数値の右寄せ
                                if cell.replace(',', '').replace('.', '').isdigit():
                                    cell_obj.alignment = Alignment(horizontal='right')
                                else:
                                    cell_obj.alignment = Alignment(horizontal='left')
                            
                            # 列幅の自動調整
                            ws.column_dimensions[get_column_letter(col_idx + 1)].width = 15
                    
                    current_row += len(table) + 1  # テーブル間に1行空ける
        
        # Excelファイルの保存
        wb.save(excel_path)
        return True
        
    except Exception as e:
        st.error(f"変換中にエラーが発生しました: {str(e)}")
        return False

# requirements.txtに追加が必要なパッケージ
"""
openpyxl
""" 