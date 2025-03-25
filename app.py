import streamlit as st
import pdfplumber
import pandas as pd
import numpy as np
import cv2
import pytesseract
from PIL import Image
import tempfile
import os
import re
from datetime import datetime

# ページ設定
st.set_page_config(
    page_title="PDF to Excel 変換ツール｜無料でPDFの表をExcelに変換",
    page_icon="📄",
    layout="wide"
)

# セッション状態の初期化
if 'users' not in st.session_state:
    st.session_state.users = {}

if 'user_state' not in st.session_state:
    st.session_state.user_state = {
        'is_logged_in': False,
        'is_premium': False,
        'email': None,
        'daily_conversions': 0,
        'last_conversion_date': None
    }

# ユーザー登録
def register_user(email, password):
    if email in st.session_state.users:
        return False, "このメールアドレスは既に登録されています"
    
    st.session_state.users[email] = {
        'password': password,
        'is_premium': False,
        'created_at': datetime.now()
    }
    return True, "登録が完了しました"

# ログイン認証
def login_user(email, password):
    if email not in st.session_state.users:
        return False, "メールアドレスが見つかりません"
    
    if st.session_state.users[email]['password'] != password:
        return False, "パスワードが正しくありません"
    
    return True, "ログインしました"

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

# 変換制限のチェック
def check_conversion_limit():
    current_date = datetime.now().date()
    
    if st.session_state.user_state['last_conversion_date'] != current_date:
        st.session_state.user_state['daily_conversions'] = 0
        st.session_state.user_state['last_conversion_date'] = current_date

    if st.session_state.user_state['is_premium']:
        return True
    elif st.session_state.user_state['is_logged_in']:
        return st.session_state.user_state['daily_conversions'] < 5
    else:
        return st.session_state.user_state['daily_conversions'] < 3

def analyze_document_structure(pdf_path):
    """帳票の構造を解析し、項目の位置を特定する"""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[0]
            
            # 罫線の検出
            edges = page.edges
            horizontals = [e for e in edges if e['orientation'] == 'horizontal']
            verticals = [e for e in edges if e['orientation'] == 'vertical']
            
            # テキストの抽出と位置情報の取得
            texts = page.extract_words(
                keep_blank_chars=True,
                x_tolerance=3,
                y_tolerance=3,
                extra_attrs=['size', 'font']
            )
            
            # 勘定科目のパターンを定義
            account_patterns = {
                '売上': r'売上|収入|営業収益',
                '経費': r'経費|販売費|一般管理費',
                '資産': r'資産|現金|預金|売掛金',
                '負債': r'負債|借入金|買掛金',
                '税金': r'税金|法人税|消費税'
            }
            
            # 項目の分類
            classified_items = {}
            for text in texts:
                for category, pattern in account_patterns.items():
                    if re.search(pattern, text['text']):
                        if category not in classified_items:
                            classified_items[category] = []
                        classified_items[category].append({
                            'text': text['text'],
                            'x0': text['x0'],
                            'y0': text['top'],
                            'x1': text['x1'],
                            'y1': text['bottom']
                        })
            
            return {
                'edges': {'horizontal': horizontals, 'vertical': verticals},
                'texts': texts,
                'classified_items': classified_items
            }
    except Exception as e:
        st.error(f"帳票構造の解析中にエラーが発生しました: {str(e)}")
        return None

def extract_numerical_values(text):
    """数値を抽出して整形する"""
    # カンマを除去して数値に変換
    numbers = re.findall(r'[\d,]+', text)
    return [int(num.replace(',', '')) for num in numbers if num]

def create_excel_output(document_structure, output_path):
    """抽出したデータをExcelファイルに出力"""
    try:
        # カテゴリごとのDataFrameを作成
        dfs = {}
        for category, items in document_structure['classified_items'].items():
            data = []
            for item in items:
                # 項目名の周辺で数値を探索
                nearby_texts = [t for t in document_structure['texts'] 
                              if abs(t['top'] - item['y0']) < 10]
                values = []
                for text in nearby_texts:
                    values.extend(extract_numerical_values(text['text']))
                
                data.append({
                    '項目': item['text'],
                    '金額': values[0] if values else 0
                })
            
            dfs[category] = pd.DataFrame(data)
        
        # Excelファイルに出力（シート分け）
        with pd.ExcelWriter(output_path, engine='openpyxdf') as writer:
            for category, df in dfs.items():
                df.to_excel(writer, sheet_name=category, index=False)
                
                # シートの書式設定
                workbook = writer.book
                worksheet = writer.sheets[category]
                
                # 列幅の自動調整
                for column in worksheet.columns:
                    max_length = 0
                    column = [cell for cell in column]
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = (max_length + 2)
                    worksheet.column_dimensions[column[0].column_letter].width = adjusted_width
                
                # 金額列の書式設定
                money_format = workbook.add_format({'num_format': '#,##0'})
                worksheet.set_column('B:B', None, money_format)
        
        return True
    except Exception as e:
        st.error(f"Excel出力中にエラーが発生しました: {str(e)}")
        return False

def process_pdf(uploaded_file):
    """PDFファイルを処理する"""
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_path = tmp_file.name
            
            # 帳票構造の解析
            document_structure = analyze_document_structure(tmp_path)
            if not document_structure:
                return None
            
            # Excelファイルの作成
            excel_path = tmp_path.replace('.pdf', '.xlsx')
            if create_excel_output(document_structure, excel_path):
                return excel_path
            
            return None
    except Exception as e:
        st.error(f"ファイル処理中にエラーが発生しました: {str(e)}")
        return None
    finally:
        if 'tmp_path' in locals():
            os.unlink(tmp_path)

# メインアプリケーション
def main():
    show_auth_ui()
    
    st.title("PDF to Excel 変換ツール")
    st.markdown("PDFファイルをExcel形式に変換できます。すべての処理はブラウザ内で行われます。")
    
    # 利用制限の表示
    if not st.session_state.user_state['is_premium']:
        remaining = 5 - st.session_state.user_state['daily_conversions'] if st.session_state.user_state['is_logged_in'] else 3 - st.session_state.user_state['daily_conversions']
        st.info(f"本日の残り変換回数: {remaining}回")
    
    # ファイルアップロード
    uploaded_file = st.file_uploader("PDFファイルを選択", type=['pdf'])

    if uploaded_file:
        if not check_conversion_limit():
            if st.session_state.user_state['is_logged_in']:
                st.error("本日の変換可能回数（5回）を超えました。プレミアムプランへのアップグレードをご検討ください。")
            else:
                st.error("本日の変換可能回数（3回）を超えました。アカウント登録で追加の2回が利用可能になります。")
            return

        with st.spinner('PDFを解析中...'):
            excel_path = process_pdf(uploaded_file)
            
            if excel_path:
                st.success("変換が完了しました！")
                
                # プレビューの表示
                excel_file = pd.ExcelFile(excel_path)
                for sheet_name in excel_file.sheet_names:
                    st.subheader(f"📊 {sheet_name}")
                    df = pd.read_excel(excel_file, sheet_name=sheet_name)
                    st.dataframe(df)
                
                # ダウンロードボタン
                with open(excel_path, 'rb') as f:
                    st.download_button(
                        label="📥 Excelファイルをダウンロード",
                        data=f,
                        file_name=f'converted_{uploaded_file.name}.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )
                
                os.remove(excel_path)
                
                if not st.session_state.user_state['is_premium']:
                    st.session_state.user_state['daily_conversions'] += 1

if __name__ == "__main__":
    main() 