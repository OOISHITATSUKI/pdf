import streamlit as st
import pdfplumber
import pandas as pd
import tempfile
import os
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

# PDFの処理
def process_pdf(uploaded_file):
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_path = tmp_file.name

        with pdfplumber.open(tmp_path) as pdf:
            all_tables = []
            for page in pdf.pages:
                tables = page.extract_tables()
                if tables:
                    all_tables.extend(tables)

        os.unlink(tmp_path)

        if all_tables:
            df = pd.DataFrame(all_tables[0])
            return df
        else:
            st.warning("PDFからテーブルを抽出できませんでした")
            return None

    except Exception as e:
        st.error(f"ファイルの処理中にエラーが発生しました: {str(e)}")
        return None

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

        with st.spinner('変換中...'):
            df = process_pdf(uploaded_file)
            
            if df is not None:
                st.success("変換が完了しました！")
                
                # プレビュー表示
                st.write("プレビュー:")
                st.dataframe(df)
                
                # Excelファイルの作成とダウンロード
                excel_file = f'converted_{uploaded_file.name}.xlsx'
                df.to_excel(excel_file, index=False)
                
                with open(excel_file, 'rb') as f:
                    st.download_button(
                        label=f"📥 Excelファイルをダウンロード",
                        data=f,
                        file_name=excel_file,
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )
                
                os.remove(excel_file)
                
                if not st.session_state.user_state['is_premium']:
                    st.session_state.user_state['daily_conversions'] += 1

if __name__ == "__main__":
    main() 