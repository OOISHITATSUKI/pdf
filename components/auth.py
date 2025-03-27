import streamlit as st
from models.user import User, PlanType
from database import SessionLocal
from passlib.context import CryptContext
import bcrypt
from utils.session import set_user_session

pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")

def show_auth_forms():
    """認証フォームを表示するコンポーネント"""
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        tab1, tab2 = st.tabs(["ログイン", "新規登録"])
        
        with tab1:
            with st.form("login_form"):
                st.subheader("ログイン")
                email = st.text_input("メールアドレス")
                password = st.text_input("パスワード", type="password")
                submit = st.form_submit_button("ログイン")
                
                if submit:
                    try:
                        db = SessionLocal()
                        user = db.query(User).filter(User.email == email).first()
                        if user and pwd_context.verify(password, user.hashed_password):
                            set_user_session(user.id, user.email, user.plan_type)
                            st.success("ログインしました！")
                            st.rerun()
                        else:
                            st.error("メールアドレスまたはパスワードが正しくありません。")
                    finally:
                        db.close()
        
        with tab2:
            with st.form("register_form"):
                st.subheader("新規登録")
                username = st.text_input("ユーザー名")
                email = st.text_input("メールアドレス")
                password = st.text_input("パスワード", type="password")
                password_confirm = st.text_input("パスワード（確認）", type="password")
                
                # 利用規約とプライバシーポリシーのリンク
                st.markdown("""
                <div style="text-align: center; margin: 10px 0;">
                    <a href="/docs/terms" target="_blank">利用規約</a> | 
                    <a href="/docs/security" target="_blank">プライバシーポリシー</a> | 
                    <a href="/contact" target="_blank">お問い合わせ</a>
                </div>
                """, unsafe_allow_html=True)
                
                submit = st.form_submit_button("登録")
                
                if submit:
                    if not username or not email or not password:
                        st.error("全ての項目を入力してください。")
                        return
                    
                    if password != password_confirm:
                        st.error("パスワードが一致しません。")
                        return
                    
                    try:
                        db = SessionLocal()
                        # ユーザー名とメールアドレスの重複チェック
                        if db.query(User).filter(User.username == username).first():
                            st.error("このユーザー名は既に使用されています。")
                            return
                        if db.query(User).filter(User.email == email).first():
                            st.error("このメールアドレスは既に登録されています。")
                            return
                        
                        # 新規ユーザーの作成
                        hashed_password = pwd_context.hash(password)
                        new_user = User(
                            username=username,
                            email=email,
                            hashed_password=hashed_password,
                            plan_type=PlanType.FREE
                        )
                        db.add(new_user)
                        db.commit()
                        db.refresh(new_user)
                        
                        # セッションの設定
                        set_user_session(new_user.id, new_user.email, new_user.plan_type)
                        st.success("アカウントが作成されました！")
                        
                        # プラン選択ページへリダイレクト
                        st.markdown('<meta http-equiv="refresh" content="2; url=/pricing">', unsafe_allow_html=True)
                    except Exception as e:
                        st.error(f"エラーが発生しました: {str(e)}")
                    finally:
                        db.close() 