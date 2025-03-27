import streamlit as st
from typing import Optional

def init_session_state():
    """セッション状態を初期化"""
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False
    if "user_id" not in st.session_state:
        st.session_state.user_id = None
    if "user_email" not in st.session_state:
        st.session_state.user_email = None
    if "plan_type" not in st.session_state:
        st.session_state.plan_type = None

def set_user_session(user_id: int, email: str, plan_type: Optional[str] = None):
    """ユーザーセッションを設定"""
    st.session_state.logged_in = True
    st.session_state.user_id = user_id
    st.session_state.user_email = email
    if plan_type:
        st.session_state.plan_type = plan_type

def clear_user_session():
    """ユーザーセッションをクリア"""
    st.session_state.logged_in = False
    st.session_state.user_id = None
    st.session_state.user_email = None
    st.session_state.plan_type = None

def is_logged_in() -> bool:
    """ログイン状態を確認"""
    return st.session_state.get("logged_in", False)

def get_user_id() -> Optional[int]:
    """ユーザーIDを取得"""
    return st.session_state.get("user_id")

def get_user_email() -> Optional[str]:
    """ユーザーメールアドレスを取得"""
    return st.session_state.get("user_email")

def get_plan_type() -> Optional[str]:
    """プランタイプを取得"""
    return st.session_state.get("plan_type") 