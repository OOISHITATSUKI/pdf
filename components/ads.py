import streamlit as st
from utils.session import get_plan_type
from models.user import PlanType

def show_ads():
    """無料プランのユーザーに広告を表示するコンポーネント"""
    plan_type = get_plan_type()
    
    if not plan_type or plan_type == PlanType.FREE:
        st.markdown("""
        <script async src="https://pagead2.googlesyndication.com/pagead/js/adsbygoogle.js?client=ca-pub-9624397569723291" crossorigin="anonymous"></script>
        <meta name="google-adsense-account" content="ca-pub-9624397569723291">
        """, unsafe_allow_html=True) 