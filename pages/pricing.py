import streamlit as st
from typing import Dict, Any
import os
from dotenv import load_dotenv

# 環境変数の読み込み
load_dotenv()

# プラン情報の定義
PLAN_INFO = {
    "free": {
        "name": "無料プラン",
        "price": "¥0",
        "color": "gray",
        "features": {
            "PDF変換回数": "1日5回（登録済）",
            "PDFページ数上限": "1ページ",
            "Excel出力品質": "通常品質（簡易整形）",
            "Vision API使用": "×",
            "会計データ構造化": "×",
            "分析機能": "×",
            "変換ファイル保存": "×",
            "広告": "表示あり"
        },
        "stripe_link": None
    },
    "pro": {
        "name": "プロプラン",
        "price": "$5/月",
        "color": "blue",
        "features": {
            "PDF変換回数": "月500回まで",
            "PDFページ数上限": "3ページ/PDF",
            "Excel出力品質": "高品質（セル結合・書式保持）",
            "Vision API使用": "○（Google Vision OCR）",
            "会計データ構造化": "あり（定義済テンプレ対応）",
            "分析機能": "売上比率などの計算含む",
            "変換ファイル保存": "保存なし",
            "広告": "非表示"
        },
        "stripe_link": "https://buy.stripe.com/aEUeYEbKhc63gUwfYZ"
    },
    "premium": {
        "name": "プレミアムプラン",
        "price": "$20/月",
        "color": "orange",
        "features": {
            "PDF変換回数": "月2000回まで",
            "PDFページ数上限": "6ページ/PDF",
            "Excel出力品質": "最上級（整形＋分析式入り）",
            "Vision API使用": "○（Google Vision OCR）",
            "会計データ構造化": "あり＋複数シート分割可能",
            "分析機能": "グラフ＋比率分析付き",
            "変換ファイル保存": "30日保存あり",
            "広告": "非表示"
        },
        "stripe_link": "https://buy.stripe.com/00g6s8aGdeeb33GaEG"
    }
}

def check_authentication() -> bool:
    """ユーザーの認証状態を確認"""
    return "user_id" in st.session_state

def render_plan_card(plan_data: Dict[str, Any]) -> None:
    """プランカードをレンダリング"""
    with st.container():
        st.markdown(f"### {plan_data['name']}")
        st.markdown(f"#### {plan_data['price']}")
        
        # 機能一覧
        for feature, value in plan_data["features"].items():
            st.markdown(f"- **{feature}:** {value}")
        
        # 購入ボタン（無料プラン以外）
        if plan_data["stripe_link"]:
            if check_authentication():
                st.link_button("プランを選択", plan_data["stripe_link"])
            else:
                st.warning("プランを選択するにはログインが必要です")

def main():
    st.title("料金プラン")
    st.markdown("### 最適なプランをお選びください")

    # プラン比較表示
    cols = st.columns(3)
    for col, (plan_type, plan_data) in zip(cols, PLAN_INFO.items()):
        with col:
            render_plan_card(plan_data)

    # 注意事項
    st.markdown("---")
    st.markdown("""
    #### 注意事項
    - 料金は税抜き価格です
    - プランは毎月自動更新されます
    - キャンセルは次回更新日の前日まで可能です
    - 無料プランへのダウングレードはいつでも可能です
    """)

if __name__ == "__main__":
    main() 