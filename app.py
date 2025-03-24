import streamlit as st
import pdfplumber
import pandas as pd
import tempfile
import os

# ページ設定
st.set_page_config(
    page_title="PDF to Excel 変換ツール",
    page_icon="📄",
    layout="wide"
)

# カスタムCSS
st.markdown("""
<style>
    .main {
        background-color: #ffffff;
        padding: 0;
    }
    .block-container {
        padding-top: 1rem;
        padding-bottom: 0rem;
        padding-left: 2rem;
        padding-right: 2rem;
    }
    .preview-box {
        background-color: #ffffff;
        border: 1px solid #eee;
        border-radius: 5px;
        min-height: 400px;
        margin-top: 0;
        padding: 10px;
    }
    .stDataFrame {
        width: 100%;
    }
    div[data-testid="stDataFrame"] > div {
        width: 100%;
    }
    .dataframe {
        width: 100%;
    }
    thead tr th {
        background-color: #f8f9fa;
    }
    tbody tr:nth-child(even) {
        background-color: #f8f9fa;
    }
    .stButton>button {
        background-color: #8D9DA2;
        color: white;
        padding: 10px 24px;
        border-radius: 5px;
        border: none;
        font-size: 14px;
    }
    .stButton>button:hover {
        background-color: #7A8A8E;
    }
    .upload-box {
        border: 2px dashed #ccc;
        padding: 40px;
        text-align: center;
        border-radius: 10px;
        background-color: #f8f9fa;
        margin: 20px 0;
    }
    .upload-icon {
        font-size: 48px;
        color: #4F8BF9;
        margin-bottom: 10px;
    }
    .section-title {
        font-size: 16px;
        color: #333;
        margin-bottom: 20px;
    }
    .ad-space {
        background-color: #f8f9fa;
        padding: 20px;
        border-radius: 5px;
        text-align: center;
        margin: 20px 0;
        color: #666;
    }
    .footer {
        text-align: center;
        color: #666;
        padding: 20px;
        font-size: 12px;
    }
    .support-text {
        font-size: 14px;
        color: #666;
        margin-top: 10px;
    }
</style>
""", unsafe_allow_html=True)

# メインレイアウト
st.title("PDF to Excel 変換ツール")
st.markdown("PDFファイルをExcel形式に変換できます。すべての処理はブラウザ内で行われます。", unsafe_allow_html=True)

# ファイルアップロード部分
uploaded_file = st.file_uploader("PDFファイルを選択してください", type=['pdf'])

# プレビュー表示（上部に配置）
if uploaded_file is not None:
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_path = tmp_file.name

        with pdfplumber.open(tmp_path) as pdf:
            tables = []
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    tables.extend(table)

            if tables:
                df = pd.DataFrame(tables[1:], columns=tables[0])
                st.markdown("### プレビュー")
                st.dataframe(df, use_container_width=True)
                
                # Excelファイルの作成とダウンロードボタン
                excel_file = 'converted_data.xlsx'
                df.to_excel(excel_file, index=False)
                
                with open(excel_file, 'rb') as f:
                    st.download_button(
                        label="📥 Excelファイルをダウンロード",
                        data=f,
                        file_name='converted_data.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )
                os.remove(excel_file)
            else:
                st.warning("テーブルデータが見つかりませんでした。")
    except Exception as e:
        st.error("エラーが発生しました。PDFの形式を確認してください。")
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
else:
    st.markdown("PDFファイルをアップロードするとここにプレビューが表示されます")

# サポート情報（下部に配置）
with st.expander("📌 サポート対象PDFについて"):
    st.markdown("""
    ### 対応PDFの種類
    - ✅ 表形式のデータを含むPDF
    - ✅ 通常のテキストデータを含むPDF
    - ✅ 複合的なコンテンツを含むPDF
    
    ### 注意事項
    - ⚠️ スキャンされたPDFや画像化されたPDFは変換できない場合があります
    - ⚠️ パスワード保護されたPDFは処理できません
    """)

# 広告スペース（最下部に配置）
st.markdown('<div class="ad-space">広告スペース</div>', unsafe_allow_html=True)

# フッター
st.markdown("""
<div class="footer">
    © 2025 PDF to Excel変換ツール - プライバシーを重視した無料のオンラインPDF変換サービス
</div>
""", unsafe_allow_html=True) 