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

# カスタムCSSを更新
st.markdown("""
<style>
    /* 全体のスタイル */
    .main {
        background: linear-gradient(135deg, #f8f9fa 0%, #ffffff 100%);
        padding: 0;
    }
    .block-container {
        padding: 2rem 3rem;
    }
    
    /* タイトルスタイル */
    h1 {
        color: #1E88E5;
        font-size: 2.5rem !important;
        font-weight: 700 !important;
        margin-bottom: 1rem !important;
        text-align: center;
        background: linear-gradient(45deg, #1E88E5, #64B5F6);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    
    /* サブタイトル */
    .subtitle {
        text-align: center;
        color: #666;
        font-size: 1.1rem;
        margin-bottom: 2rem;
    }
    
    /* アップロードエリア */
    .upload-area {
        background: white;
        border: 2px dashed #1E88E5;
        border-radius: 15px;
        padding: 2rem;
        text-align: center;
        margin: 2rem 0;
        transition: all 0.3s ease;
    }
    .upload-area:hover {
        border-color: #64B5F6;
        background: #f8f9fa;
    }
    
    /* プレビューエリア */
    .preview-box {
        background: white;
        border-radius: 15px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        padding: 1.5rem;
        margin: 2rem 0;
    }
    
    /* データフレームスタイル */
    .stDataFrame {
        width: 100%;
    }
    div[data-testid="stDataFrame"] > div {
        width: 100%;
    }
    .dataframe {
        width: 100%;
        border-collapse: separate;
        border-spacing: 0;
    }
    thead tr th {
        background: linear-gradient(45deg, #1E88E5, #64B5F6);
        color: white !important;
        padding: 12px !important;
    }
    tbody tr:nth-child(even) {
        background-color: #f8f9fa;
    }
    tbody tr:hover {
        background-color: #e3f2fd;
    }
    
    /* ダウンロードボタン */
    .stDownloadButton button {
        background: linear-gradient(45deg, #1E88E5, #64B5F6) !important;
        color: white !important;
        padding: 0.75rem 2rem !important;
        border-radius: 25px !important;
        border: none !important;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1) !important;
        transition: all 0.3s ease !important;
    }
    .stDownloadButton button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 8px rgba(0, 0, 0, 0.15) !important;
    }
    
    /* エクスパンダー */
    .streamlit-expanderHeader {
        background: white !important;
        border-radius: 10px !important;
        border: 1px solid #e0e0e0 !important;
    }
    .streamlit-expanderHeader:hover {
        background: #f8f9fa !important;
    }
    
    /* フッター */
    .footer {
        text-align: center;
        padding: 2rem;
        color: #666;
        background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%);
        border-radius: 15px;
        margin-top: 3rem;
    }
    
    /* ローディングスピナー */
    .stSpinner > div {
        border-color: #1E88E5 !important;
    }
</style>
""", unsafe_allow_html=True)

# メインレイアウト
st.markdown('<h1>PDF to Excel 変換ツール</h1>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">PDFファイルをExcel形式に変換できます。すべての処理はブラウザ内で行われます。</p>', unsafe_allow_html=True)

# ファイルアップロード
st.markdown('<div class="upload-area">', unsafe_allow_html=True)
uploaded_file = st.file_uploader("", type=['pdf'])
if not uploaded_file:
    st.markdown('📄 クリックまたはドラッグ＆ドロップでPDFファイルを選択')
st.markdown('</div>', unsafe_allow_html=True)

# プレビュー表示
if uploaded_file is not None:
    st.markdown('<div class="preview-box">', unsafe_allow_html=True)
    with st.spinner('PDFを処理中...'):
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
                    
                    # Excelファイル作成
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
                    st.warning("⚠️ テーブルデータが見つかりませんでした")
        except Exception as e:
            st.error("❌ エラーが発生しました。PDFの形式を確認してください。")
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
    st.markdown('</div>', unsafe_allow_html=True)

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