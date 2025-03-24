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
    .preview-box {
        background-color: #f8f9fa;
        padding: 20px;
        border-radius: 10px;
        min-height: 400px;
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

# メインコンテンツ
st.title("PDF to Excel 変換ツール")
st.markdown("PDFファイルをExcel形式に変換できます。すべての処理はブラウザ内で行われます。", unsafe_allow_html=True)

# 2列レイアウト
col1, col2 = st.columns([1, 1])

with col1:
    st.markdown("### 変換設定")
    
    # ファイルアップロードエリア
    st.markdown('<div class="upload-box">', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("", type=['pdf'])
    if not uploaded_file:
        st.markdown('📄<br>クリックまたはドラッグ＆ドロップでPDFファイルを選択', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)
    
    if uploaded_file:
        st.button("Excelに変換する")
    
    # サポート情報
    with st.expander("📌 サポート対象PDFについて"):
        st.markdown("""
        - テキストが含まれるPDF
        - 表形式のデータが含まれるPDF
        - シンプルな表構造のPDFが最適
        - セル結合が少ないPDFの方が良好な結果が得られます
        """)
    
    # 広告スペース1
    st.markdown('<div class="ad-space">広告スペース 1</div>', unsafe_allow_html=True)

with col2:
    st.markdown("### プレビュー")
    st.markdown('<div class="preview-box">', unsafe_allow_html=True)
    if not uploaded_file:
        st.markdown("PDFファイルをアップロードするとここにプレビューが表示されます")
    else:
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
                    st.dataframe(df)
                    
                    # Excelファイルとして出力
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
    st.markdown('</div>', unsafe_allow_html=True)

    # 広告スペース2
    st.markdown('<div class="ad-space">広告スペース 2</div>', unsafe_allow_html=True)

# フッター
st.markdown("""
<div class="footer">
    © 2025 PDF to Excel変換ツール - プライバシーを重視した無料のオンラインPDF変換サービス
</div>
""", unsafe_allow_html=True) 