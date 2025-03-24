import streamlit as st
import pdfplumber
import pandas as pd
import tempfile
import os

# ページ設定とスタイル
st.set_page_config(
    page_title="PDF to Excel コンバーター",
    page_icon="📄",
    layout="wide"
)

# カスタムCSS
st.markdown("""
<style>
    .main {
        background-color: #f5f7f9;
    }
    .stButton>button {
        background-color: #4CAF50;
        color: white;
        padding: 10px 24px;
        border-radius: 8px;
        border: none;
    }
    .stButton>button:hover {
        background-color: #45a049;
    }
    .info-box {
        background-color: #e8f4f8;
        padding: 20px;
        border-radius: 10px;
        margin: 10px 0;
    }
    .ad-container {
        background-color: #ffffff;
        padding: 15px;
        border-radius: 8px;
        margin: 20px 0;
        text-align: center;
    }
</style>
""", unsafe_allow_html=True)

# ヘッダー
st.title("🔄 PDF to Excel コンバーター")
st.markdown("---")

# 広告スペース1
st.markdown('<div class="ad-container">広告スペース 1</div>', unsafe_allow_html=True)

# Google AdSenseのコード
ad_code = """
<script async src="https://pagead2.googlesyndication.com/pagead/js/adsbygoogle.js?client=YOUR-CLIENT-ID"
     crossorigin="anonymous"></script>
<!-- 広告ユニット -->
<ins class="adsbygoogle"
     style="display:block"
     data-ad-client="YOUR-CLIENT-ID"
     data-ad-slot="YOUR-AD-SLOT"
     data-ad-format="auto"
     data-full-width-responsive="true"></ins>
<script>
     (adsbygoogle = window.adsbygoogle || []).push({});
</script>
"""

# サポート情報
with st.expander("📌 サポート対象PDFについて"):
    st.markdown("""
    ### 対応PDFの種類
    - ✅ テキストが含まれるPDF
    - ✅ 表形式のデータが含まれるPDF
    - ⚠️ 画像化されたPDFは変換精度が低下する可能性があります
    - ⚠️ 複雑なレイアウトは正確に変換できない場合があります
    
    ### 変換のヒント
    - 📝 シンプルな表構造のPDFが最適です
    - 📝 セル結合が少ないPDFの方が良好な結果が得られます
    - ⚠️ 変換に失敗する場合は、PDFの品質や互換性を確認してください
    """)

# 抽出設定
with st.expander("⚙️ 詳細設定"):
    extraction_mode = st.radio(
        "抽出モード",
        ["テーブルモード", "テキストモード"],
        help="テーブルモード：表形式のデータを抽出\nテキストモード：すべてのテキストを抽出"
    )
    
    if extraction_mode == "テーブルモード":
        vertical_strategy = st.selectbox("縦方向の抽出方法", ["text", "lines", "explicit"])
        horizontal_strategy = st.selectbox("横方向の抽出方法", ["text", "lines", "explicit"])

# ファイルアップローダー
uploaded_file = st.file_uploader("PDFファイルを選択してください", type=['pdf'])

if uploaded_file is not None:
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_path = tmp_file.name

        # 広告スペース2
        st.markdown('<div class="ad-container">広告スペース 2</div>', unsafe_allow_html=True)

        with pdfplumber.open(tmp_path) as pdf:
            progress_bar = st.progress(0)
            total_pages = len(pdf.pages)
            
            if extraction_mode == "テーブルモード":
                tables = []
                for i, page in enumerate(pdf.pages):
                    table = page.extract_table(
                        vertical_strategy=vertical_strategy,
                        horizontal_strategy=horizontal_strategy
                    )
                    if table:
                        tables.extend(table)
                    progress_bar.progress((i + 1) / total_pages)

                if tables:
                    df = pd.DataFrame(tables[1:], columns=tables[0])
                else:
                    st.warning("テーブルが見つかりませんでした。テキストモードを試してください。")
                    df = None

            else:  # テキストモード
                texts = []
                for i, page in enumerate(pdf.pages):
                    text = page.extract_text()
                    if text:
                        texts.append([text])
                    progress_bar.progress((i + 1) / total_pages)
                
                if texts:
                    df = pd.DataFrame(texts, columns=['テキスト内容'])
                else:
                    st.warning("テキストが見つかりませんでした。")
                    df = None

            if df is not None:
                st.write("### プレビュー:")
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

    except Exception as e:
        st.error(f"エラーが発生しました: {str(e)}")
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)

# フッター
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666;'>
    <p>PDFからExcelへの変換ツール | © 2024</p>
</div>
""", unsafe_allow_html=True)

# 広告の表示
st.markdown(ad_code, unsafe_allow_html=True) 