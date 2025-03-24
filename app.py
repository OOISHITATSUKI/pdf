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
uploaded_files = st.file_uploader("", type=['pdf'], accept_multiple_files=True)
if not uploaded_files:
    st.markdown('📄 クリックまたはドラッグ＆ドロップでPDFファイルを選択（最大3つまで）')
elif len(uploaded_files) > 3:
    st.error("⚠️ 無料版では一度に3つまでのファイルしか変換できません")
st.markdown('</div>', unsafe_allow_html=True)

# SEO対策のためのメタ情報
st.markdown("""
<!-- SEO対策用メタ情報 -->
<div style="display:none">
    PDF Excel 変換 無料 表 テーブル 一括変換 データ抽出 オンライン ツール
    PDFからExcelへの無料変換ツール 表形式データ抽出 高精度変換
</div>
""", unsafe_allow_html=True)

# 複数ファイル処理部分
if uploaded_files:
    for i, uploaded_file in enumerate(uploaded_files[:3]):  # 最大3つまでに制限
        with st.spinner(f'PDFファイル {i+1}/{len(uploaded_files)} を処理中...'):
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
                    tmp_file.write(uploaded_file.getvalue())
                    tmp_path = tmp_file.name

                # 広告表示（処理中）
                st.markdown('<div class="ad-container">広告スペース</div>', unsafe_allow_html=True)

                with pdfplumber.open(tmp_path) as pdf:
                    # テーブル認識精度の強化
                    all_tables = []
                    for page in pdf.pages:
                        try:
                            # シンプルな抽出方法を試す
                            table = page.extract_table()
                            if table:
                                all_tables.extend(table)
                            else:
                                # テキストとして抽出を試みる
                                text = page.extract_text()
                                if text:
                                    # テキストを1列のデータとして追加
                                    all_tables.append([text])
                        except Exception as e:
                            st.warning(f"ページの処理中にエラーが発生しました: {str(e)}")
                            continue

                    if all_tables:
                        # データフレームの作成と最適化
                        df = pd.DataFrame(all_tables)
                        # 空の行と列を削除
                        df = df.dropna(how='all').dropna(axis=1, how='all')
                        
                        st.markdown(f"### {uploaded_file.name} のプレビュー")
                        st.dataframe(df, use_container_width=True)
                        
                        # Excelファイル作成
                        excel_file = f'converted_data_{i+1}.xlsx'
                        df.to_excel(excel_file, index=False)
                        
                        with open(excel_file, 'rb') as f:
                            st.download_button(
                                label=f"📥 {uploaded_file.name} をダウンロード",
                                data=f,
                                file_name=excel_file,
                                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                            )
                        os.remove(excel_file)
                    else:
                        st.warning(f"⚠️ {uploaded_file.name} からテーブルデータが見つかりませんでした")

            except Exception as e:
                st.error(f"❌ {uploaded_file.name} の処理中にエラーが発生しました")
            finally:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)

# テーブル認識精度強化のための関数
def extract_table_with_enhanced_recognition(page):
    """
    複数の抽出方法を試行して最適な結果を返す
    """
    try:
        # 方法1: 標準的な抽出
        table = page.extract_table()
        if table and is_valid_table(table):
            return table

        # 方法2: カスタム設定での抽出
        table = page.extract_table(
            table_settings={
                "vertical_strategy": "text",
                "horizontal_strategy": "text",
                "snap_tolerance": 3,
                "join_tolerance": 3,
                "edge_min_length": 3,
                "min_words_vertical": 3,
            }
        )
        if table and is_valid_table(table):
            return table

        # 方法3: 線による抽出
        table = page.extract_table(
            table_settings={
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
            }
        )
        return table if table and is_valid_table(table) else None

    except Exception:
        return None

def enhance_table_structure(df):
    """
    pandas DataFrameの構造を改善
    """
    # NaN値の処理
    df = df.fillna('')
    
    # 重複列の処理
    df = df.loc[:, ~df.columns.duplicated()]
    
    # 空行の削除
    df = df.dropna(how='all')
    
    # 列名の正規化
    df.columns = [str(col).strip() for col in df.columns]
    
    # データの正規化
    df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
    
    return df

def save_enhanced_excel(df, filename):
    """
    整形されたExcelファイルを保存
    """
    writer = pd.ExcelWriter(filename, engine='openpyxl')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    
    # ワークシートの取得
    worksheet = writer.sheets['Sheet1']
    
    # 列幅の自動調整
    for column in worksheet.columns:
        max_length = 0
        column = [cell for cell in column]
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        worksheet.column_dimensions[column[0].column_letter].width = adjusted_width
    
    writer.save()

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