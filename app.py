import streamlit as st
import pdfplumber
import pandas as pd
import tempfile
import os

# ページ設定
st.set_page_config(
    page_title="PDF to Excel コンバーター",
    page_icon="📄",
    layout="wide"
)

# タイトルと説明
st.title("PDF to Excel コンバーター")
st.write("PDFファイルをアップロードして、Excelに変換します。")

# ファイルアップローダー
uploaded_file = st.file_uploader("PDFファイルを選択してください", type=['pdf'])

if uploaded_file is not None:
    try:
        # 一時ファイルとして保存
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_path = tmp_file.name

        # PDFの処理
        tables = []
        with pdfplumber.open(tmp_path) as pdf:
            # プログレスバーの表示
            progress_bar = st.progress(0)
            total_pages = len(pdf.pages)
            
            for i, page in enumerate(pdf.pages):
                table = page.extract_table()
                if table:
                    tables.extend(table)
                # プログレスバーの更新
                progress_bar.progress((i + 1) / total_pages)

        if tables:
            # DataFrameに変換
            df = pd.DataFrame(tables[1:], columns=tables[0])
            
            # プレビューを表示
            st.write("データプレビュー:")
            st.dataframe(df)

            # Excelファイルとして出力
            excel_file = 'converted_data.xlsx'
            df.to_excel(excel_file, index=False)

            # ダウンロードボタン
            with open(excel_file, 'rb') as f:
                st.download_button(
                    label="📥 Excelファイルをダウンロード",
                    data=f,
                    file_name='converted_data.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
            
            # 一時ファイルの削除
            os.remove(excel_file)
        else:
            st.error("テーブルデータが見つかりませんでした。")

    except Exception as e:
        st.error(f"エラーが発生しました: {str(e)}")
    finally:
        # 一時ファイルの削除
        if os.path.exists(tmp_path):
            os.remove(tmp_path)

# フッター
st.markdown("---")
st.markdown("PDFからExcelへの変換ツール") 