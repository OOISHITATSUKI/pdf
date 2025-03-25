import streamlit as st
import pdfplumber
import pandas as pd
import numpy as np
from PIL import Image
import tempfile
import os
import re
from datetime import datetime
from openpyxl.utils import get_column_letter

# ページ設定
st.set_page_config(
    page_title="PDF to Excel 変換ツール｜無料でPDFの表をExcelに変換",
    page_icon="📄",
    layout="wide"
)

# セッション状態の初期化
if 'users' not in st.session_state:
    st.session_state.users = {}

if 'user_state' not in st.session_state:
    st.session_state.user_state = {
        'is_logged_in': False,
        'is_premium': False,
        'email': None,
        'daily_conversions': 0,
        'last_conversion_date': None
    }

# ユーザー登録
def register_user(email, password):
    if email in st.session_state.users:
        return False, "このメールアドレスは既に登録されています"
    
    st.session_state.users[email] = {
        'password': password,
        'is_premium': False,
        'created_at': datetime.now()
    }
    return True, "登録が完了しました"

# ログイン認証
def login_user(email, password):
    if email not in st.session_state.users:
        return False, "メールアドレスが見つかりません"
    
    if st.session_state.users[email]['password'] != password:
        return False, "パスワードが正しくありません"
    
    return True, "ログインしました"

# 認証UI
def show_auth_ui():
    st.sidebar.markdown("### アカウント管理")
    
    if not st.session_state.user_state['is_logged_in']:
        tab1, tab2 = st.sidebar.tabs(["ログイン", "新規登録"])
        
        with tab1:
            with st.form("login_form"):
                login_email = st.text_input("メールアドレス", key="login_email")
                login_password = st.text_input("パスワード", type="password", key="login_password")
                login_submit = st.form_submit_button("ログイン")
                
                if login_submit:
                    success, message = login_user(login_email, login_password)
                    if success:
                        st.session_state.user_state['is_logged_in'] = True
                        st.session_state.user_state['email'] = login_email
                        st.session_state.user_state['is_premium'] = st.session_state.users[login_email]['is_premium']
                        st.success(message)
                        st.rerun()
                    else:
                        st.error(message)
        
        with tab2:
            with st.form("register_form"):
                reg_email = st.text_input("メールアドレス", key="reg_email")
                reg_password = st.text_input("パスワード", type="password", key="reg_password")
                reg_password_confirm = st.text_input("パスワード（確認）", type="password")
                register_submit = st.form_submit_button("新規登録")
                
                if register_submit:
                    if reg_password != reg_password_confirm:
                        st.error("パスワードが一致しません")
                    else:
                        success, message = register_user(reg_email, reg_password)
                        if success:
                            st.success(message)
                            st.session_state.user_state['is_logged_in'] = True
                            st.session_state.user_state['email'] = reg_email
                            st.rerun()
                        else:
                            st.error(message)
    
    else:
        st.sidebar.markdown(f"### ようこそ！")
        st.sidebar.markdown(f"ログイン中: {st.session_state.user_state['email']}")
        
        if not st.session_state.user_state['is_premium']:
            st.sidebar.markdown("### 🌟 プレミアムにアップグレード")
            if st.sidebar.button("プレミアム会員に登録"):
                st.sidebar.info("準備中です...")
        
        if st.sidebar.button("ログアウト"):
            st.session_state.user_state = {
                'is_logged_in': False,
                'is_premium': False,
                'email': None,
                'daily_conversions': 0,
                'last_conversion_date': None
            }
            st.rerun()

# 変換制限のチェック
def check_conversion_limit():
    current_date = datetime.now().date()
    
    if st.session_state.user_state['last_conversion_date'] != current_date:
        st.session_state.user_state['daily_conversions'] = 0
        st.session_state.user_state['last_conversion_date'] = current_date

    if st.session_state.user_state['is_premium']:
        return True
    elif st.session_state.user_state['is_logged_in']:
        return st.session_state.user_state['daily_conversions'] < 5
    else:
        return st.session_state.user_state['daily_conversions'] < 3

def extract_text_with_settings(page):
    """より正確なテキスト抽出のための設定"""
    return page.extract_text(
        x_tolerance=3,  # 文字間の水平方向の許容値
        y_tolerance=3,  # 文字間の垂直方向の許容値
        layout=True,    # レイアウトを考慮
        keep_blank_chars=False,  # 空白文字を除去
        use_text_flow=True,  # テキストの流れを考慮
        horizontal_ltr=True,  # 左から右への読み取り
        vertical_ttb=True,    # 上から下への読み取り
        extra_attrs=['fontname', 'size']  # フォント情報も取得
    )

def analyze_document_structure(pdf_path):
    """PDFの構造を解析する"""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[0]
            
            # テーブルの検出を試みる
            tables = page.extract_tables(
                table_settings={
                    "vertical_strategy": "text",
                    "horizontal_strategy": "text",
                    "intersection_y_tolerance": 10,
                    "intersection_x_tolerance": 10,
                    "snap_y_tolerance": 3,
                    "snap_x_tolerance": 3,
                    "join_y_tolerance": 3,
                    "join_x_tolerance": 3,
                    "edge_min_length": 3,
                    "min_words_vertical": 3,
                    "min_words_horizontal": 1
                }
            )
            
            if tables:
                # テーブルが見つかった場合の処理
                items = []
                for table in tables:
                    for row in table:
                        if any(row):  # 空でない行のみ処理
                            cleaned_row = [
                                str(cell).strip() if cell is not None else ""
                                for cell in row
                            ]
                            if any(cleaned_row):  # 空でない行のみ追加
                                items.append({
                                    'text': ' '.join(cleaned_row),
                                    'type': 'table_row'
                                })
            else:
                # テーブルが見つからない場合はテキストとして抽出
                text = extract_text_with_settings(page)
                items = [
                    {'text': line.strip(), 'type': 'text'}
                    for line in text.split('\n')
                    if line.strip()
                ]
            
            return {'items': items}
            
    except Exception as e:
        st.error(f"PDF解析中にエラーが発生しました: {str(e)}")
        return None

def extract_numerical_values(text):
    """数値を抽出して整形する"""
    # カンマを除去して数値に変換
    numbers = re.findall(r'[\d,]+', text)
    cleaned_numbers = []
    for num in numbers:
        try:
            cleaned_numbers.append(int(num.replace(',', '')))
        except ValueError:
            continue
    return cleaned_numbers

def create_excel_output(items, output_path):
    """抽出したデータをExcelに出力"""
    try:
        # DataFrameの作成
        df = pd.DataFrame([{'内容': item['text']} for item in items])
        
        # Excelファイルとして保存
        df.to_excel(output_path, index=False, engine='openpyxl')
        return True
    except Exception as e:
        st.error(f"Excel作成中にエラーが発生しました: {str(e)}")
        return False

def extract_exact_layout(pdf_path):
    """PDFの完全なレイアウトを抽出する"""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[0]
            
            # テキストの抽出（より詳細な設定で）
            texts = page.extract_words(
                keep_blank_chars=False,
                x_tolerance=1,
                y_tolerance=1,
                extra_attrs=['fontname', 'size'],
                use_text_flow=True
            )
            
            # 罫線情報の取得
            edges = page.edges
            horizontals = sorted([e for e in edges if e['orientation'] == 'horizontal'], key=lambda x: x['y0'])
            verticals = sorted([e for e in edges if e['orientation'] == 'vertical'], key=lambda x: x['x0'])
            
            # テーブルの抽出（より詳細な設定で）
            tables = page.extract_tables(
                table_settings={
                    "vertical_strategy": "text",
                    "horizontal_strategy": "text",
                    "intersection_y_tolerance": 10,
                    "intersection_x_tolerance": 10,
                    "snap_y_tolerance": 3,
                    "snap_x_tolerance": 3,
                    "join_y_tolerance": 3,
                    "join_x_tolerance": 3,
                    "edge_min_length": 3,
                    "min_words_vertical": 3,
                    "min_words_horizontal": 1
                }
            )
            
            # テキストの前処理
            processed_texts = []
            for text in texts:
                # cidの除去
                cleaned_text = re.sub(r'\(cid:\d+\)', '', text['text'])
                if cleaned_text.strip():
                    text['text'] = cleaned_text.strip()
                    processed_texts.append(text)
            
            return {
                'texts': processed_texts,
                'edges': {'horizontal': horizontals, 'vertical': verticals},
                'tables': tables
            }
            
    except Exception as e:
        st.error(f"レイアウト抽出中にエラーが発生しました: {str(e)}")
        return None

def create_layout_excel(layout_info, output_path):
    """レイアウト情報をExcelに出力"""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, Alignment, Border, Side
        
        wb = Workbook()
        ws = wb.active
        ws.title = "完全レイアウト"
        
        # 罫線スタイル
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # テキストの配置
        for text in layout_info['texts']:
            # 座標を行と列に変換
            row = int(text['top'] // 20) + 1  # 20ピクセルを1行とする
            col = int(text['x0'] // 50) + 1   # 50ピクセルを1列とする
            
            cell = ws.cell(row=row, column=col, value=text['text'])
            
            # スタイルの適用
            cell.border = thin_border
            
            # 数値の場合は右寄せ
            if text['text'].replace(',', '').replace('¥', '').replace('(', '').replace(')', '').strip().isdigit():
                cell.alignment = Alignment(horizontal='right', vertical='center')
            else:
                cell.alignment = Alignment(horizontal='left', vertical='center')
        
        # 罫線の配置
        if layout_info['edges']:
            # 水平線
            for h_line in layout_info['edges']['horizontal']:
                row = int(h_line['y0'] // 20) + 1
                start_col = int(h_line['x0'] // 50) + 1
                end_col = int(h_line['x1'] // 50) + 1
                
                for col in range(start_col, end_col + 1):
                    cell = ws.cell(row=row, column=col)
                    if not cell.value:
                        cell.value = ''
                    cell.border = thin_border
            
            # 垂直線
            for v_line in layout_info['edges']['vertical']:
                col = int(v_line['x0'] // 50) + 1
                start_row = int(v_line['y0'] // 20) + 1
                end_row = int(v_line['y1'] // 20) + 1
                
                for row in range(start_row, end_row + 1):
                    cell = ws.cell(row=row, column=col)
                    if not cell.value:
                        cell.value = ''
                    cell.border = thin_border
        
        # 列幅の調整
        for col in ws.columns:
            max_length = 0
            column = get_column_letter(col[0].column)
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2) * 1.2
            ws.column_dimensions[column].width = adjusted_width
        
        # 行の高さを統一
        for row in ws.rows:
            ws.row_dimensions[row[0].row].height = 20
        
        # ファイルを保存
        wb.save(output_path)
        return True
        
    except Exception as e:
        st.error(f"レイアウトExcel作成中にエラーが発生しました: {str(e)}")
        return False

def process_pdf(uploaded_file):
    """PDFファイルを処理する"""
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_pdf:
            tmp_pdf.write(uploaded_file.getvalue())
            pdf_path = tmp_pdf.name
            
            # 構造解析
            doc_structure = analyze_document_structure(pdf_path)
            # レイアウト解析
            layout_info = extract_exact_layout(pdf_path)
            
            if doc_structure or layout_info:
                # 通常版Excel
                normal_excel = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
                normal_path = normal_excel.name
                if doc_structure:
                    create_excel_output(doc_structure['items'], normal_path)
                
                # レイアウト版Excel
                layout_excel = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
                layout_path = layout_excel.name
                if layout_info:
                    create_layout_excel(layout_info, layout_path)
                
                return normal_path, layout_path
            
            return None, None
            
    except Exception as e:
        st.error(f"ファイル処理中にエラーが発生しました: {str(e)}")
        return None, None
    finally:
        if 'pdf_path' in locals():
            try:
                os.unlink(pdf_path)
            except:
                pass

def process_multiple_pdfs(uploaded_files):
    """複数のPDFファイルを処理する"""
    temp_dir = tempfile.mkdtemp()  # 一時ディレクトリを作成
    try:
        all_results = []
        
        for uploaded_file in uploaded_files:
            # 一時PDFファイルを作成
            pdf_path = os.path.join(temp_dir, uploaded_file.name)
            with open(pdf_path, 'wb') as f:
                f.write(uploaded_file.getvalue())
            
            # PDFの処理
            document_structure = analyze_document_structure(pdf_path)
            layout_info = extract_exact_layout(pdf_path)
            
            if document_structure and layout_info:
                result = {
                    'filename': uploaded_file.name,
                    'document_structure': document_structure,
                    'layout_info': layout_info
                }
                all_results.append(result)
            
            # 一時PDFファイルを削除
            os.remove(pdf_path)
        
        if all_results:
            # カテゴリ分類版Excelの作成
            categorized_path = os.path.join(temp_dir, 'categorized_results.xlsx')
            create_combined_excel(all_results, categorized_path)
            
            # 完全レイアウト版Excelの作成
            layout_path = os.path.join(temp_dir, 'layout_results.xlsx')
            create_combined_layout_excel(all_results, layout_path)
            
            # Excelファイルの内容を読み込む
            with open(categorized_path, 'rb') as f:
                categorized_data = f.read()
            with open(layout_path, 'rb') as f:
                layout_data = f.read()
            
            return categorized_data, layout_data
        
        return None, None
        
    except Exception as e:
        st.error(f"ファイル処理中にエラーが発生しました: {str(e)}")
        return None, None
    finally:
        # 一時ディレクトリとファイルの削除
        try:
            import shutil
            shutil.rmtree(temp_dir)
        except:
            pass

def create_combined_excel(results, output_path):
    """複数のPDFの結果を1つのExcelファイルにまとめる"""
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for idx, result in enumerate(results):
                filename = result['filename']
                doc_structure = result['document_structure']
                
                # カテゴリごとのDataFrameを作成
                for category, items in doc_structure['classified_items'].items():
                    sheet_name = f"{filename}_{category}"[:31]  # Excelのシート名制限
                    
                    df = pd.DataFrame(items)
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        return True
    except Exception as e:
        st.error(f"Excel作成中にエラーが発生しました: {str(e)}")
        return False

def create_combined_layout_excel(results, output_path):
    """複数のPDFの完全レイアウトを1つのExcelファイルにまとめる"""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, Alignment, Border, Side
        
        wb = Workbook()
        
        for idx, result in enumerate(results):
            filename = result['filename']
            layout_info = result['layout_info']
            
            # 各PDFに対して新しいシートを作成
            if idx == 0:
                ws = wb.active
                ws.title = f"Layout_{filename}"[:31]
            else:
                ws = wb.create_sheet(f"Layout_{filename}"[:31])
            
            # グリッドデータの配置
            for i, row in enumerate(layout_info['grid']):
                for j, cell in enumerate(row):
                    if not cell['merged']:
                        excel_cell = ws.cell(row=i+1, column=j+1, value=cell['text'])
                        
                        # スタイルの設定
                        if cell['text'].replace(',', '').replace('¥', '').replace('(', '').replace(')', '').strip().isdigit():
                            excel_cell.alignment = Alignment(horizontal='right', vertical='center')
                        else:
                            excel_cell.alignment = Alignment(horizontal='left', vertical='center')
                        
                        # 罫線の設定
                        excel_cell.border = Border(
                            left=Side(style='thin'),
                            right=Side(style='thin'),
                            top=Side(style='thin'),
                            bottom=Side(style='thin')
                        )
            
            # セル結合の適用
            for merged_cell in layout_info['merged_cells']:
                try:
                    ws.merge_cells(
                        start_row=merged_cell['start_row'] + 1,
                        start_column=merged_cell['start_col'] + 1,
                        end_row=merged_cell['end_row'],
                        end_column=merged_cell['end_col']
                    )
                    
                    cell = ws.cell(
                        row=merged_cell['start_row'] + 1,
                        column=merged_cell['start_col'] + 1,
                        value=merged_cell['text']
                    )
                    cell.alignment = Alignment(horizontal='left', vertical='center')
                except:
                    continue
            
            # 列幅の調整
            for col in ws.columns:
                max_length = 0
                column = get_column_letter(col[0].column)
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 2) * 1.2
                ws.column_dimensions[column].width = adjusted_width
            
            # 行の高さを統一
            for row in ws.rows:
                ws.row_dimensions[row[0].row].height = 20
        
        wb.save(output_path)
        return True
        
    except Exception as e:
        st.error(f"Excel作成中にエラーが発生しました: {str(e)}")
        return False

# メインアプリケーション
def main():
    show_auth_ui()
    
    st.title("PDF to Excel 変換ツール")
    st.markdown("PDFファイルをExcel形式に変換できます。")
    
    uploaded_file = st.file_uploader("PDFファイルを選択", type=['pdf'])
    
    if uploaded_file:
        with st.spinner('PDFを解析中...'):
            normal_path, layout_path = process_pdf(uploaded_file)
            
            if normal_path or layout_path:
                st.success("変換が完了しました！")
                
                # 通常版の表示
                if normal_path and os.path.exists(normal_path):
                    st.subheader("📊 通常データ")
                    try:
                        df = pd.read_excel(normal_path)
                        st.dataframe(df)
                    except Exception as e:
                        st.error(f"通常データの表示中にエラーが発生しました: {str(e)}")
                
                # レイアウト版の表示
                if layout_path and os.path.exists(layout_path):
                    st.subheader("📄 完全レイアウト")
                    try:
                        df = pd.read_excel(layout_path)
                        st.dataframe(df)
                    except Exception as e:
                        st.error(f"レイアウトデータの表示中にエラーが発生しました: {str(e)}")
                
                # ダウンロードボタン
                col1, col2 = st.columns(2)
                
                with col1:
                    if normal_path and os.path.exists(normal_path):
                        with open(normal_path, 'rb') as f:
                            normal_data = f.read()
                            st.download_button(
                                label="📥 通常データをダウンロード",
                                data=normal_data,
                                file_name=f'normal_{uploaded_file.name}.xlsx',
                                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                            )
                
                with col2:
                    if layout_path and os.path.exists(layout_path):
                        with open(layout_path, 'rb') as f:
                            layout_data = f.read()
                            st.download_button(
                                label="📥 完全レイアウトをダウンロード",
                                data=layout_data,
                                file_name=f'layout_{uploaded_file.name}.xlsx',
                                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                            )
                
                # 一時ファイルの削除
                try:
                    if normal_path and os.path.exists(normal_path):
                        os.unlink(normal_path)
                    if layout_path and os.path.exists(layout_path):
                        os.unlink(layout_path)
                except:
                    pass

if __name__ == "__main__":
    main() 