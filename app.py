import streamlit as st
import pdfplumber
import pandas as pd
import numpy as np
from PIL import Image
import tempfile
import os
import re
from datetime import datetime

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

def analyze_document_structure(pdf_path):
    """帳票の構造を解析し、項目の位置を特定する"""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[0]
            
            # テキストの抽出と位置情報の取得
            texts = page.extract_words(
                keep_blank_chars=True,
                x_tolerance=3,
                y_tolerance=3
            )
            
            # 勘定科目のパターンを定義
            account_patterns = {
                '売上': r'売上|収入|営業収益',
                '経費': r'経費|販売費|一般管理費',
                '資産': r'資産|現金|預金|売掛金',
                '負債': r'負債|借入金|買掛金',
                '税金': r'税金|法人税|消費税'
            }
            
            # 項目の分類
            classified_items = {}
            for text in texts:
                for category, pattern in account_patterns.items():
                    if re.search(pattern, text['text']):
                        if category not in classified_items:
                            classified_items[category] = []
                        classified_items[category].append({
                            'text': text['text'],
                            'x0': text['x0'],
                            'y0': text['top'],
                            'x1': text['x1'],
                            'y1': text['bottom']
                        })
            
            # 表の検出
            tables = page.extract_tables()
            
            return {
                'texts': texts,
                'classified_items': classified_items,
                'tables': tables
            }
    except Exception as e:
        st.error(f"帳票構造の解析中にエラーが発生しました: {str(e)}")
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

def create_excel_output(document_structure, output_path):
    """抽出したデータをExcelファイルに出力"""
    try:
        # カテゴリごとのDataFrameを作成
        dfs = {}
        
        # 分類された項目の処理
        for category, items in document_structure['classified_items'].items():
            data = []
            for item in items:
                # 項目名の周辺で数値を探索
                nearby_texts = [t for t in document_structure['texts'] 
                              if abs(t['top'] - item['y0']) < 10]
                values = []
                for text in nearby_texts:
                    values.extend(extract_numerical_values(text['text']))
                
                data.append({
                    '項目': item['text'],
                    '金額': values[0] if values else 0
                })
            
            if data:
                dfs[category] = pd.DataFrame(data)
        
        # テーブルデータの処理
        if document_structure['tables']:
            table_data = []
            for table in document_structure['tables']:
                if table:  # テーブルが空でない場合
                    df = pd.DataFrame(table[1:], columns=table[0] if table[0] else None)
                    table_data.append(df)
            
            if table_data:
                dfs['テーブルデータ'] = pd.concat(table_data, ignore_index=True)
        
        # Excelファイルに出力
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for category, df in dfs.items():
                df.to_excel(writer, sheet_name=category, index=False)
        
        return True
    except Exception as e:
        st.error(f"Excel出力中にエラーが発生しました: {str(e)}")
        return False

def extract_exact_layout(pdf_path):
    """PDFの完全なレイアウトを抽出してExcelに再現する"""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[0]
            
            # 全ての要素を抽出
            texts = page.extract_words(
                keep_blank_chars=True,
                x_tolerance=1,
                y_tolerance=1
            )
            
            # 罫線情報の取得と整理
            edges = page.edges
            horizontals = sorted([e for e in edges if e['orientation'] == 'horizontal'], key=lambda x: x['y0'])
            verticals = sorted([e for e in edges if e['orientation'] == 'vertical'], key=lambda x: x['x0'])
            
            # グリッドの作成
            grid = []
            for i in range(len(horizontals) - 1):
                row = []
                for j in range(len(verticals) - 1):
                    # セルの境界を定義
                    cell = {
                        'x0': verticals[j]['x0'],
                        'x1': verticals[j + 1]['x0'],
                        'y0': horizontals[i]['y0'],
                        'y1': horizontals[i + 1]['y0'],
                        'merged': False,
                        'text': ''
                    }
                    
                    # セル内のテキストを検索
                    cell_texts = [
                        t for t in texts
                        if t['x0'] >= cell['x0'] - 2 and t['x1'] <= cell['x1'] + 2
                        and t['top'] >= cell['y0'] - 2 and t['bottom'] <= cell['y1'] + 2
                    ]
                    
                    if cell_texts:
                        cell['text'] = ' '.join(t['text'] for t in cell_texts)
                    
                    row.append(cell)
                grid.append(row)
            
            # セル結合の検出
            merged_cells = []
            for i in range(len(grid)):
                for j in range(len(grid[i])):
                    if grid[i][j]['merged']:
                        continue
                    
                    # 横方向の結合を検出
                    merge_width = 1
                    while j + merge_width < len(grid[i]):
                        next_cell = grid[i][j + merge_width]
                        if next_cell['text'] == '' and not next_cell['merged']:
                            merge_width += 1
                        else:
                            break
                    
                    # 縦方向の結合を検出
                    merge_height = 1
                    while i + merge_height < len(grid):
                        next_row_cell = grid[i + merge_height][j]
                        if next_row_cell['text'] == '' and not next_row_cell['merged']:
                            merge_height += 1
                        else:
                            break
                    
                    # 結合セルとして記録
                    if merge_width > 1 or merge_height > 1:
                        merged_cell = {
                            'start_row': i,
                            'end_row': i + merge_height,
                            'start_col': j,
                            'end_col': j + merge_width,
                            'text': grid[i][j]['text']
                        }
                        merged_cells.append(merged_cell)
                        
                        # 結合されたセルをマーク
                        for mi in range(i, i + merge_height):
                            for mj in range(j, j + merge_width):
                                if mi < len(grid) and mj < len(grid[mi]):
                                    grid[mi][mj]['merged'] = True
            
            return {
                'grid': grid,
                'merged_cells': merged_cells,
                'edges': {'horizontal': horizontals, 'vertical': verticals}
            }
            
    except Exception as e:
        st.error(f"レイアウト抽出中にエラーが発生しました: {str(e)}")
        return None

def create_exact_excel_layout(layout_info, output_path):
    """抽出したレイアウト情報を元にExcelファイルを作成"""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
        from openpyxl.utils import get_column_letter
        
        wb = Workbook()
        ws = wb.active
        ws.title = "完全レイアウト"
        
        # 基本の罫線スタイル
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # グリッドデータの配置
        for i, row in enumerate(layout_info['grid']):
            for j, cell in enumerate(row):
                if not cell['merged']:
                    excel_cell = ws.cell(row=i+1, column=j+1, value=cell['text'])
                    
                    # 数値の判定と右寄せ
                    if cell['text'].replace(',', '').replace('¥', '').replace('(', '').replace(')', '').strip().isdigit():
                        excel_cell.alignment = Alignment(horizontal='right', vertical='center')
                    else:
                        excel_cell.alignment = Alignment(horizontal='left', vertical='center')
                    
                    # 罫線の設定
                    excel_cell.border = thin_border
        
        # セル結合の適用
        for merged_cell in layout_info['merged_cells']:
            try:
                ws.merge_cells(
                    start_row=merged_cell['start_row'] + 1,
                    start_column=merged_cell['start_col'] + 1,
                    end_row=merged_cell['end_row'],
                    end_column=merged_cell['end_col']
                )
                
                # 結合したセルのスタイル設定
                cell = ws.cell(
                    row=merged_cell['start_row'] + 1,
                    column=merged_cell['start_col'] + 1,
                    value=merged_cell['text']
                )
                cell.alignment = Alignment(horizontal='left', vertical='center')
                cell.border = thin_border
            except:
                continue
        
        # 列幅の自動調整
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

def process_pdf(uploaded_file):
    """PDFファイルを処理する"""
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_path = tmp_file.name
            
            # 既存の処理（カテゴリ分類）
            document_structure = analyze_document_structure(tmp_path)
            
            # 完全なレイアウト抽出
            layout_info = extract_exact_layout(tmp_path)
            
            # Excelファイルの作成
            excel_path = tmp_path.replace('.pdf', '.xlsx')
            
            if document_structure and layout_info:
                # 既存のシートを作成
                create_excel_output(document_structure, excel_path)
                
                # 完全レイアウトシートを追加
                create_exact_excel_layout(layout_info, excel_path.replace('.xlsx', '_exact.xlsx'))
                
                return excel_path, excel_path.replace('.xlsx', '_exact.xlsx')
            
            return None, None
            
    except Exception as e:
        st.error(f"ファイル処理中にエラーが発生しました: {str(e)}")
        return None, None
    finally:
        if 'tmp_path' in locals():
            os.unlink(tmp_path)

def process_multiple_pdfs(uploaded_files):
    """複数のPDFファイルを処理する"""
    try:
        # 一時ファイル用のディレクトリを作成
        with tempfile.TemporaryDirectory() as temp_dir:
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
            
            if all_results:
                # カテゴリ分類版Excelの作成
                categorized_path = os.path.join(temp_dir, 'categorized_results.xlsx')
                create_combined_excel(all_results, categorized_path)
                
                # 完全レイアウト版Excelの作成
                layout_path = os.path.join(temp_dir, 'layout_results.xlsx')
                create_combined_layout_excel(all_results, layout_path)
                
                return categorized_path, layout_path
            
            return None, None
            
    except Exception as e:
        st.error(f"ファイル処理中にエラーが発生しました: {str(e)}")
        return None, None

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
    
    uploaded_files = st.file_uploader(
        "PDFファイルを選択（複数可）", 
        type=['pdf'],
        accept_multiple_files=True
    )
    
    if uploaded_files:
        with st.spinner('PDFを解析中...'):
            categorized_path, layout_path = process_multiple_pdfs(uploaded_files)
            
            if categorized_path and layout_path:
                st.success("変換が完了しました！")
                
                # カテゴリ分類版のプレビューと出力
                st.subheader("📊 カテゴリ分類データ")
                excel_file = pd.ExcelFile(categorized_path)
                for sheet_name in excel_file.sheet_names:
                    st.write(f"シート: {sheet_name}")
                    df = pd.read_excel(excel_file, sheet_name=sheet_name)
                    st.dataframe(df)
                
                # 完全レイアウト版のプレビューと出力
                st.subheader("📄 完全レイアウト")
                layout_excel = pd.ExcelFile(layout_path)
                for sheet_name in layout_excel.sheet_names:
                    st.write(f"シート: {sheet_name}")
                    df = pd.read_excel(layout_excel, sheet_name=sheet_name)
                    st.dataframe(df)
                
                # ダウンロードボタン
                col1, col2 = st.columns(2)
                with col1:
                    with open(categorized_path, 'rb') as f:
                        st.download_button(
                            label="📥 カテゴリ分類データをダウンロード",
                            data=f,
                            file_name='categorized_results.xlsx',
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                        )
                
                with col2:
                    with open(layout_path, 'rb') as f:
                        st.download_button(
                            label="📥 完全レイアウトをダウンロード",
                            data=f,
                            file_name='layout_results.xlsx',
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                        )

if __name__ == "__main__":
    main() 