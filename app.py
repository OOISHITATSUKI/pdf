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
            
            # テキストとその詳細な属性を抽出（フォント関連の属性を修正）
            texts = page.extract_words(
                keep_blank_chars=True,
                x_tolerance=3,
                y_tolerance=3
            )
            
            # 詳細なテキスト情報の抽出
            chars = page.chars
            
            # テキスト情報にフォントサイズを追加
            for text in texts:
                # テキストの位置に基づいて対応する文字を検索
                matching_chars = [
                    char for char in chars
                    if char['x0'] >= text['x0'] and char['x1'] <= text['x1']
                    and char['top'] >= text['top'] and char['bottom'] <= text['bottom']
                ]
                
                # フォントサイズの平均を計算
                if matching_chars:
                    text['size'] = sum(char.get('size', 11) for char in matching_chars) / len(matching_chars)
                else:
                    text['size'] = 11  # デフォルトのフォントサイズ
            
            # 罫線情報の取得
            edges = page.edges
            horizontals = [e for e in edges if e['orientation'] == 'horizontal']
            verticals = [e for e in edges if e['orientation'] == 'vertical']
            
            # テーブル構造の検出
            tables = page.find_tables(
                table_settings={
                    "vertical_strategy": "text",
                    "horizontal_strategy": "text",
                    "intersection_y_tolerance": 10,
                    "intersection_x_tolerance": 10
                }
            )
            
            # セル結合の検出
            merged_cells = []
            if tables:
                for table in tables:
                    for cell in table.cells:
                        if cell.rowspan > 1 or cell.colspan > 1:
                            merged_cells.append({
                                'top': cell.bbox[1],
                                'bottom': cell.bbox[3],
                                'left': cell.bbox[0],
                                'right': cell.bbox[2],
                                'rowspan': cell.rowspan,
                                'colspan': cell.colspan
                            })
            
            return {
                'texts': texts,
                'merged_cells': merged_cells,
                'edges': {'horizontal': horizontals, 'vertical': verticals},
                'tables': tables
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
        
        # グリッドの作成（デフォルトのセルサイズを設定）
        default_row_height = 20
        default_col_width = 10
        
        # テキストの配置
        for text in layout_info['texts']:
            # 位置の計算
            col = int((text['x0']) // (default_col_width * 7)) + 1  # 7はおよそのピクセル/ポイント比
            row = int((text['top']) // default_row_height) + 1
            
            cell = ws.cell(row=row, column=col, value=text['text'])
            
            # スタイルの設定
            font_size = min(max(int(text.get('size', 11)), 8), 16)  # フォントサイズを8-16の範囲に制限
            cell.font = Font(size=font_size)
            
            # 数値の右寄せ
            if text['text'].replace(',', '').replace('¥', '').strip().isdigit():
                cell.alignment = Alignment(horizontal='right')
            else:
                cell.alignment = Alignment(vertical='center')
        
        # テーブルの罫線を設定
        if layout_info.get('tables'):
            for table in layout_info['tables']:
                for cell in table.cells:
                    row = int(cell.bbox[1] // default_row_height) + 1
                    col = int(cell.bbox[0] // (default_col_width * 7)) + 1
                    excel_cell = ws.cell(row=row, column=col)
                    excel_cell.border = Border(
                        left=Side(style='thin'),
                        right=Side(style='thin'),
                        top=Side(style='thin'),
                        bottom=Side(style='thin')
                    )
        
        # セル結合の適用
        for cell in layout_info['merged_cells']:
            start_row = int(cell['top'] // default_row_height) + 1
            end_row = int(cell['bottom'] // default_row_height) + 1
            start_col = int(cell['left'] // (default_col_width * 7)) + 1
            end_col = int(cell['right'] // (default_col_width * 7)) + 1
            
            if start_row < end_row or start_col < end_col:
                ws.merge_cells(
                    start_row=start_row,
                    start_column=start_col,
                    end_row=end_row,
                    end_column=end_col
                )
        
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
            adjusted_width = (max_length + 2) * 1.2  # 少し余裕を持たせる
            ws.column_dimensions[column].width = adjusted_width
        
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

# メインアプリケーション
def main():
    show_auth_ui()
    
    st.title("PDF to Excel 変換ツール")
    st.markdown("PDFファイルをExcel形式に変換できます。すべての処理はブラウザ内で行われます。")
    
    # 利用制限の表示
    if not st.session_state.user_state['is_premium']:
        remaining = 5 - st.session_state.user_state['daily_conversions'] if st.session_state.user_state['is_logged_in'] else 3 - st.session_state.user_state['daily_conversions']
        st.info(f"本日の残り変換回数: {remaining}回")
    
    # ファイルアップロード
    uploaded_file = st.file_uploader("PDFファイルを選択", type=['pdf'])

    if uploaded_file:
        if not check_conversion_limit():
            if st.session_state.user_state['is_logged_in']:
                st.error("本日の変換可能回数（5回）を超えました。プレミアムプランへのアップグレードをご検討ください。")
            else:
                st.error("本日の変換可能回数（3回）を超えました。アカウント登録で追加の2回が利用可能になります。")
            return

        with st.spinner('PDFを解析中...'):
            excel_path, exact_excel_path = process_pdf(uploaded_file)
            
            if excel_path and exact_excel_path:
                st.success("変換が完了しました！")
                
                # カテゴリ分類されたExcelのプレビュー
                st.subheader("📊 カテゴリ分類データ")
                excel_file = pd.ExcelFile(excel_path)
                for sheet_name in excel_file.sheet_names:
                    st.write(f"シート: {sheet_name}")
                    df = pd.read_excel(excel_file, sheet_name=sheet_name)
                    st.dataframe(df)
                
                # 完全レイアウトExcelのプレビュー
                st.subheader("📄 完全レイアウト")
                exact_df = pd.read_excel(exact_excel_path)
                st.dataframe(exact_df)
                
                # ダウンロードボタン
                col1, col2 = st.columns(2)
                with col1:
                    with open(excel_path, 'rb') as f:
                        st.download_button(
                            label="📥 カテゴリ分類データをダウンロード",
                            data=f,
                            file_name=f'categorized_{uploaded_file.name}.xlsx',
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                        )
                
                with col2:
                    with open(exact_excel_path, 'rb') as f:
                        st.download_button(
                            label="📥 完全レイアウトをダウンロード",
                            data=f,
                            file_name=f'exact_{uploaded_file.name}.xlsx',
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                        )
                
                os.remove(excel_path)
                os.remove(exact_excel_path)
                
                if not st.session_state.user_state['is_premium']:
                    st.session_state.user_state['daily_conversions'] += 1

if __name__ == "__main__":
    main() 