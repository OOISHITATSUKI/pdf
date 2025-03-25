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
import uuid
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import Column, String, DateTime, Enum, JSON, ForeignKey, Text
from sqlalchemy.sql import func
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

# ページ設定
st.set_page_config(
    page_title="PDF to Excel 変換ツール｜無料でPDFの表をExcelに変換",
    page_icon="📄",
    layout="wide"
)

# セッション状態の初期化
if 'users' not in st.session_state:
    st.session_state.users = {}

def initialize_session_state():
    """セッション状態の初期化とローカルストレージとの同期"""
    if 'user_state' not in st.session_state:
        st.session_state.user_state = {
            'is_logged_in': False,
            'is_premium': False,
            'email': None,
            'daily_conversions': 0,
            'last_conversion_date': None,
            'device_id': None  # デバイス識別用
        }
    
    # ローカルストレージからの読み込み用JavaScript
    st.markdown("""
        <script>
            const deviceId = localStorage.getItem('deviceId') || Date.now().toString();
            localStorage.setItem('deviceId', deviceId);
            
            const conversions = localStorage.getItem('dailyConversions') || '0';
            const lastDate = localStorage.getItem('lastConversionDate');
            
            window.parent.postMessage({
                type: 'getLocalStorage',
                deviceId: deviceId,
                conversions: conversions,
                lastDate: lastDate
            }, '*');
        </script>
    """, unsafe_allow_html=True)

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

def is_tax_return_pdf(pdf_path):
    """確定申告書かどうかを判定"""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[0]
            text = page.extract_text()
            # 確定申告書に特有の文字列をチェック
            tax_keywords = ['確定申告書', '所得税', '法人税', '消費税', '源泉所得税']
            return any(keyword in text for keyword in tax_keywords)
    except:
        return False

def process_tax_return_pdf(page):
    """確定申告書専用の処理"""
    try:
        # テキストとレイアウト情報を抽出
        text = page.extract_text()
        words = page.extract_words(
            keep_blank_chars=False,
            x_tolerance=1,
            y_tolerance=1,
            extra_attrs=['fontname', 'size']
        )

        # テキストブロックを構造化
        blocks = []
        for word in words:
            if word['text'].strip():
                # CIDフォントの処理
                cleaned_text = re.sub(r'\(cid:\d+\)', '', word['text'])
                
                # 数値の処理
                numbers = re.findall(r'[\d,]+', cleaned_text)
                for num in numbers:
                    try:
                        value = int(num.replace(',', ''))
                        cleaned_text = cleaned_text.replace(num, f'{value:,}')
                    except ValueError:
                        continue
                
                # 位置情報をタプルとして保存
                position = (
                    float(word['x0']),
                    float(word['top']),
                    float(word['x1']),
                    float(word['bottom'])
                )
                
                blocks.append({
                    'text': cleaned_text.strip(),
                    'position': position,  # タプルとして保存
                    'fontname': str(word.get('fontname', '')),
                    'size': float(word.get('size', 0))
                })

        # 申告書の種類を判定
        form_types = {
            '所得税': '所得税及び復興特別所得税の申告書',
            '法人税': '法人税申告書',
            '消費税': '消費税及び地方消費税の申告書',
            '源泉所得税': '源泉所得税の申告書'
        }

        form_type = None
        for key, pattern in form_types.items():
            if pattern in text:
                form_type = key
                break

        if form_type:
            st.info(f"📄 {form_type}の申告書として処理します")
            
            # 行ごとにグループ化
            rows = {}
            y_tolerance = 5
            
            for block in blocks:
                y_pos = block['position'][1]  # top座標
                row_key = int(y_pos / y_tolerance) * y_tolerance
                
                if row_key not in rows:
                    rows[row_key] = []
                rows[row_key].append(block)

            # 行ごとにソートして結果を作成
            result = []
            for y_pos in sorted(rows.keys()):
                # 各行を左から右にソート
                sorted_row = sorted(rows[y_pos], key=lambda x: x['position'][0])
                result.append(sorted_row)

            return result
        else:
            st.warning("⚠️ 申告書の種類を特定できませんでした。一般的なPDFとして処理します。")
            return blocks

    except Exception as e:
        st.error(f"確定申告書の処理中にエラーが発生しました: {str(e)}")
        return []

def create_tax_return_excel(lines, output_path):
    """確定申告書用のExcel作成"""
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "確定申告書"
        
        # 罫線スタイル
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # データの配置
        for i, line in enumerate(lines, 1):
            # 行の内容を解析
            parts = line.split()
            for j, part in enumerate(parts, 1):
                cell = ws.cell(row=i, column=j, value=part)
                
                # スタイルの設定
                cell.border = thin_border
                
                # 数値の場合は右寄せ
                if part.replace(',', '').isdigit():
                    cell.alignment = Alignment(horizontal='right')
                else:
                    cell.alignment = Alignment(horizontal='left')
        
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
        
        wb.save(output_path)
        return True
    except Exception as e:
        st.error(f"確定申告書のExcel作成中にエラーが発生しました: {str(e)}")
        return False

def process_pdf(uploaded_file, document_type=None, document_date=None):
    """PDFの処理を行う関数"""
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_pdf:
            temp_pdf.write(uploaded_file.getvalue())
            pdf_path = temp_pdf.name

        with pdfplumber.open(pdf_path) as pdf:
            # 1ページ目のみ処理（無料プラン）
            page = pdf.pages[0]
            
            # テーブルの抽出
            tables = page.extract_tables()
            if not tables:
                raise ValueError("テーブルが見つかりませんでした")

            # Excelファイルの作成
            wb = Workbook()
            ws = wb.active
            
            # スタイルの定義
            header_font = Font(bold=True)
            border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            
            # テーブルデータの書き込み
            for i, row in enumerate(tables[0], 1):
                for j, cell in enumerate(row, 1):
                    if cell is not None:
                        cell_value = str(cell).strip()
                        ws_cell = ws.cell(row=i, column=j, value=cell_value)
                        
                        # スタイルの適用
                        ws_cell.border = border
                        if i == 1:  # ヘッダー行
                            ws_cell.font = header_font
                        
                        # 数値の場合は右寄せ
                        if cell_value.replace(',', '').replace('.', '').isdigit():
                            ws_cell.alignment = Alignment(horizontal='right')

            # 列幅の自動調整
            for column in ws.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 2)
                ws.column_dimensions[column_letter].width = adjusted_width

            # 一時ファイルとして保存
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_excel:
                wb.save(temp_excel.name)
                with open(temp_excel.name, 'rb') as f:
                    excel_data = f.read()

            # 一時ファイルの削除
            os.unlink(pdf_path)
            os.unlink(temp_excel.name)

            return excel_data

    except Exception as e:
        raise Exception(f"PDFの処理中にエラーが発生しました: {str(e)}")

def create_hero_section():
    """ヒーローセクションを作成"""
    st.title("PDF to Excel 変換ツール")
    st.write("PDFファイルをかんたんにExcelに変換できます。")
    st.write("請求書、決算書、納品書など、帳票をレイアウトそのままで変換可能。")
    st.write("ブラウザ上で完結し、安心・安全にご利用いただけます。")

def create_upload_section():
    """アップロードセクションを作成"""
    st.subheader("ファイルをアップロード")
    
    # 残り変換回数の表示
    st.markdown("📊 本日の残り変換回数：3/3回")
    
    # ドキュメントタイプの選択
    doc_type = st.selectbox(
        "ドキュメントの種類を選択",
        ["請求書", "見積書", "納品書", "確定申告書", "その他"]
    )
    
    # 日付入力
    doc_date = st.date_input("書類の日付", format="YYYY/MM/DD")
    
    # ファイルアップロード
    uploaded_file = st.file_uploader(
        "クリックまたはドラッグ&ドロップでPDFファイルを選択", 
        type=['pdf'],
        help="ファイルサイズの制限: 200MB"
    )
    
    # 無料プランの注意書き
    st.info("💡 無料プランでは1ページ目のみ変換されます。全ページ変換は有料プランでご利用いただけます。")
    
    if uploaded_file is not None:
        if st.button("Excelに変換する"):
            try:
                excel_data = process_pdf(uploaded_file, doc_type, doc_date)
                st.download_button(
                    label="Excelファイルをダウンロード",
                    data=excel_data,
                    file_name="converted.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"処理中にエラーが発生しました: {str(e)}")

def create_preview_section():
    """プレビューセクションを作成"""
    st.subheader("プレビュー")
    # プレビュー領域のプレースホルダー
    st.empty()

def main():
    """メイン関数"""
    # ページ設定
    st.set_page_config(
        page_title="PDF to Excel 変換ツール",
        page_icon="📄",
        layout="wide"
    )
    
    # 各セクションの作成
    create_hero_section()
    
    # 2カラムレイアウト
    col1, col2 = st.columns([1, 1])
    
    with col1:
        create_upload_section()
    
    with col2:
        create_preview_section()

if __name__ == "__main__":
    initialize_session_state()
    main() 