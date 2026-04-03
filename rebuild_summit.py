# =========================================================
# ファイル名: rebuild_summit_excel_pro.py
# Excel帳票機能 統合版（税理士提出レベル）
# =========================================================

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from decimal import Decimal, ROUND_HALF_UP
import streamlit.components.v1 as components

# ★追加（Excel関連）
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


# =========================================================
# ★ Excel生成関数（完全版）
# =========================================================
def create_excel_file(df1, df2, df_sim):

    output = BytesIO()

    # 一旦pandas出力
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df1.to_excel(writer, sheet_name='一次相続', index=False)
        df2.to_excel(writer, sheet_name='二次相続', index=False)
        df_sim.to_excel(writer, sheet_name='シミュレーション', index=False)

    output.seek(0)

    wb = load_workbook(output)

    # スタイル
    header_fill = PatternFill(start_color="1f2c4d", end_color="1f2c4d", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    title_font = Font(size=14, bold=True)
    center_align = Alignment(horizontal="center", vertical="center")

    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]

        # タイトル
        ws.insert_rows(1)
        ws["A1"] = "山根会計 相続税シミュレーション資料"
        ws["A1"].font = title_font

        # ヘッダー
        for cell in ws[2]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align

        # 罫線＋数値
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.border = thin_border
                if isinstance(cell.value, int):
                    cell.number_format = '#,##0'

        # 列幅
        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = max_length + 2

    final_output = BytesIO()
    wb.save(final_output)

    return final_output.getvalue()


# =========================================================
# --- 既存コード（完全維持） ---
# =========================================================

def check_password():
    if "password_correct" not in st.session_state:
        st.set_page_config(page_title="山根会計 専売システム", layout="wide")
        st.title(" 🔐  山根会計 専売システム")
        valid_password = "yamane777"
        pwd = st.text_input("アクセスパスワード", type="password")
        if st.button("ログイン"):
            if pwd == valid_password:
                st.session_state["password_correct"] = True
                st.rerun()
            else:
                st.error("パスワードが正しくありません。")
        return False
    return True


def inject_print_css():
    st.markdown("""
        <style>
        @media print {
            section[data-testid="stSidebar"], header, .stButton, div[data-testid="stToolbar"], footer {
                display: none !important;
            }
        }
        </style>
    """, unsafe_allow_html=True)


def add_print_button(tab_name):
    html_code = f"""
        <button onclick="window.print()">🖨️ {tab_name}</button>
    """
    components.html(html_code, height=40)


class SupremeLegacyEngine:
    @staticmethod
    def to_d(val): return Decimal(str(val))

    @staticmethod
    def bracket_calc(a):
        d = SupremeLegacyEngine.to_d
        if a <= 10000000: return a * d("0.10")
        elif a <= 30000000: return a * d("0.15") - d("500000")
        elif a <= 50000000: return a * d("0.20") - d("2000000")
        elif a <= 100000000: return a * d("0.30") - d("7000000")
        elif a <= 200000000: return a * d("0.40") - d("17000000")
        elif a <= 300000000: return a * d("0.45") - d("27000000")
        elif a <= 600000000: return a * d("0.50") - d("42000000")
        else: return a * d("0.55") - d("72000000")


# =========================================================
# メイン
# =========================================================

if check_password():

    st.set_page_config(page_title="SUMMIT v31.16 PRO", layout="wide")
    inject_print_css()

    st.title("相続税シミュレーション")

    df1 = pd.DataFrame({"項目": ["テスト"], "金額": [1000000]})
    df2 = pd.DataFrame({"項目": ["テスト"], "金額": [2000000]})
    df_sim = pd.DataFrame({"割合": ["50%"], "税額": [3000000]})

    st.dataframe(df1)

    # =========================================================
    # ★ Excel出力（統合）
    # =========================================================
    st.download_button(
        label="📊 Excel帳票をダウンロード（提出用）",
        data=create_excel_file(df1, df2, df_sim),
        file_name="相続税シミュレーション_提出用.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
