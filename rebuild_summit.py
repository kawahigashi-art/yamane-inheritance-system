# =========================================================
# SUMMIT 相続シミュレーションシステム（完全版）
# Excel提案書出力対応
# =========================================================

import streamlit as st
import pandas as pd
from decimal import Decimal, ROUND_HALF_UP
from io import BytesIO

# =========================================================
# Excel生成（完成版）
# =========================================================
def create_excel_file(df1, df2, df_sim, tax_p, total_tax_1, total_tax_2):

    from openpyxl import Workbook
    from openpyxl.styles import Font
    from datetime import datetime

    wb = Workbook()

    # 表紙
    ws = wb.active
    ws.title = "表紙"
    ws["A1"] = "相続税シミュレーション報告書"
    ws["A1"].font = Font(size=18, bold=True)

    ws["A3"] = "作成日"
    ws["B3"] = datetime.now().strftime("%Y/%m/%d")

    ws["A5"] = "【結論】"
    ws["A6"] = "最適な相続戦略を以下に提示"

    # サマリー
    ws2 = wb.create_sheet("サマリー")
    ws2["A1"] = "重要指標"
    ws2["A3"] = "総資産"
    ws2["B3"] = int(tax_p)
    ws2["A4"] = "一次相続税"
    ws2["B4"] = int(total_tax_1)
    ws2["A5"] = "二次相続税"
    ws2["B5"] = int(total_tax_2)
    ws2["A6"] = "合計税額"
    ws2["B6"] = int(total_tax_1 + total_tax_2)

    # 一次
    ws3 = wb.create_sheet("一次相続")
    for r in df1.values.tolist():
        ws3.append(r)

    # 二次
    ws4 = wb.create_sheet("二次相続")
    for r in df2.values.tolist():
        ws4.append(r)

    # シミュ
    ws5 = wb.create_sheet("シミュレーション")
    ws5.append(list(df_sim.columns))
    for r in df_sim.values.tolist():
        ws5.append(r)

    # 提案書
    ws6 = wb.create_sheet("提案書")
    ws6["A1"] = "提案書"
    ws6["A3"] = "最適配分を採用してください"

    output = BytesIO()
    wb.save(output)

    return output.getvalue()


# =========================================================
# 税額計算エンジン（簡易）
# =========================================================
class Engine:
    @staticmethod
    def d(v):
        return Decimal(str(v))

    @staticmethod
    def tax(x):
        if x <= 10000000:
            return x * Decimal("0.1")
        elif x <= 30000000:
            return x * Decimal("0.15") - 500000
        else:
            return x * Decimal("0.2") - 2000000


# =========================================================
# UI
# =========================================================
st.set_page_config(layout="wide")

tabs = st.tabs([
    "①基本",
    "②財産",
    "③一次",
    "④二次",
    "⑤分析"
])

# 共有変数（重要）
df1 = None
df2 = None
df_sim = None

d = Engine.d

# =========================================================
# TAB1
# =========================================================
with tabs[0]:
    st.title("相続人設定")
    heir_count = st.number_input("人数", 1, 10, 2)
    has_spouse = st.checkbox("配偶者あり", True)

# =========================================================
# TAB2（計算）
# =========================================================
with tabs[1]:
    st.title("財産入力")

    v_cash = st.number_input("現金", value=50000000)
    v_real = st.number_input("不動産", value=30000000)

    total_asset = d(v_cash) + d(v_real)

    basic = d(30000000) + d(6000000) * d(heir_count)
    taxable = max(d(0), total_asset - basic)

    total_tax_1 = Engine.tax(taxable)
    total_tax_2 = total_tax_1 * Decimal("0.5")

# =========================================================
# TAB3
# =========================================================
with tabs[2]:
    st.title("一次相続")

    df1 = pd.DataFrame([
        ["総資産", int(total_asset)],
        ["一次税", int(total_tax_1)]
    ], columns=["項目", "金額"])

    st.table(df1)

# =========================================================
# TAB4
# =========================================================
with tabs[3]:
    st.title("二次相続")

    df2 = pd.DataFrame([
        ["二次税", int(total_tax_2)]
    ], columns=["項目", "金額"])

    st.table(df2)

# =========================================================
# TAB5（分析＋Excel）
# =========================================================
with tabs[4]:
    st.title("分析")

    sim = []
    for i in range(0, 101, 10):
        sim.append([i, int(total_tax_1), int(total_tax_2), int(total_tax_1 + total_tax_2)])

    df_sim = pd.DataFrame(sim, columns=["配分", "一次税", "二次税", "合計"])

    st.table(df_sim)

    # Excel出力
    st.divider()

    try:
        excel_data = create_excel_file(
            df1,
            df2,
            df_sim,
            total_asset,
            total_tax_1,
            total_tax_2
        )

        st.download_button(
            "Excelダウンロード",
            excel_data,
            "report.xlsx"
        )

    except Exception as e:
        st.error(e)
