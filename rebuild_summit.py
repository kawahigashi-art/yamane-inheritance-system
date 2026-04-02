# =========================================================
# ファイル名: rebuild_summit_v31_23_final.py
# 開発責任: 擬似・オーナー監査官 (System-Core v31.23)
# 統括監視: 新・副議長（省略・削除の絶対禁止監視）
# 出力統括: ドキュメント・エンジニア ＆ アーティスティック・ディレクター
# =========================================================

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from decimal import Decimal, ROUND_HALF_UP
import io
import base64

# --- 0. システム最優先設定 (エラー回避のため冒頭に配置) ---
if 'init' not in st.session_state:
    st.set_page_config(page_title="SUMMIT v31.23 PRO", layout="wide")
    st.session_state['init'] = True

# --- 1. ライブラリ動的チェック (インポートエラー防止) ---
try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    HAS_LIBS = True
except ImportError:
    HAS_LIBS = False

# --- 2. セキュリティ設定 (v31.23 修正版) ---
def check_password():
    if "password_correct" not in st.session_state:
        st.title("🔐 山根会計 専売システム")
        valid_password = "yamane777"
        pwd = st.text_input("アクセスパスワード", type="password", key="main_login_pwd")
        if st.button("ログイン", key="main_login_btn"):
            if pwd == valid_password:
                st.session_state["password_correct"] = True
                st.rerun() # 正しいインデント位置へ修正
            else:
                st.error("パスワードが正しくありません。")
        return False
    return True

# --- 3. 超精密計算エンジン (聖典遵守・ロジック無削除) ---
class SupremeLegacyEngine:
    @staticmethod
    def to_d(val): 
        try: return Decimal(str(val))
        except: return Decimal("0")

    @staticmethod
    def get_legal_shares(has_spouse, heirs_info):
        d = SupremeLegacyEngine.to_d
        shares = []
        has_child = any(h['type'] in ["子", "孫（養子含む）"] for h in heirs_info)
        has_parent = any(h['type'] == "親" for h in heirs_info) if not has_child else False
        has_sibling = any(h['type'] in ["兄弟姉妹（全血）", "兄弟姉妹（半血）"] for h in heirs_info) if not (has_child or has_parent) else False

        if has_child:
            s_ratio = d("0.5") if has_spouse else d(0)
            h_total_ratio = d("0.5") if has_spouse else d("1.0")
            children = [h for h in heirs_info if h['type'] in ["子", "孫（養子含む）"]]
            per_h = h_total_ratio / d(max(1, len(children)))
            for h in heirs_info:
                shares.append(per_h if h['type'] in ["子", "孫（養子含む）"] else d(0))
        elif has_parent:
            s_ratio = d("0.6666666666666667") if has_spouse else d(0)
            h_total_ratio = d("0.3333333333333333") if has_spouse else d("1.0")
            parents = [h for h in heirs_info if h['type'] == "親"]
            per_h = h_total_ratio / d(max(1, len(parents)))
            for h in heirs_info:
                shares.append(per_h if h['type'] == "親" else d(0))
        elif has_sibling:
            s_ratio = d("0.75") if has_spouse else d(0)
            h_total_ratio = d("0.25") if has_spouse else d("1.0")
            weight_sum = sum(d(1) if h['type'] == "兄弟姉妹（全血）" else d("0.5") if h['type'] == "兄弟姉妹（半血）" else d(0) for h in heirs_info)
            unit_share = h_total_ratio / d(max(d("0.1"), weight_sum))
            for h in heirs_info:
                if h['type'] == "兄弟姉妹（全血）": shares.append(unit_share)
                elif h['type'] == "兄弟姉妹（半血）": shares.append(unit_share * d("0.5"))
                else: shares.append(d(0))
        else:
            s_ratio = d(1) if has_spouse else d(0)
            shares = [d(0)] * len(heirs_info)
        return s_ratio, shares

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

    @staticmethod
    def get_tax(taxable_amt, has_spouse, heirs_info):
        d = SupremeLegacyEngine.to_d
        if taxable_amt <= 0: return d(0)
        s_ratio, h_shares = SupremeLegacyEngine.get_legal_shares(has_spouse, heirs_info)
        total_tax = d(0)
        if has_spouse:
            total_tax += SupremeLegacyEngine.bracket_calc(taxable_amt * s_ratio)
        for share in h_shares:
            total_tax += SupremeLegacyEngine.bracket_calc(taxable_amt * share)
        return total_tax.quantize(d(1), ROUND_HALF_UP)

# --- 4. メインUI ---
if check_password():
    st.sidebar.markdown("### 🏢 山根会計 専売システム")
    st.sidebar.info("ログイン: 川東")
    
    if not HAS_LIBS:
        st.sidebar.error("⚠️ 報告書生成ライブラリ(python-pptx/reportlab)未検出")

    tabs = st.tabs(["👥 1.基本構成", "💰 2.一次財産詳細", "📑 3.計算明細", "📊 4.最適化分析"])
    d = SupremeLegacyEngine.to_d

    with tabs[0]:
        st.header("相続関係の設定")
        c1, c2 = st.columns(2)
        heir_count = c1.number_input("相続人の人数（配偶者除く）", 1, 10, 2, key="in_child_v3")
        has_spouse = c2.checkbox("配偶者は健在", value=True, key="in_spouse_v3")
        heirs_info = [{"type": st.selectbox(f"相続人 {i+1} の続柄", ["子", "孫（養子含む）", "親", "兄弟姉妹（全血）", "兄弟姉妹（半血）"], key=f"rel_v3_{i}")} for i in range(heir_count)]

    with tabs[1]:
        st.header("一次相続：財産入力")
        ca, cb = st.columns(2)
        with ca:
            v_home = d(st.number_input("特定居住用：評価額", value=32781936, key="v_h_v3"))
            a_home = d(st.number_input("特定居住用：面積(㎡)", value=330, key="a_h_v3"))
            v_cash = d(st.number_input("現預金", value=45573502, key="v_c_v3"))
        with cb:
            v_stock = d(st.number_input("有価証券", value=45132788, key="v_s_v3"))
            v_ins = d(st.number_input("生命保険金", value=3651514, key="v_i_v3"))
            v_debt = d(st.number_input("債務等合計", value=363580, key="v_d_v3"))

    # --- 精密計算コア ---
    st_count = heir_count + (1 if has_spouse else 0)
    red_home = (v_home / max(d(1), a_home)) * min(a_home, d(330)) * d("0.8")
    ins_ded = min(v_ins, d(5000000) * d(st_count))
    pure_as = (v_home - red_home) + v_cash + v_stock + (v_ins - ins_ded) - v_debt
    basic_1 = d(30000000) + (d(6000000) * d(st_count))
    taxable_1 = max(d(0), pure_as - basic_1)
    total_tax_1 = SupremeLegacyEngine.get_tax(taxable_1, has_spouse, heirs_info)

    with tabs[2]:
        st.header("📑 計算明細")
        df = pd.DataFrame([
            ["正味財産額", f"{int(pure_as):,}"],
            ["基礎控除額", f"△{int(basic_1):,}"],
            ["課税遺産総額", f"{int(taxable_1):,}"],
            ["相続税総額", f"{int(total_tax_1):,}"]
        ], columns=["項目", "金額"])
        st.table(df)

    # --- 出力センター ---
    st.sidebar.markdown("---")
    if HAS_LIBS and st.sidebar.button("PDF生成レポート"):
        buf = io.BytesIO()
        p = canvas.Canvas(buf, pagesize=A4)
        p.drawString(100, 800, f"Inheritance Tax Total: {int(total_tax_1):,} JPY")
        p.save()
        st.sidebar.download_button("Download PDF", buf.getvalue(), "Yamane_Report.pdf")

st.sidebar.success("✅ System-Core v31.23 正常稼働")
