# =========================================================
# ファイル名: rebuild_summit_v31_27_final.py
# 開発責任: 擬似・オーナー監査官 (System-Core v31.27)
# 統括監視: エラー対策専門チーム（デバッグ・ガーディアンズ） [cite: 12]
# 出力統括: ドキュメント・エンジニア ＆ アーティスティック・ディレクター [cite: 6, 9]
# =========================================================

import streamlit as st
import pandas as pd
from decimal import Decimal, ROUND_HALF_UP
import io

# --- 0. システム最優先設定 (エラー回避のため冒頭に配置) ---
def init_system():
    if 'init' not in st.session_state:
        st.set_page_config(page_title="SUMMIT v31.27 PRO", layout="wide")
        st.session_state['init'] = True

init_system()

# --- 1. ライブラリ動的チェック (環境依存エラー防止)  ---
try:
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
    HAS_LIBS = True
except ImportError:
    HAS_LIBS = False

# --- 2. セキュリティ設定 ---
def check_password():
    if "password_correct" not in st.session_state:
        st.title("🔐 山根会計 専売システム")
        valid_password = "yamane777"
        pwd = st.text_input("アクセスパスワード", type="password", key="main_login_pwd")
        if st.button("ログイン", key="main_login_btn"):
            if pwd == valid_password:
                st.session_state["password_correct"] = True
                st.rerun()
            else:
                st.error("パスワードが正しくありません。")
        return False
    return True

# --- 3. 超精密計算エンジン (聖典遵守・ロジック無削除) [cite: 2, 4] ---
class SupremeLegacyEngine:
    @staticmethod
    def to_d(val): 
        try: return Decimal(str(val))
        except: return Decimal("0")

    @staticmethod
    def get_tax(taxable_amt, has_spouse, heir_count):
        d = SupremeLegacyEngine.to_d
        if taxable_amt <= 0: return d(0)
        
        # 法定相続分計算 (v31.23からの継承ロジックを維持) [cite: 4]
        st_count = d(heir_count) + (d(1) if has_spouse else d(0))
        s_ratio = d("0.5") if has_spouse else d(0)
        h_total_ratio = d("0.5") if has_spouse else d("1.0")
        per_h_ratio = h_total_ratio / d(max(1, heir_count))

        def bracket(a):
            if a <= 10000000: return a * d("0.10")
            elif a <= 30000000: return a * d("0.15") - d("500000")
            elif a <= 50000000: return a * d("0.20") - d("2000000")
            elif a <= 100000000: return a * d("0.30") - d("7000000")
            elif a <= 200000000: return a * d("0.40") - d("17000000")
            elif a <= 300000000: return a * d("0.45") - d("27000000")
            elif a <= 600000000: return a * d("0.50") - d("42000000")
            else: return a * d("0.55") - d("72000000")

        total_tax = d(0)
        if has_spouse:
            total_tax += bracket(taxable_amt * s_ratio)
        for _ in range(int(heir_count)):
            total_tax += bracket(taxable_amt * per_h_ratio)
            
        return total_tax.quantize(d(1), ROUND_HALF_UP)

# --- 4. メインUI ---
if check_password():
    st.sidebar.markdown("### 🏢 山根会計 専売システム")
    st.sidebar.info("ログイン: 川東")
    d = SupremeLegacyEngine.to_d

    # エラー対策：全タブ共通で必要な変数をサイドバーで一元管理 
    st.sidebar.header("基本設定")
    in_child = st.sidebar.number_input("相続人の人数（子）", 1, 10, 2, key="in_child_v3")
    in_spouse = st.sidebar.checkbox("配偶者は健在", value=True, key="in_spouse_v3")

    tabs = st.tabs(["💰 財産入力", "📑 計算明細", "📄 報告書出力"])

    with tabs[0]:
        st.header("一次相続：精密財産入力")
        ca, cb = st.columns(2)
        with ca:
            v_home = d(st.number_input("居住用宅地：評価額", value=32781936))
            a_home = d(st.number_input("居住用宅地：面積(㎡)", value=330))
            v_cash = d(st.number_input("現預金合計", value=45573502))
        with cb:
            v_stock = d(st.number_input("有価証券合計", value=45132788))
            v_ins = d(st.number_input("生命保険金", value=3651514))
            v_debt = d(st.number_input("債務・葬式費用", value=363580))

    # --- 計算実行コア (エラー対策チームによる堅牢化)  ---
    st_count = (1 if in_spouse else 0) + in_child
    # ゼロ除算ガード
    safe_area = max(d("0.1"), a_home)
    red_home = (v_home / safe_area) * min(a_home, d(330)) * d("0.8")
    ins_ded = min(v_ins, d(5000000) * d(st_count))
    
    # 正味財産の算出 (削除・中略禁止) [cite: 2]
    pure_as = max(d(0), (v_home - red_home) + v_cash + v_stock + max(d(0), v_ins - ins_ded) - v_debt)
    basic_1 = d(30000000) + (d(6000000) * d(st_count))
    taxable_1 = max(d(0), pure_as - basic_1)
    total_tax_1 = SupremeLegacyEngine.get_tax(taxable_1, in_spouse, in_child)

    with tabs[1]:
        st.subheader("山根会計 専売：相続税計算明細")
        res_df = pd.DataFrame({
            "項目": ["正味財産価格", "基礎控除額", "課税遺産総額", "相続税の総額"],
            "金額（円）": [f"{int(pure_as):,}", f"△{int(basic_1):,}", f"{int(taxable_1):,}", f"{int(total_tax_1):,}"]
        })
        st.table(res_df)

    with tabs[2]:
        st.header("エグゼクティブ報告書")
        if HAS_LIBS:
            if st.button("📄 報告書PDF生成"):
                try:
                    buf = io.BytesIO()
                    canvas_obj = canvas.Canvas(buf, pagesize=A4)
                    canvas_obj.setFont("Helvetica-Bold", 16)
                    canvas_obj.drawString(50, 800, "Yamane Accounting Inheritance Report")
                    canvas_obj.setFont("Helvetica", 12)
                    canvas_obj.drawString(50, 780, f"Net Assets: {int(pure_as):,} JPY")
                    canvas_obj.drawString(50, 760, f"Total Tax: {int(total_tax_1):,} JPY")
                    canvas_obj.save()
                    st.download_button("PDFをダウンロード", buf.getvalue(), "Yamane_Summit_v31_27.pdf")
                except Exception as e:
                    st.error(f"PDF生成中にエラーが発生しました: {e}")
        else:
            st.warning("PDF生成ライブラリ(ReportLab)が未検出のため、この機能は無効です。")

st.sidebar.success("✅ System-Core v31.27 正常稼働（エラー対策済）")
