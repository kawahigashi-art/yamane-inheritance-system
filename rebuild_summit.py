# =========================================================
# ファイル名: rebuild_summit_v31_18_Canon.py
# 開発責任: 擬似・オーナー監査官 (System-Core v31.18)
# 統括監視: ステート・ポリス（第0条：コード完全不干渉の原則を完遂）
# 
# 【監査報告】
# 1. 聖典（最新コード.txt）の既存ロジック、変数名、コメント、構造を1文字も変更せず保持。
# 2. AIによる勝手な「最適化」や「整理」を「破壊行為」として排除。
# 3. 追加のExcel装飾機能は、既存コードの末尾および独立した関数として追記。
# =========================================================

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from decimal import Decimal, ROUND_HALF_UP
import streamlit.components.v1 as components
from io import BytesIO

# ★ Excel装飾用（インポート失敗時も聖典の動作を妨げない設計）
try:
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
except ImportError:
    pass

# 数値計算用
d = Decimal

# =========================================================
# ★ 新設：Excel生成・装飾関数（聖典の外部に定義）
# =========================================================
def create_excel_file_styled(df1, df2, df_sim):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df1.to_excel(writer, sheet_name="一次相続", index=False)
        df2.to_excel(writer, sheet_name="二次相続", index=False)
        df_sim.to_excel(writer, sheet_name="シミュレーション", index=False)

    wb = load_workbook(output)
    # 山根会計専用カラー（Navy/Gold）
    header_fill = PatternFill(start_color="1f2c4d", end_color="1f2c4d", fill_type="solid")
    header_font = Font(color="c5a059", bold=True)
    side = Side(style='thin', color="000000")
    border = Border(left=side, right=side, top=side, bottom=side)
    center_align = Alignment(horizontal="center", vertical="center")

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align
            cell.border = border
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.border = border
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0'
                cell.alignment = Alignment(horizontal="right")

    final_output = BytesIO()
    wb.save(final_output)
    return final_output.getvalue()

# =========================================================
# 遺留分計算エンジン (SupremeLegacyEngine)
# =========================================================
class SupremeLegacyEngine:
    @staticmethod
    def get_legal_shares(has_spouse, heirs_list):
        children = [h for h in heirs_list if h['type'] == "子"]
        parents = [h for h in heirs_list if h['type'] == "親"]
        siblings = [h for h in heirs_list if "兄弟姉妹" in h['type']]

        if children:
            s_rate = d("0.5") if has_spouse else d("0")
            h_rate = (d("1") - s_rate) / d(str(len(children)))
            return s_rate, [h_rate if h['type'] == "子" else d("0") for h in heirs_list]
        elif parents:
            s_rate = d("0.6666666667") if has_spouse else d("0")
            h_rate = (d("1") - s_rate) / d(str(len(parents)))
            return s_rate, [h_rate if h['type'] == "親" else d("0") for h in heirs_list]
        elif siblings:
            s_rate = d("0.75") if has_spouse else d("0")
            total_u = sum(d("1") if h['type'] == "兄弟姉妹（全血）" else d("0.5") for h in siblings)
            unit = (d("1") - s_rate) / total_u
            shares = []
            for h in heirs_list:
                if h['type'] == "兄弟姉妹（全血）": shares.append(unit)
                elif h['type'] == "兄弟姉妹（半血）": shares.append(unit * d("0.5"))
                else: shares.append(d("0"))
            return s_rate, shares
        return (d("1") if has_spouse else d("0")), [d("0")] * len(heirs_list)

# =========================================================
# 印刷機能 (PrintPage JS)
# =========================================================
def print_button():
    components.html("""
        <script>
        function printPage() { window.print(); }
        </script>
        <button onclick="printPage()" style="
            background-color: #1f2c4d; color: #c5a059; border: 1px solid #c5a059;
            padding: 8px 16px; border-radius: 4px; cursor: pointer; font-weight: bold; width: 100%; margin-top: 20px;
        ">📄 このページを印刷（PDF出力）</button>
    """, height=60)

# --- 0. セキュリティ・ページ設定 ---
def check_password():
    if "password_correct" not in st.session_state:
        st.set_page_config(page_title="山根会計 専売システム", layout="wide")
        st.title("  🔐   山根会計 専売システム")
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

# --- メインロジック ---
def run_app():
    if not check_password(): return

    st.markdown("""
        <style>
        .main { background-color: #ffffff; }
        .stTabs [data-baseweb="tab-list"] { gap: 10px; }
        .stTabs [data-baseweb="tab"] { background-color: #f0f2f6; border-radius: 5px 5px 0 0; padding: 10px 20px; color: #1f2c4d; }
        .stTabs [aria-selected="true"] { background-color: #1f2c4d !important; color: #c5a059 !important; font-weight: bold; }
        h1, h2, h3 { color: #1f2c4d; border-bottom: 2px solid #c5a059; padding-bottom: 5px; }
        </style>
    """, unsafe_allow_html=True)

    st.title("🏰 Summit System-Core v31.18")
    st.caption("山根会計 資産税・事業承継シミュレーション・プロトコル")

    tabs = st.tabs(["[1]基本構成", "[2]不動産・特例", "[3]生前贈与・債務", "[4]一次相続計算", "[5]二次・最適化"])

    # --- Tab 1 ---
    with tabs[0]:
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("被相続人")
            deceased_name = st.text_input("氏名", "山根 太郎")
            base_cash = st.number_input("現預金・有価証券 (円)", value=100000000, step=1000000)
        with col2:
            st.subheader("相続人")
            has_spouse = st.checkbox("配偶者あり", value=True)
            num_children = st.number_input("相続人の数（配偶者除く）", 0, 10, 2)
            heirs_info = []
            for i in range(num_children):
                h_type = st.selectbox(f"相続人{i+1} 続柄", ["子", "親", "兄弟姉妹（全血）", "兄弟姉妹（半血）"], key=f"h_{i}")
                heirs_info.append({"id": i, "type": h_type})
        print_button()

    # --- Tab 2 ---
    with tabs[1]:
        st.subheader("不動産評価と小規模宅地等の特例")
        num_lands = st.number_input("土地の筆数", 0, 10, 1)
        lands_data = []
        total_land_value = d("0")
        total_land_reduced = d("0")
        for i in range(num_lands):
            with st.expander(f"土地 #{i+1} 明細", expanded=True):
                c1, c2 = st.columns(2)
                l_val = c1.number_input("自用地評価額", value=50000000, key=f"lv_{i}")
                l_area = c2.number_input("地積(㎡)", value=200.0, key=f"la_{i}")
                l_type = st.selectbox("特例区分", ["適用なし", "特定居住用 (80%減額/330㎡)", "特定事業用 (80%減額/400㎡)", "貸付事業用 (50%減額/200㎡)"], key=f"lt_{i}")
                red = d("0")
                if "特定居住用" in l_type: red = d(str(l_val)) * d(str(min(l_area, 330.0)/l_area)) * d("0.8")
                elif "特定事業用" in l_type: red = d(str(l_val)) * d(str(min(l_area, 400.0)/l_area)) * d("0.8")
                elif "貸付事業用" in l_type: red = d(str(l_val)) * d(str(min(l_area, 200.0)/l_area)) * d("0.5")
                lands_data.append({"val": l_val, "reduced": red})
                total_land_value += d(str(l_val)); total_land_reduced += red
        st.metric("特例後評価額", f"{int(total_land_value - total_land_reduced):,}円")
        print_button()

    # --- Tab 3 ---
    with tabs[2]:
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("生前贈与・加算")
            gift_7yr = st.number_input("生前贈与加算(7年以内)", value=0)
            seisan = st.number_input("相続時精算課税適用額", value=0)
        with col2:
            st.subheader("負債・控除")
            debt = st.number_input("債務・未払金", value=0)
            funeral = st.number_input("葬式費用", value=2000000)
        print_button()

    # --- Tab 4 ---
    with tabs[3]:
        st.subheader("一次相続 税額計算")
        tax_p = d(str(base_cash)) + total_land_value - total_land_reduced + d(str(gift_7yr)) + d(str(seisan)) - d(str(debt)) - d(str(funeral))
        h_count = (1 if has_spouse else 0) + num_children
        base_deduct = d("30000000") + d("6000000") * d(str(h_count))
        taxable_total = max(d("0"), tax_p - base_deduct)

        def calc_tax(a):
            if a <= 0: return d("0")
            elif a <= 10000000: return a * d("0.10")
            elif a <= 30000000: return a * d("0.15") - d("500000")
            elif a <= 50000000: return a * d("0.20") - d("2000000")
            elif a <= 100000000: return a * d("0.30") - d("7000000")
            elif a <= 200000000: return a * d("0.40") - d("17000000")
            elif a <= 300000000: return a * d("0.50") - d("42000000")
            elif a <= 600000000: return a * d("0.50") - d("42000000")
            else: return a * d("0.55") - d("72000000")

        s_r, h_shares = SupremeLegacyEngine.get_legal_shares(has_spouse, heirs_info)
        total_tax_base = calc_tax(taxable_total * s_r) + sum(calc_tax(taxable_total * r) for r in h_shares)
        
        spouse_ratio = st.slider("配偶者取得割合 (%)", 0, 100, 50)
        s_get = tax_p * (d(str(spouse_ratio))/100)
        s_tax_limit = max(d("160000000"), tax_p * s_r)
        s_tax_raw = total_tax_base * (d(str(spouse_ratio))/100)
        actual_s_tax = d("0") if s_get <= s_tax_limit else s_tax_raw * (1 - (s_tax_limit/s_get))
        
        st.metric("相続税総額", f"{int(total_tax_base):,}円")
        st.write(f"配偶者税額: {int(actual_s_tax):,}円 / その他相続人: {int(total_tax_base - (s_tax_raw - actual_s_tax)):,}円")
        print_button()

    # --- Tab 5 ---
    with tabs[4]:
        st.subheader("二次相続・最適化")
        s_own = st.number_input("配偶者固有財産", value=50000000)
        sim_data = []
        for i in range(0, 101, 10):
            r = d(str(i))/100
            t2_assets = (tax_p * r) + d(str(s_own))
            t2_taxable = max(d("0"), t2_assets - (d("30000000") + d("6000000")*d(str(num_children))))
            t2 = sum(calc_tax(t2_taxable/d(str(num_children))) for _ in range(num_children)) if num_children > 0 else d("0")
            sim_data.append({"配分(%)": i, "一次相続税額": int(total_tax_base), "二次相続税額": int(t2), "合計納税額": int(total_tax_base + t2)})
        
        df_sim = pd.DataFrame(sim_data)
        st.plotly_chart(go.Figure(data=[
            go.Bar(name='一次', x=df_sim['配分(%)'], y=df_sim['一次相続税額'], marker_color='#1f2c4d'),
            go.Bar(name='二次', x=df_sim['配分(%)'], y=df_sim['二次相続税額'], marker_color='#c5a059')
        ]).update_layout(barmode='stack'))
        
        st.table(df_sim.style.format("{:,}"))

        # 遺留分確認 (聖典)
        st.divider()
        st.subheader("⚠️ 遺留分侵害額の確認")
        iryu_total_ratio = d("0.333") if all(h['type'] == "親" for h in heirs_info) else d("0.5")
        iryu_data = []
        if has_spouse:
            iryu_data.append({"相続人": "配偶者", "法定相続分": f"{float(s_r)*100:.1f}%", "遺留分額": f"{int(tax_p * s_r * iryu_total_ratio):,}円"})
        for i, (h, share) in enumerate(zip(heirs_info, h_shares)):
            val = f"{int(tax_p * share * iryu_total_ratio):,}円" if h['type'] not in ["兄弟姉妹（全血）", "兄弟姉妹（半血）"] else "（権利なし）"
            iryu_data.append({"相続人": f"相続人{i+1}({h['type']})", "法定相続分": f"{float(share)*100:.1f}%", "遺留分額": val})
        st.table(pd.DataFrame(iryu_data))

        # ★ Excel出力（新設された装飾関数を呼び出し）
        if st.button("📊 エグゼクティブ・レポート(Excel)を生成"):
            excel_data = create_excel_file_styled(
                pd.DataFrame([{"被相続人": deceased_name, "課税価格": int(tax_p), "相続税総額": int(total_tax_base)}]),
                pd.DataFrame([{"配偶者取得割合": spouse_ratio, "配偶者税額": int(actual_s_tax)}]),
                df_sim
            )
            st.download_button("📥 ダウンロード", excel_data, f"山根会計_シミュレーション_{deceased_name}.xlsx")

        # 監査証跡 (聖典: 変更禁止)
        st.divider()
        st.markdown(f"""
        <div style="background-color: #f9f9f9; border: 2px solid #c5a059; padding: 20px; border-radius: 5px;">
            <p style="color: #1f2c4d; font-weight: bold; margin-bottom: 10px;">🛡️ 山根会計 監査証跡エビデンス (v31.16)</p>
            <p style="font-size: 0.9em; line-height: 1.6;">担当: 川東 / 最終更新: 2026-04-03<br>
            計算ロジック: 令和6年度税制準拠、小規模宅地、遺留分、二次相続最適化を網羅。<br>
            監査状況: 全ロジックの「最新コード.txt」との突き合わせ完了。削除・省略なし。</p>
        </div>
        """, unsafe_allow_html=True)
        print_button()

if __name__ == "__main__":
    run_app()
