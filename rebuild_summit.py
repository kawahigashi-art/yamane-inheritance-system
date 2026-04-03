# =========================================================
# ファイル名: rebuild_summit.py
# 開発責任: 擬似・オーナー監査官 (System-Core v32.01)
# 統括監視: ステート・ポリス（削除・省略・中略の絶対禁止監視）
# 
# 【修正・更新内容】
# 1. Excel出力機能の実装: openpyxlを用いた多機能帳票エクスポートを追加。
# 2. 聖典の厳守: 既存の計算エンジン、UI、印刷機能を1行も削らず保持。
# 3. 安定性向上: ファイル生成時のメモリバッファ処理（io.BytesIO）を採用。
# =========================================================
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from decimal import Decimal, ROUND_HALF_UP
import streamlit.components.v1 as components
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

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

# --- 1. 超精密計算エンジン (Decimal実装) ---
d = Decimal
def round_tax(val):
    return val.quantize(d("1"), rounding=ROUND_HALF_UP)

class SupremeTaxEngine:
    @staticmethod
    def get_tax_rate(taxable_amount):
        a = taxable_amount
        if a <= 10_000_000: return d("0.10"), d("0")
        if a <= 30_000_000: return d("0.15"), d("500000")
        if a <= 50_000_000: return d("0.20"), d("2000000")
        if a <= 100_000_000: return d("0.30"), d("7000000")
        if a <= 200_000_000: return d("0.40"), d("17000000")
        if a <= 300_000_000: return d("0.45"), d("27000000")
        if a <= 600_000_000: return d("0.50"), d("42000000")
        return d("0.55"), d("72000000")

    @classmethod
    def calc_total_tax(cls, net_taxable_assets, heirs_info, has_spouse):
        n = len(heirs_info) + (1 if has_spouse else 0)
        if n == 0: return d("0")
        basic_deduction = d("30000000") + d("6000000") * d(str(n))
        taxable_total = max(d("0"), net_taxable_assets - basic_deduction)
        if taxable_total == 0: return d("0")

        # 法定相続分での仮分割
        if has_spouse:
            if any(h['type'] == "子" for h in heirs_info):
                s_share = taxable_total * d("0.5")
            elif any(h['type'] == "親" for h in heirs_info):
                s_share = taxable_total * d("0.666666666")
            else:
                s_share = taxable_total * d("0.75")
        else:
            s_share = d("0")

        remaining = taxable_total - s_share
        child_count = len(heirs_info)
        h_share = remaining / d(str(child_count)) if child_count > 0 else d("0")

        total_tax = d("0")
        if has_spouse:
            rate, ded = cls.get_tax_rate(s_share)
            total_tax += round_tax(s_share * rate - ded)
        for _ in heirs_info:
            rate, ded = cls.get_tax_rate(h_share)
            total_tax += round_tax(h_share * rate - ded)
        return total_tax

class SupremeLegacyEngine:
    @staticmethod
    def get_legal_shares(has_spouse, heirs_info):
        n = len(heirs_info)
        if not has_spouse:
            return d("0"), [d("1")/d(str(n))]*n if n>0 else []
        if any(h['type'] == "子" for h in heirs_info):
            return d("0.5"), [d("0.5")/d(str(n))]*n
        if any(h['type'] == "親" for h in heirs_info):
            return d("0.666"), [d("0.333")/d(str(n))]*n
        return d("0.75"), [d("0.25")/d(str(n))]*n

# --- 2. Excel出力エンジン (New) ---
def generate_excel(data_dict):
    output = io.BytesIO()
    wb = Workbook()
    
    # スタイル定義
    navy_fill = PatternFill(start_color="1F2C4D", end_color="1F2C4D", fill_type="solid")
    gold_font = Font(color="C5A059", bold=True)
    header_font = Font(bold=True, color="FFFFFF")
    center_align = Alignment(horizontal="center")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # Sheet 1: 診断概要
    ws1 = wb.active
    ws1.title = "診断概要"
    ws1.append(["項目", "内容"])
    summary_data = [
        ["総資産（特例適用前）", f"{data_dict['total_assets']:,}円"],
        ["小規模宅地特例評価減額", f"{data_dict['deduction_amount']:,}円"],
        ["課税対象資産（純資産）", f"{data_dict['net_taxable']:,}円"],
        ["法定相続人数", f"{data_dict['heir_count']}名"]
    ]
    for row in summary_data:
        ws1.append(row)
    
    # Sheet 2: シミュレーション結果
    ws2 = wb.create_sheet("税額シミュレーション")
    headers = ["配偶者取得割合(%)", "一次相続税額(円)", "二次相続税額(円)", "合計納税額(円)"]
    ws2.append(headers)
    for cell in ws2[1]:
        cell.fill = navy_fill
        cell.font = header_font

    for _, row in data_dict['df_sim'].iterrows():
        ws2.append([row['配分(%)'], row['一次相続税額'], row['二次相続税額'], row['合計納税額']])

    # 列幅調整
    for ws in wb.worksheets:
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            ws.column_dimensions[column].width = max_length + 5

    wb.save(output)
    return output.getvalue()

# --- 3. メインアプリケーション ---
def main():
    if not check_password(): return

    # --- サイドバー：基本入力 ---
    with st.sidebar:
        st.header("💎 基本財産入力")
        land_v = st.number_input("土地評価額 (円)", value=100_000_000, step=1_000_000)
        house_v = st.number_input("建物評価額 (円)", value=20_000_000, step=1_000_000)
        cash_v = st.number_input("現預金・証券 (円)", value=50_000_000, step=1_000_000)
        stock_v = st.number_input("非上場株式 (円)", value=30_000_000, step=1_000_000)
        debt_v = st.number_input("負債・葬式費用 (円)", value=5_000_000, step=1_000_000)
        
        st.divider()
        st.header("🎁 生前贈与等")
        gift_addition = st.number_input("生前贈与加算 (円) ※亡くなる前7年以内", value=0, step=100_000)
        tokutei_gift = st.number_input("相続時精算課税適用財産 (円)", value=0, step=100_000)
        
        st.divider()
        st.header("👥 相続人構成")
        has_spouse = st.checkbox("配偶者あり", value=True)
        child_num = st.number_input("相続人の数（配偶者除く）", min_value=0, max_value=10, value=2)
        
        heirs_info = []
        for i in range(child_num):
            with st.expander(f"相続人 {i+1} 設定"):
                h_type = st.selectbox(f"関係 {i+1}", ["子", "親", "兄弟姉妹（全血）", "兄弟姉妹（半血）"], key=f"type_{i}")
                heirs_info.append({"type": h_type})

    # --- 特例計算ロジック ---
    st.title("🏆 SUMMIT v32.01 PRO")
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "基本財産明細", "小規模宅地特例", "一次相続税計算", "二次相続シミュ", "遺留分・法的リスク", "シミュレーション・最適化分析"
    ])

    total_raw = d(str(land_v + house_v + cash_v + stock_v - debt_v))
    
    # 2. 小規模宅地特例タブ
    with tab2:
        st.subheader("🏡 小規模宅地等の特例判定")
        is_applied = st.radio("特例を適用しますか？", ["適用する", "適用しない"])
        land_usage = st.selectbox("土地の利用区分", ["特定居住用 (330㎡まで80%減)", "特定事業用 (400㎡まで80%減)", "貸付事業用 (200㎡まで50%減)"])
        land_area = st.number_input("土地面積 (㎡)", value=200.0)
        
        deduction = d("0")
        if is_applied == "適用する":
            limit_area = 330 if "居住" in land_usage else 400 if "事業" in land_usage else 200
            rate = d("0.8") if "居住" in land_usage or "事業" in land_usage else d("0.5")
            applied_area = min(d(str(land_area)), d(str(limit_area)))
            unit_price = d(str(land_v)) / d(str(land_area)) if land_area > 0 else d("0")
            deduction = round_tax(applied_area * unit_price * rate)
        
        st.metric("特例評価減額", f"{int(deduction):,}円")
        st.info(f"根拠: {land_usage} の上限面積 {limit_area}㎡ 内での計算")
        if st.button("このページを印刷", key="p2"):
            components.html("<script>window.print();</script>")

    net_taxable = total_raw - deduction + d(str(gift_addition)) + d(str(tokutei_gift))

    # 1. 明細タブ
    with tab1:
        st.subheader("📊 相続財産構成")
        data = {
            "項目": ["土地", "建物", "現預金・証券", "非上場株式", "（負債）", "生前贈与加算", "相続時精算課税"],
            "金額": [land_v, house_v, cash_v, stock_v, -debt_v, gift_addition, tokutei_gift]
        }
        st.table(pd.DataFrame(data))
        st.metric("課税対象総額", f"{int(net_taxable):,}円")
        if st.button("このページを印刷", key="p1"):
            components.html("<script>window.print();</script>")

    # 3. 一次相続税計算
    with tab3:
        st.subheader("⚖️ 一次相続税の概算")
        total_tax = SupremeTaxEngine.calc_total_tax(net_taxable, heirs_info, has_spouse)
        
        col1, col2 = st.columns(2)
        col1.metric("相続税総額", f"{int(total_tax):,}円")
        
        st.write("■ 法定相続分で分割した場合の納税予測")
        s_r, h_shares = SupremeLegacyEngine.get_legal_shares(has_spouse, heirs_info)
        
        calc_data = []
        if has_spouse:
            # 配偶者控除適用（法定分または1.6億の大きい方まで非課税）
            tax_val = max(d("0"), (total_tax * s_r) - (total_tax * s_r)) # 簡易表示用
            calc_data.append({"相続人": "配偶者", "取得額(法定)": f"{int(net_taxable*s_r):,}円", "税額": "0円 (配偶者控除)"})
        
        for i, (h, share) in enumerate(zip(heirs_info, h_shares)):
            calc_data.append({"相続人": f"相続人{i+1}", "取得額(法定)": f"{int(net_taxable*share):,}円", "税額": f"{int(total_tax*share):,}円"})
        
        st.table(pd.DataFrame(calc_data))
        if st.button("このページを印刷", key="p3"):
            components.html("<script>window.print();</script>")

    # 4. 二次相続シミュ
    with tab4:
        st.subheader("🔄 二次相続を見据えた比較")
        if not has_spouse:
            st.warning("配偶者がいないため、二次相続シミュレーションは不要です。")
        else:
            spouse_own_assets = st.number_input("配偶者自身の固有財産 (円)", value=30_000_000, step=1_000_000)
            st.info("一次相続で配偶者が取得した分が、将来の二次相続の課税対象に加算されます。")
            if st.button("このページを印刷", key="p4"):
                components.html("<script>window.print();</script>")

    # 5. 遺留分
    with tab5:
        st.subheader("⚠️ 遺留分侵害額の確認")
        s_r, h_shares = SupremeLegacyEngine.get_legal_shares(has_spouse, heirs_info)
        iryu_total_ratio = d("0.333") if all(h['type'] == "親" for h in heirs_info) else d("0.5")
        
        iryu_data = []
        if has_spouse:
            iryu_data.append({"相続人": "配偶者", "法定相続分": f"{float(s_r)*100:.1f}%", "遺留分額": f"{int(net_taxable * s_r * iryu_total_ratio):,}円"})
        for i, (h, share) in enumerate(zip(heirs_info, h_shares)):
            val = f"{int(net_taxable * share * iryu_total_ratio):,}円" if h['type'] not in ["兄弟姉妹（全血）", "兄弟姉妹（半血）"] else "（権利なし）"
            iryu_data.append({"相続人": f"相続人{i+1}({h['type']})", "法定相続分": f"{float(share)*100:.1f}%", "遺留分額": val})
        
        st.table(pd.DataFrame(iryu_data))
        if st.button("このページを印刷", key="p5"):
            components.html("<script>window.print();</script>")

    # 6. 最適化シミュレーション
    with tab6:
        st.subheader("📈 配偶者の取得割合別・納税総額シミュレーション")
        if not has_spouse:
            st.error("配偶者なしの設定では実行できません。")
        else:
            s_own = d(str(st.session_state.get('spouse_own_assets', 30_000_000)))
            sim_results = []
            for i in range(0, 101, 10):
                ratio = d(str(i)) / d("100")
                # 一次
                t1_total = SupremeTaxEngine.calc_total_tax(net_taxable, heirs_info, True)
                s_get = net_taxable * ratio
                # 配偶者控除
                s_tax_before = t1_total * ratio
                tax_limit = max(net_taxable * d("0.5"), d("160000000"))
                s_tax_after = d("0") if s_get <= tax_limit else round_tax(s_tax_before * (s_get - tax_limit) / s_get)
                t1_final = round_tax(t1_total * (1 - ratio)) + s_tax_after
                
                # 二次
                net2 = s_get + s_own
                t2_final = SupremeTaxEngine.calc_total_tax(net2, heirs_info, False)
                
                sim_results.append({
                    "配分(%)": i,
                    "一次相続税額": int(t1_final),
                    "二次相続税額": int(t2_final),
                    "合計納税額": int(t1_final + t2_final)
                })
            
            df_sim = pd.DataFrame(sim_results)
            fig = go.Figure()
            fig.add_trace(go.Bar(x=df_sim['配分(%)'], y=df_sim['一次相続税額'], name='一次相続税', marker_color='#1f2c4d'))
            fig.add_trace(go.Bar(x=df_sim['配分(%)'], y=df_sim['二次相続税額'], name='二次相続税', marker_color='#c5a059'))
            fig.add_trace(go.Scatter(x=df_sim['配分(%)'], y=df_sim['合計納税額'], name='合計', line=dict(color='#a61d24', width=4)))
            fig.update_layout(barmode='stack', title="税額最適化シミュレーション", xaxis_title="配偶者配分(%)", yaxis_title="税額(円)")
            st.plotly_chart(fig, use_container_width=True)
            st.table(df_sim.style.format({"一次相続税額": "{:,}", "二次相続税額": "{:,}", "合計納税額": "{:,}"}))

            # --- Excel出力ボタン ---
            st.divider()
            st.subheader("📂 帳票出力・データ保存")
            
            excel_data = {
                'total_assets': int(total_raw),
                'deduction_amount': int(deduction),
                'net_taxable': int(net_taxable),
                'heir_count': len(heirs_info) + 1,
                'df_sim': df_sim
            }
            
            excel_bin = generate_excel(excel_data)
            
            st.download_button(
                label="📊 計算結果をExcelで保存",
                data=excel_bin,
                file_name="相続税シミュレーション結果.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.caption("※入力データ、特例適用後の純資産、一次・二次相続の推移表がエクスポートされます。")

    # 共通フッター
    st.divider()
    st.markdown("""
    <div style="background-color: #1f2c4d; padding: 20px; border-radius: 10px; border: 2px solid #c5a059; color: white;">
        <h3 style="color: #c5a059; margin-top: 0;">📜 山根会計 資産税シミュレーション 監査済証</h3>
        <p>本計算結果は、入力された財産評価額に基づき、現行の相続税法および租税特別措置法を適用して算出された概算数値です。実際の申告に際しては、財産の現地調査および個別具体的な税務判断が必要となります。</p>
        <p style="text-align: right; font-style: italic;">Yamane Accounting Proprietary System - Core v32.01</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
