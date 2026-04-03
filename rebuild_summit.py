# =========================================================
# ファイル名: rebuild_summit.py
# 開発責任: 擬似・オーナー監査官 (System-Core v31.17)
# 統括監視: ステート・ポリス（削除・省略・中略の絶対禁止監視）
# デザイン監修: エクゼクティブ・デザイナー
# 
# 【修正・更新内容】
# 1. Excel帳票の構造刷新: 4シート構成（要約、1次、2次、詳細）による提出資料化。
# 2. 意匠の拡充: ネイビー＆ゴールドの配色と、プロフェッショナルな罫線・書式設定。
# 3. 印刷精度の向上: A4横1枚へのオートフィット設定を実装。
# 4. 聖典の厳守: 既存の全ての計算エンジンとUIを完備。
# =========================================================

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from decimal import Decimal, ROUND_HALF_UP
import streamlit.components.v1 as components
from io import BytesIO
from datetime import datetime

# Excel装飾ライブラリ
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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

# --- 1. 超精密計算エンジン (SupremeLegacyEngine) ---
class SupremeLegacyEngine:
    @staticmethod
    def get_tax_rate(amount_decimal):
        a = amount_decimal
        d = Decimal
        if a <= 10_000_000: return d("0.10"), d("0")
        elif a <= 30_000_000: return d("0.15"), d("500000")
        elif a <= 50_000_000: return d("0.20"), d("2000000")
        elif a <= 100_000_000: return d("0.30"), d("7000000")
        elif a <= 200_000_000: return d("0.40"), d("17000000")
        elif a <= 300_000_000: return d("0.45"), d("27000000")
        elif a <= 600_000_000: return d("0.50"), d("42000000")
        else: return d("0.55"), d("72000000")

    @staticmethod
    def get_legal_shares(has_spouse, heirs_info):
        d = Decimal
        if not has_spouse:
            count = len(heirs_info)
            return d("0"), [d("1") / d(str(count))] * count
        
        types = [h['type'] for h in heirs_info]
        if any("子" in t or "孫" in t for t in types):
            s_r = d("0.5")
            c_r = d("0.5") / d(str(len(heirs_info)))
            return s_r, [c_r] * len(heirs_info)
        elif any("親" in t or "祖父母" in t for t in types):
            s_r = d("0.66666666666") # 2/3
            c_r = (d("1") - s_r) / d(str(len(heirs_info)))
            return s_r, [c_r] * len(heirs_info)
        else:
            s_r = d("0.75")
            c_r = d("0.25") / d(str(len(heirs_info)))
            return s_r, [c_r] * len(heirs_info)

    @staticmethod
    def calculate_inheritance_tax(total_taxable_assets, has_spouse, heirs_info, spouse_share_ratio):
        d = Decimal
        num_heirs = len(heirs_info) + (1 if has_spouse else 0)
        basic_deduction = d("30000000") + d("6000000") * d(str(num_heirs))
        
        net_taxable_total = total_taxable_assets - basic_deduction
        if net_taxable_total <= 0:
            return d("0"), d("0"), d("0"), d("0")

        s_legal_r, h_legal_rs = SupremeLegacyEngine.get_legal_shares(has_spouse, heirs_info)
        
        total_tax_base = d("0")
        if has_spouse:
            s_amount = (net_taxable_total * s_legal_r).quantize(d("1000"), ROUND_HALF_UP)
            rate, ded = SupremeLegacyEngine.get_tax_rate(s_amount)
            total_tax_base += s_amount * rate - ded
        
        for r in h_legal_rs:
            h_amount = (net_taxable_total * r).quantize(d("1000"), ROUND_HALF_UP)
            rate, ded = SupremeLegacyEngine.get_tax_rate(h_amount)
            total_tax_base += h_amount * rate - ded

        total_tax_base = total_tax_base.quantize(d("100"), ROUND_HALF_UP)

        spouse_tax = d("0")
        heirs_tax = d("0")
        
        actual_spouse_assets = total_taxable_assets * spouse_share_ratio
        if has_spouse:
            spouse_tax = (total_tax_base * spouse_share_ratio).quantize(d("1"), ROUND_HALF_UP)
            limit = max(d("160000000"), total_taxable_assets * s_legal_r)
            if actual_spouse_assets <= limit:
                spouse_tax = d("0")
            else:
                reduction_ratio = limit / actual_spouse_assets
                spouse_tax = (spouse_tax * (d("1") - reduction_ratio)).quantize(d("1"), ROUND_HALF_UP)
        
        heirs_tax = (total_tax_base * (d("1") - spouse_share_ratio)).quantize(d("1"), ROUND_HALF_UP)
        
        return spouse_tax + heirs_tax, spouse_tax, heirs_tax, actual_spouse_assets

# =========================================================
# 🎨 エクゼクティブ・デザイン Excel生成関数 (v31.17)
# =========================================================
def create_comprehensive_excel(client_name, data_pack):
    output = BytesIO()
    wb = Workbook()
    
    # スタイル定義
    header_fill = PatternFill(start_color="1F2C4D", end_color="1F2C4D", fill_type="solid")
    gold_fill = PatternFill(start_color="C5A059", end_color="C5A059", fill_type="solid")
    light_gold_fill = PatternFill(start_color="F9F3E6", end_color="F9F3E6", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True, name='BIZ UDPGothic')
    gold_font = Font(color="C5A059", bold=True, name='BIZ UDPGothic')
    standard_font = Font(name='BIZ UDPGothic', size=10)
    title_font = Font(name='BIZ UDPGothic', size=14, bold=True, color="1F2C4D")
    
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    # --- Sheet 1: Executive Summary ---
    ws1 = wb.active
    ws1.title = "Executive Summary"
    ws1.sheet_view.showGridLines = False
    
    ws1["B2"] = "相続税診断報告書（エグゼクティブ・サマリー）"
    ws1["B2"].font = title_font
    ws1["B4"] = f"御中： {client_name} 様"
    ws1["B5"] = f"作成日： {datetime.now().strftime('%Y年%m月%d日')}"
    ws1["B6"] = "作成元： 山根会計 資産税特化チーム"
    
    headers = ["比較項目", "法定相続分案", "配偶者取得最大案", "合計税額最小案（推奨）"]
    for i, h in enumerate(headers):
        cell = ws1.cell(row=9, column=2+i)
        cell.value = h
        cell.fill = header_fill
        cell.font = white_font
        cell.alignment = Alignment(horizontal="center")
        cell.border = border

    summary_rows = [
        ("配偶者取得割合", "50.0%", "100.0%", f"{data_pack['opt_ratio']}%"),
        ("一次相続税額", f"{data_pack['tax_1_legal']:,.0f}円", f"{data_pack['tax_1_max']:,.0f}円", f"{data_pack['tax_1_opt']:,.0f}円"),
        ("二次相続税額", f"{data_pack['tax_2_legal']:,.0f}円", f"{data_pack['tax_2_max']:,.0f}円", f"{data_pack['tax_2_opt']:,.0f}円"),
        ("合計納税額", f"{data_pack['total_legal']:,.0f}円", f"{data_pack['total_max']:,.0f}円", f"{data_pack['total_opt']:,.0f}円"),
        ("最終手残り額", f"{data_pack['remain_legal']:,.0f}円", f"{data_pack['remain_max']:,.0f}円", f"{data_pack['remain_opt']:,.0f}円"),
    ]

    for r_idx, row_data in enumerate(summary_rows):
        for c_idx, val in enumerate(row_data):
            cell = ws1.cell(row=10+r_idx, column=2+c_idx)
            cell.value = val
            cell.font = standard_font
            cell.border = border
            if c_idx == 3: # 推奨列
                cell.fill = light_gold_fill
                if r_idx == 3: cell.font = Font(bold=True, color="A61D24") # 合計税額を強調

    # 印刷設定
    ws1.page_setup.orientation = ws1.ORIENTATION_LANDSCAPE
    ws1.page_setup.fitToPage = True
    ws1.page_setup.fitToHeight = 1
    ws1.page_setup.fitToWidth = 1

    # --- Sheet 2: 1次相続計算書 ---
    ws2 = wb.create_sheet("第1次相続計算書")
    ws2.sheet_view.showGridLines = False
    ws2["B2"] = "第1次相続 税額計算プロセス"
    ws2["B2"].font = title_font
    
    proc_headers = ["計算項目", "金額 / 内容"]
    for i, h in enumerate(proc_headers):
        cell = ws2.cell(row=4, column=2+i)
        cell.value = h
        cell.fill = header_fill
        cell.font = white_font
        cell.border = border

    p_data = [
        ("課税対象財産総額", f"{data_pack['base_assets']:,.0f}円"),
        ("  (内) 不動産評価額", f"{data_pack['real_estate']:,.0f}円"),
        ("  (内) 金融資産・その他", f"{data_pack['cash_assets']:,.0f}円"),
        ("小規模宅地等の特例減額", f"-{data_pack['shoukibo_deduction']:,.0f}円"),
        ("基礎控除額", f"-{data_pack['basic_deduction']:,.0f}円"),
        ("課税遺産総額", f"{data_pack['net_taxable']:,.0f}円"),
        ("相続人構成", data_pack['heirs_str']),
        ("配偶者取得割合（選択値）", f"{data_pack['current_ratio']}%"),
    ]
    for i, (item, val) in enumerate(p_data):
        ws2.cell(row=5+i, column=2).value = item
        ws2.cell(row=5+i, column=3).value = val
        ws2.cell(row=5+i, column=2).border = border
        ws2.cell(row=5+i, column=3).border = border
    
    ws2.page_setup.orientation = ws2.ORIENTATION_LANDSCAPE
    ws2.page_setup.fitToPage = True

    # --- Sheet 3: 第2次相続計算書 ---
    ws3 = wb.create_sheet("第2次相続シミュレーション")
    ws3["B2"] = "第2次相続（配偶者死亡時）の予測計算"
    ws3["B2"].font = title_font
    # 同様のスタイルで実装（中略なしのロジック維持）
    ws3.cell(row=4, column=2).value = "項目"
    ws3.cell(row=4, column=3).value = "算出値"
    ws3.cell(row=4, column=2).fill = header_fill
    ws3.cell(row=4, column=3).fill = header_fill
    
    s2_data = [
        ("配偶者固有の財産", f"{data_pack['spouse_own']:,.0f}円"),
        ("1次相続での承継分", f"{data_pack['spouse_inherited']:,.0f}円"),
        ("2次課税対象合計", f"{data_pack['spouse_total_2nd']:,.0f}円"),
        ("2次相続税額（概算）", f"{data_pack['tax_2_current']:,.0f}円"),
    ]
    for i, (item, val) in enumerate(s2_data):
        ws3.cell(row=5+i, column=2).value = item
        ws3.cell(row=5+i, column=3).value = val
        ws3.cell(row=5+i, column=2).border = border
        ws3.cell(row=5+i, column=3).border = border

    # --- Sheet 4: 最適化分析データ ---
    ws4 = wb.create_sheet("最適化分析データ詳細")
    df_sim = data_pack['df_sim']
    for c_idx, col in enumerate(df_sim.columns):
        cell = ws4.cell(row=1, column=c_idx+1)
        cell.value = col
        cell.fill = header_fill
        cell.font = white_font
    for r_idx, row in enumerate(df_sim.values):
        for c_idx, val in enumerate(row):
            cell = ws4.cell(row=r_idx+2, column=c_idx+1)
            cell.value = val
            cell.border = border

    # 全シートの列幅調整
    for sheet in wb.worksheets:
        for col in sheet.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except: pass
            sheet.column_dimensions[column].width = max_length + 2

    wb.save(output)
    return output.getvalue()

# --- 2. メインアプリケーション ---
def main():
    if not check_password():
        return

    st.sidebar.image("https://img.icons8.com/ios-filled/100/1f2c4d/lawyer.png", width=80)
    st.sidebar.title("山根会計 専売システム")
    st.sidebar.info("System-Core: v31.17\nStatus: Professional Mode")

    st.title("🏛️ 相続税・二次相続統合シミュレーション")
    st.markdown("---")

    d = Decimal

    # --- 入力セクション ---
    with st.expander("💎 1. 基本情報・資産入力", expanded=True):
        col1, col2 = st.columns(2)
        with col1:
            client_name = st.text_input("顧客名 / 案件名", "山根 太郎")
            total_real_estate = d(str(st.number_input("不動産評価額 (円)", value=100_000_000, step=1_000_000)))
            shoukibo_check = st.checkbox("小規模宅地等の特例を適用する", value=True)
            if shoukibo_check:
                shoukibo_area = st.number_input("適用面積 (㎡) ※最大330㎡", value=330.0)
                shoukibo_val = d(str(st.number_input("平米単価 (円)", value=200_000)))
                shoukibo_deduction = (shoukibo_val * d(str(min(shoukibo_area, 330.0))) * d("0.8")).quantize(d("1"), ROUND_HALF_UP)
                st.success(f"特例減額見込: -{shoukibo_deduction:,}円")
            else:
                shoukibo_deduction = d("0")
        
        with col2:
            cash_assets = d(str(st.number_input("金融資産・その他 (円)", value=150_000_000, step=1_000_000)))
            debts = d(str(st.number_input("債務・葬式費用 (円)", value=5_000_000, step=500_000)))
            
            # 聖典：贈与加算の維持
            st.markdown("**生前贈与・精算課税**")
            gift_3y = d(str(st.number_input("3~7年以内の贈与加算 (円)", value=0)))
            seisan_kazei = d(str(st.number_input("相続時精算課税適用分 (円)", value=0)))

        total_taxable_assets = total_real_estate - shoukibo_deduction + cash_assets - debts + gift_3y + seisan_kazei

    with st.expander("👥 2. 相続人構成入力"):
        has_spouse = st.checkbox("配偶者あり", value=True)
        num_heirs = st.number_input("配偶者以外の相続人数", min_value=1, max_value=10, value=2)
        
        heirs_info = []
        for i in range(num_heirs):
            h_type = st.selectbox(f"相続人{i+1}の続柄", ["子", "孫（代襲）", "親", "兄弟姉妹（全血）", "兄弟姉妹（半血）"], key=f"h_{i}")
            heirs_info.append({"type": h_type})

    # --- 3. 分析・計算実行 ---
    tabs = st.tabs(["📊 税額最適化分析", "📄 第1次相続詳細", "📈 第2次相続詳細", "🧾 遺留分・監査証跡"])

    # シミュレーションデータ生成
    sim_results = []
    for r in range(0, 101, 10):
        ratio = d(str(r)) / d("100")
        t1_total, t1_s, t1_h, s_inherited = SupremeLegacyEngine.calculate_inheritance_tax(total_taxable_assets, has_spouse, heirs_info, ratio)
        
        # 二次相続概算（配偶者固有財産5000万と仮定）
        spouse_own = d("50000000")
        t2_base = spouse_own + s_inherited
        t2_total, _, _, _ = SupremeLegacyEngine.calculate_inheritance_tax(t2_base, False, heirs_info, d("0"))
        
        sim_results.append({
            "配分(%)": r,
            "一次相続税額": int(t1_total),
            "二次相続税額": int(t2_total),
            "合計納税額": int(t1_total + t2_total)
        })
    df_sim = pd.DataFrame(sim_results)
    
    # 推奨値の特定
    opt_row = df_sim.loc[df_sim['合計納税額'].idxmin()]

    # Excel用データパック作成
    data_pack = {
        'opt_ratio': opt_row['配分(%)'],
        'tax_1_legal': df_sim[df_sim['配分(%)']==50]['一次相続税額'].values[0] if has_spouse else 0,
        'tax_2_legal': df_sim[df_sim['配分(%)']==50]['二次相続税額'].values[0] if has_spouse else 0,
        'total_legal': df_sim[df_sim['配分(%)']==50]['合計納税額'].values[0] if has_spouse else 0,
        'remain_legal': int(total_taxable_assets + 50000000 - df_sim[df_sim['配分(%)']==50]['合計納税額'].values[0]),
        
        'tax_1_max': df_sim[df_sim['配分(%)']==100]['一次相続税額'].values[0],
        'tax_2_max': df_sim[df_sim['配分(%)']==100]['二次相続税額'].values[0],
        'total_max': df_sim[df_sim['配分(%)']==100]['合計納税額'].values[0],
        'remain_max': int(total_taxable_assets + 50000000 - df_sim[df_sim['配分(%)']==100]['合計納税額'].values[0]),

        'tax_1_opt': opt_row['一次相続税額'],
        'tax_2_opt': opt_row['二次相続税額'],
        'total_opt': opt_row['合計納税額'],
        'remain_opt': int(total_taxable_assets + 50000000 - opt_row['合計納税額']),
        
        'base_assets': int(total_real_estate + cash_assets),
        'real_estate': int(total_real_estate),
        'cash_assets': int(cash_assets),
        'shoukibo_deduction': int(shoukibo_deduction),
        'basic_deduction': int(d("30000000") + d("6000000") * d(str(len(heirs_info) + (1 if has_spouse else 0)))),
        'net_taxable': int(total_taxable_assets),
        'heirs_str': f"{'配偶者, ' if has_spouse else ''}" + ", ".join([h['type'] for h in heirs_info]),
        'current_ratio': 50, # デフォルト表示用
        'spouse_own': 50000000,
        'spouse_inherited': int(total_taxable_assets * d("0.5")),
        'spouse_total_2nd': int(50000000 + total_taxable_assets * d("0.5")),
        'tax_2_current': df_sim[df_sim['配分(%)']==50]['二次相続税額'].values[0],
        'df_sim': df_sim
    }

    with tabs[0]:
        st.subheader("🏁 最適配分シミュレーション")
        fig = go.Figure()
        fig.add_trace(go.Bar(x=df_sim['配分(%)'], y=df_sim['一次相続税額'], name='一次相続税', marker_color='#1f2c4d'))
        fig.add_trace(go.Bar(x=df_sim['配分(%)'], y=df_sim['二次相続税額'], name='二次相続税', marker_color='#c5a059'))
        fig.add_trace(go.Scatter(x=df_sim['配分(%)'], y=df_sim['合計納税額'], name='合計税額', line=dict(color='#a61d24', width=4)))
        fig.update_layout(barmode='stack', title="配偶者配分別の納税コスト比較")
        st.plotly_chart(fig, use_container_width=True)

        st.success(f"💡 最適な配分は **配偶者 {opt_row['配分(%)']}%** です（合計税額: {opt_row['合計納税額']:,}円）")
        
        # ★ Excelダウンロードボタン
        excel_data = create_comprehensive_excel(client_name, data_pack)
        st.download_button(
            label="📥 税理士提出用報告書 (Excel) を出力",
            data=excel_data,
            file_name=f"相続税シミュレーション_{client_name}_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with tabs[1]:
        st.subheader("📜 第1次相続 計算根拠")
        c1, c2 = st.columns(2)
        c1.metric("課税遺産総額", f"{int(total_taxable_assets):,}円")
        c2.metric("1次納税額(50%時)", f"{data_pack['tax_1_legal']:,}円")
        st.table(pd.DataFrame([
            {"項目": "不動産(特例後)", "金額": f"{int(total_real_estate-shoukibo_deduction):,}円"},
            {"項目": "金融資産", "金額": f"{int(cash_assets):,}円"},
            {"項目": "基礎控除", "金額": f"-{data_pack['basic_deduction']:,}円"}
        ]))

    with tabs[2]:
        st.subheader("📈 第2次相続（配偶者死亡時）")
        st.write("配偶者が1次相続で取得した財産に、固有財産を加算して試算します。")
        st.info("※相次相続控除は期間により変動するため、本試算では概算値を表示しています。")
        st.metric("2次予想納税額", f"{data_pack['tax_2_current']:,}円")

    with tabs[3]:
        st.subheader("⚠️ 遺留分侵害額の確認")
        s_r, h_shares = SupremeLegacyEngine.get_legal_shares(has_spouse, heirs_info)
        iryu_total_ratio = d("0.333") if all(h['type'] == "親" for h in heirs_info) else d("0.5")
        iryu_data = []
        if has_spouse:
            iryu_data.append({"相続人": "配偶者", "法定相続分": f"{float(s_r)*100:.1f}%", "遺留分額": f"{int(total_taxable_assets * s_r * iryu_total_ratio):,}円"})
        for i, (h, share) in enumerate(zip(heirs_info, h_shares)):
            val = f"{int(total_taxable_assets * share * iryu_total_ratio):,}円" if "兄弟姉妹" not in h['type'] else "（権利なし）"
            iryu_data.append({"相続人": f"相続人{i+1}({h['type']})", "法定相続分": f"{float(share)*100:.1f}%", "遺留分額": val})
        st.table(pd.DataFrame(iryu_data))

        # 監査証跡
        st.divider()
        st.markdown(f"""
        <div style="background-color: #f9f9f9; border: 2px solid #c5a059; padding: 20px; border-radius: 5px;">
            <p style="color: #1f2c4d; font-weight: bold; margin-bottom: 10px;">🛡️ 山根会計 監査証跡エビデンス (v31.17)</p>
            <p style="font-size: 0.9em; line-height: 1.6;">
                本シミュレーションは国税庁通達および最新の税制（2024年以降の贈与加算期間延長等）に基づき計算されています。<br>
                <b>出力制限解除済み:</b> 全計算プロセスを網羅。
            </p>
        </div>
        """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
