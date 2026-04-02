# =========================================================
# ファイル名: rebuild_summit.py
# 開発責任: 擬似・オーナー監査官 (System-Core v29.0)
# 【緊急復元】小規模宅地特例、二次相続・遺留分計算根拠を完全復旧。
# 【真・聖典準拠】中略・省略を一切排除したフルスタック・コード。
# =========================================================

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from decimal import Decimal, ROUND_HALF_UP

# --- 0. セキュリティ設定 ---
def check_password():
    if "password_correct" not in st.session_state:
        st.set_page_config(page_title="山根会計 専売システム", layout="wide")
        st.title("🔐 山根会計 専売システム")
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

# --- 1. 超精密計算エンジン ---
class SupremeLegacyEngine:
    @staticmethod
    def to_d(val): return Decimal(str(val))

    @staticmethod
    def get_tax(taxable_amt, count):
        d = SupremeLegacyEngine.to_d
        if taxable_amt <= 0: return d(0)
        # 聖典第四則：分母ルール
        per_heir = (taxable_amt / d(count)).quantize(d(1), ROUND_HALF_UP)
        def bracket(a):
            if a <= 10000000: return a * d("0.10")
            elif a <= 30000000: return a * d("0.15") - d("500000")
            elif a <= 50000000: return a * d("0.20") - d("2000000")
            elif a <= 100000000: return a * d("0.30") - d("7000000")
            elif a <= 200000000: return a * d("0.40") - d("17000000")
            elif a <= 300000000: return a * d("0.45") - d("27000000")
            elif a <= 600000000: return a * d("0.50") - d("42000000")
            else: return a * d("0.55") - d("72000000")
        return (bracket(per_heir) * d(count)).quantize(d(1), ROUND_HALF_UP)

# --- 2. メインUI ---
if check_password():
    st.set_page_config(page_title="SUMMIT v29.0 PRO", layout="wide")
    st.sidebar.markdown("### 🏢 山根会計 専売システム")
    st.sidebar.info("ログインユーザー: 川東")
    
    if st.sidebar.button("入力内容をリセット"):
        for key in list(st.session_state.keys()):
            if key != "password_correct":
                del st.session_state[key]
        st.rerun()

    st.title("🏛️ Summit System-Core v29.0")
    
    # 聖典第二則：6タブ構成
    tabs = st.tabs([
        "👥 1.基本構成", 
        "💰 2.一次財産詳細", 
        "📑 3.一次相続明細", 
        "📑 4.二次相続明細", 
        "⏳ 5.二次推移予測", 
        "📊 6.精密分析結果"
    ])

    d = SupremeLegacyEngine.to_d

    # -- TAB 1: 基本構成 --
    with tabs[0]:
        st.header("相続関係の設定")
        c1, c2 = st.columns(2)
        child_count = c1.number_input("子供の人数", 1, 10, 2, key="in_child")
        has_spouse = c2.checkbox("配偶者は健在", value=True, key="in_spouse")
        st.subheader("相続人の属性設定")
        for i in range(child_count):
            st.selectbox(f"子供 {i+1} の続柄", ["子", "孫（養子含む）", "その他"], key=f"rel_{i}")

    # -- TAB 2: 一次詳細（特例入力復元） --
    with tabs[1]:
        st.header("一次相続：財産入力")
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            st.subheader("🏗️ 不動産（小規模宅地特例対応）")
            v_home = st.number_input("特定居住用：評価額", value=32781936, key="v_home")
            a_home = st.number_input("特定居住用：面積(㎡)", value=330, key="a_home")
            v_biz = st.number_input("特定事業用：評価額", value=0, key="v_biz")
            a_biz = st.number_input("特定事業用：面積(㎡)", value=400, key="a_biz")
            v_rent = st.number_input("貸付事業用：評価額", value=0, key="v_rent")
            a_rent = st.number_input("貸付事業用：面積(㎡)", value=200, key="a_rent")
            v_build = st.number_input("建物（評価額）", value=1700044, key="v_build")
            v_land_others = st.number_input("その他の土地", value=0, key="v_land_others")
        with col_b:
            st.subheader("💵 金融・有価証券")
            v_stock = st.number_input("有価証券", value=45132788, key="v_stock")
            v_cash = st.number_input("現預金", value=45573502, key="v_cash")
            v_ins = st.number_input("生命保険金受取額", value=3651514, key="v_ins")
            v_others = st.number_input("その他財産", value=1662687, key="v_others")
        with col_c:
            st.subheader("📉 債務・葬式")
            v_debt = st.number_input("債務合計", value=322179, key="v_debt")
            v_funeral = st.number_input("葬式費用", value=41401, key="v_funeral")

    # -- 共通計算ブロック（小規模宅地ロジック復元） --
    st_count = child_count + (1 if has_spouse else 0)
    
    # 小規模宅地等の特例：併用判定ロジック
    a_lim = d(330); b_lim = d(400); c_lim = d(200)
    a_app = min(d(a_home), a_lim)
    b_app = min(d(a_biz), b_lim)
    # 併用調整（居住用・事業用を優先的に適用し、残枠を貸付用へ）
    used_ratio = (a_app / a_lim) + (b_app / b_lim)
    rem_c_area = max(d(0), c_lim * (d(1) - used_ratio))
    c_app = min(d(a_rent), rem_c_area)
    
    # 減額額算出
    red_home = (d(v_home) / d(max(1, a_home))) * a_app * d("0.8")
    red_biz = (d(v_biz) / d(max(1, a_biz))) * b_app * d("0.8")
    red_rent = (d(v_rent) / d(max(1, a_rent))) * c_app * d("0.5")
    total_reduction = red_home + red_biz + red_rent
    
    land_eval_final = d(v_home) + d(v_biz) + d(v_rent) + d(v_land_others) - total_reduction
    ins_deduct = min(d(v_ins), d(5000000) * d(st_count))
    total_assets = land_eval_final + d(v_build) + d(v_stock) + d(v_cash) + d(v_ins) + d(v_others)
    tax_price = total_assets - ins_deduct - d(v_debt) - d(v_funeral)
    basic_ded_1 = d(30000000) + (d(6000000) * d(st_count))
    taxable_estate = max(d(0), tax_price - basic_ded_1)
    total_tax_1 = SupremeLegacyEngine.get_tax(taxable_estate, st_count)

    # -- TAB 3: 一次相続明細 --
    with tabs[2]:
        st.header("一次相続：計算明細")
        detail_data = [
            ["1", "土地（特例適用後）", f"{int(land_eval_final):,}", "小規模宅地減額反映済"],
            ["2", "建物", f"{int(v_build):,}", ""],
            ["3", "有価証券", f"{int(v_stock):,}", ""],
            ["4", "現預金", f"{int(v_cash):,}", ""],
            ["5", "生命保険金", f"{int(v_ins):,}", ""],
            ["6", "その他財産", f"{int(v_others):,}", ""],
            ["7", "積極財産合計", f"{int(total_assets):,}", ""],
            ["8", "生命保険非課税額", f"△{int(ins_deduct):,}", ""],
            ["9", "債務", f"△{int(v_debt):,}", ""],
            ["10", "葬式費用", f"△{int(v_funeral):,}", ""],
            ["11", "課税価格", f"{int(tax_price):,}", ""],
            ["12", "基礎控除額", f"△{int(basic_ded_1):,}", ""],
            ["13", "課税遺産総額", f"{int(taxable_estate):,}", ""],
            ["14", "相続税の総額", f"{int(total_tax_1):,}", ""],
        ]
        st.table(pd.DataFrame(detail_data, columns=["No", "項目", "金額", "備考"]))

    # -- TAB 4: 二次相続明細 --
    with tabs[3]:
        st.header("二次相続：計算明細予測")
        share_s = d("0.5") # デフォルト50%
        s_acq_base = tax_price * share_s
        s_limit = max(d(160000000), taxable_estate * share_s)
        s_tax_1 = d(0) if s_acq_base <= s_limit else (total_tax_1 * share_s * d("0.5"))
        s_acq_net = s_acq_base - s_tax_1
        s_own = d(st.session_state.get("in_s_own", 50000000))
        s_spend_total = d(st.session_state.get("in_s_spend", 5000000)) * d(st.session_state.get("in_interval", 10))
        s_debt_2 = d(st.session_state.get("in_debt2", 5000000))
        tax_price_2 = max(d(0), s_acq_net + s_own - s_spend_total - s_debt_2)
        basic_ded_2 = d(30000000) + (d(6000000) * d(child_count))
        total_tax_2 = SupremeLegacyEngine.get_tax(max(d(0), tax_price_2 - basic_ded_2), child_count)

        detail_data_2 = [
            ["1", "一次承継分（税引後）", f"{int(s_acq_net):,}", ""],
            ["2", "配偶者固有財産", f"{int(s_own):,}", ""],
            ["3", "生活費消費合計", f"△{int(s_spend_total):,}", ""],
            ["4", "二次債務・葬式費用", f"△{int(s_debt_2):,}", ""],
            ["5", "二次課税価格", f"{int(tax_price_2):,}", ""],
            ["6", "二次基礎控除", f"△{int(basic_ded_2):,}", ""],
            ["7", "二次税額総額", f"{int(total_tax_2):,}", ""],
        ]
        st.table(pd.DataFrame(detail_data_2, columns=["No", "項目", "金額", "備考"]))

    # -- TAB 5: 二次推移予測 --
    with tabs[4]:
        st.header("二次推移パラメータ")
        col_s1, col_s2 = st.columns(2)
        spouse_own_assets = col_s1.number_input("配偶者の固有財産", value=50000000, key="in_s_own")
        life_expectancy = col_s1.slider("想定期間(年)", 0, 20, 10, key="in_interval")
        spouse_spending = col_s2.number_input("年間生活費", value=5000000, key="in_s_spend")
        debt_2nd = col_s2.number_input("二次：債務・葬式", value=5000000, key="in_debt2")

    # -- TAB 6: 分析結果 --
    with tabs[5]:
        st.header("納税コスト最適化分析")
        def run_full_logic(s_ratio):
            ratio = d(s_ratio) / d(100)
            acq_s = tax_price * ratio
            limit_s = max(d(160000000), taxable_estate * ratio)
            tax_s_1 = d(0) if acq_s <= limit_s else (total_tax_1 * ratio * d("0.5"))
            tax_c_1 = d(0)
            ratio_c = (d(1) - ratio) / d(child_count)
            for i in range(child_count):
                rel = st.session_state.get(f"rel_{i}", "子")
                raw = total_tax_1 * ratio_c
                tax_c_1 += (raw * d("1.2")) if rel != "子" else raw
            net_s_2 = max(d(0), acq_s - tax_s_1 + d(spouse_own_assets) - (d(spouse_spending)*d(life_expectancy)) - d(debt_2nd))
            tax_2 = SupremeLegacyEngine.get_tax(max(d(0), net_s_2 - (d(30000000) + d(6000000)*d(child_count))), child_count)
            iru_base = tax_price + ins_deduct
            iru = (iru_base * d("0.5") * (d("0.5") if has_spouse else d("1.0"))) / d(child_count)
            return int(tax_s_1 + tax_c_1), int(tax_2), int(iru)

        res_list = []
        for r in range(0, 101, 10):
            t1, t2, iru = run_full_logic(r)
            res_list.append({"配分(%)": r, "一次税額": t1, "二次税額": t2, "合計": t1+t2, "子1人遺留分": iru})
        df = pd.DataFrame(res_list)
        st.plotly_chart(go.Figure(data=[
            go.Bar(x=df['配分(%)'], y=df['一次税額'], name="一次税額", marker_color='#1f2c4d'),
            go.Bar(x=df['配分(%)'], y=df['二次税額'], name="二次税額", marker_color='#c5a059'),
            go.Scatter(x=df['配分(%)'], y=df['合計'], name="合計", line=dict(color='#a61d24', width=4))
        ], layout=go.Layout(barmode='stack')), use_container_width=True)
        st.table(df.style.format({k: "{:,}円" for k in df.columns if k != "配分(%)"}))

    # --- エビデンス・パネル（聖典第五則：完全復元版） ---
    st.divider()
    st.markdown("""
    <style>
    .logic-box { background-color: #001f3f; color: #d4af37; padding: 25px; border-radius: 10px; border: 2px solid #d4af37; font-family: 'serif'; }
    .formula-card { background-color: rgba(255, 255, 255, 0.05); padding: 15px; border-radius: 8px; margin: 15px 0; border-left: 4px solid #d4af37; }
    .formula-text { font-family: 'Hiragino Mincho ProN', serif; color: #ffffff; font-size: 1.1em; line-height: 1.8; }
    </style>
    <div class="logic-box">
        <h3 style="text-align: center; border-bottom: 1px solid #d4af37; padding-bottom: 10px; margin-bottom: 20px;">🏛️ 相続税精密シミュレーション：計算根拠</h3>
        <div class="formula-card">
            <strong style="color: #d4af37;">1. 一次相続税額および小規模宅地特例</strong><br>
            <div class="formula-text">特定居住用(80%)、特定事業用(80%)、貸付事業用(50%)の各減額割合に基づき、限度面積までの併用調整を精密に行っています。算出された課税価格に対し、法定相続分按分による累進税率を適用しています。</div>
        </div>
        <div class="formula-card">
            <strong style="color: #d4af37;">2. 二次相続税額の推計（復元）</strong><br>
            <div class="formula-text">一次相続での配偶者取得分（税引後）を起点とし、配偶者固有の財産、想定される将来の生活費消費、および二次相続時の債務・葬式費用を反映した「将来の課税対象額」に基づき、二次基礎控除を適用して算出しています。</div>
        </div>
        <div class="formula-card">
            <strong style="color: #d4af37;">3. 遺留分相当額の算定根拠（復元）</strong><br>
            <div class="formula-text">「正味の財産 ＋ 生命保険金の非課税枠」を遺留分算定の基礎とし、総遺留分率（1/2）に各相続人の法定相続分を乗じて、一人当たりの権利行使可能額を算出しています。</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

st.sidebar.success("✅ System-Core v29.0 緊急復旧完了")
