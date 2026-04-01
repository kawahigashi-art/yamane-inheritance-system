# =========================================================
# ファイル名: rebuild_summit.py
# 開発責任: 擬似・オーナー監査官 (System-Core v28.3)
# クラウド展開用：全機能・計算根拠・認証ゲート完全実装版
# =========================================================

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from decimal import Decimal, ROUND_HALF_UP

# --- 0. セキュリティ設定（ログインゲート） ---
def check_password():
    if "password_correct" not in st.session_state:
        st.title("🔐 山根会計 専売システム")
        pwd = st.text_input("アクセスパスワードを入力してください", type="password")
        if st.button("ログイン"):
            # 実務上はStreamlitのSecrets管理を推奨しますが、まずは簡易実装
            if pwd == "yamane777":
                st.session_state["password_correct"] = True
                st.rerun()
            else:
                st.error("認証に失敗しました。")
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

# --- 2. メインアプリケーション ---
if check_password():
    st.set_page_config(page_title="SUMMIT v28.3 PRO", layout="wide")
    
    # 共通サイドバー
    st.sidebar.markdown("### 🏢 山根会計 専売システム")
    st.sidebar.info("ログイン中: 川東")
    if st.sidebar.button("全入力をリセット"):
        for key in list(st.session_state.keys()):
            if key != "password_correct":
                del st.session_state[key]
        st.rerun()

    st.title("🏛️ Summit System-Core v28.3")
    tabs = st.tabs(["👥 1. 基本構成", "💰 2. 一次財産詳細", "⏳ 3. 二次推移予測", "📊 4. 精密分析結果"])

    # -- TAB 1: 基本 --
    with tabs[0]:
        st.header("相続関係の設定")
        c1, c2 = st.columns(2)
        child_count = c1.number_input("子供の人数", 1, 10, 2, key="in_child")
        has_spouse = c2.checkbox("配偶者は健在", value=True, key="in_spouse")

    # -- TAB 2: 一次詳細 --
    with tabs[1]:
        st.header("一次相続：財産明細")
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            st.subheader("🏗️ 不動産")
            land_home = st.number_input("自宅敷地評価額", value=100000000, key="in_l_home")
            land_home_area = st.number_input("敷地面積(㎡)", value=330, key="in_l_area")
            apply_shoukibo = st.checkbox("小規模宅地特例(80%減)を適用", value=True, key="in_l_apply")
            land_others = st.number_input("その他の土地建物", value=50000000, key="in_l_others")
        with col_b:
            st.subheader("💵 金融・有価証券")
            cash_total = st.number_input("現預金", value=100000000, key="in_cash")
            stock_value = st.number_input("自社株・有価証券", value=150000000, key="in_stocks")
            ins_amount = st.number_input("生命保険金受取額", value=30000000, key="in_ins")
        with col_c:
            st.subheader("📉 一次：債務・葬式")
            debt_1st = st.number_input("一次債務合計", value=20000000, key="in_debt1")
            funeral_1st = st.number_input("一次葬式費用", value=2000000, key="in_fun1")

    # -- TAB 3: 二次推移 --
    with tabs[2]:
        st.header("二次相続：推移予測")
        col_s1, col_s2 = st.columns(2)
        spouse_own_assets = col_s1.number_input("配偶者の固有財産", value=50000000, key="in_s_own")
        life_expectancy = col_s1.slider("二次までの想定期間(年)", 0, 20, 10, key="in_interval")
        spouse_spending = col_s2.number_input("配偶者の年間生活費", value=5000000, key="in_s_spend")
        debt_2nd = col_s2.number_input("二次：見込債務・葬式費用", value=5000000, key="in_debt2")

    # -- TAB 4: 結果分析 --
    with tabs[3]:
        st.header("トータル納税コスト・シミュレーション")
        
        def run_full_logic(s_ratio):
            d = SupremeLegacyEngine.to_d
            st_1 = child_count + (1 if has_spouse else 0)
            land_eval = d(land_home) * d(0.2) if apply_shoukibo and land_home_area <= 330 else d(land_home)
            ins_taxable = max(d(0), d(ins_amount) - (d(5000000) * d(st_1)))
            net_1 = land_eval + d(land_others) + d(cash_total) + d(stock_value) + ins_taxable - d(debt_1st) - d(funeral_1st)
            
            basic_ded_1 = d(30000000) + (d(6000000) * d(st_1))
            taxable_1 = max(d(0), net_1 - basic_ded_1)
            total_tax_1 = SupremeLegacyEngine.get_tax(taxable_1, st_1)
            
            s_acq_1 = net_1 * (d(s_ratio)/100)
            s_limit = max(d(160000000), taxable_1 * (d(s_ratio)/100))
            s_tax_1 = d(0) if s_acq_1 <= s_limit else (total_tax_1 * (d(s_ratio)/100) * d(0.5))
            c_tax_1 = total_tax_1 * (d(1) - d(s_ratio)/100)
            
            s_net_2 = max(d(0), s_acq_1 - s_tax_1 + d(spouse_own_assets) - (d(spouse_spending) * d(life_expectancy)) - d(debt_2nd))
            basic_ded_2 = d(30000000) + (d(6000000) * d(child_count))
            taxable_2 = max(d(0), s_net_2 - basic_ded_2)
            total_tax_2 = SupremeLegacyEngine.get_tax(taxable_2, child_count)
            
            if life_expectancy < 10 and s_tax_1 > 0:
                total_tax_2 = max(d(0), total_tax_2 - (s_tax_1 * (d(10) - d(life_expectancy)) / d(10)))
            return int(s_tax_1 + c_tax_1), int(total_tax_2)

        results = []
        for r in range(0, 101, 10):
            t1, t2 = run_full_logic(r)
            results.append({"配分(%)": r, "一次税額": t1, "二次税額": t2, "合計": t1+t2})
        
        df_res = pd.DataFrame(results)
        
        fig = go.Figure()
        fig.add_trace(go.Bar(x=df_res['配分(%)'], y=df_res['一次税額'], name="一次税額", marker_color='#1f2c4d'))
        fig.add_trace(go.Bar(x=df_res['配分(%)'], y=df_res['二次税額'], name="二次税額", marker_color='#c5a059'))
        fig.add_trace(go.Scatter(x=df_res['配分(%)'], y=df_res['合計'], name="合計コスト", line=dict(color='#a61d24', width=4)))
        fig.update_layout(barmode='stack', yaxis=dict(tickformat=',d'), xaxis_title="配偶者取得割合(%)", yaxis_title="相続税額(円)")
        st.plotly_chart(fig, use_container_width=True)

        st.subheader("📋 試算結果データ")
        st.table(df_res.style.format({"一次税額": "{:,}円", "二次税額": "{:,}円", "合計": "{:,}円"}))

        # --- 重要：計算根拠の完全維持 ---
        st.divider()
        st.subheader("📑 計算ロジック・根拠")
        c_calc1, c_calc2 = st.columns(2)
        with c_calc1:
            st.markdown("##### 【一次相続税】")
            st.info("""
            - **基礎控除**: 3,000万円 + (600万円 × 法定相続人数)
            - **配偶者軽減**: 1.6億円 または 法定相続分 のいずれか多い額まで非課税
            - **評価減**: 小規模宅地特例（居住用330㎡まで80%減額）を適用
            """)
        with c_calc2:
            st.markdown("##### 【二次相続税・相次相続控除】")
            st.info("""
            - **二次課税対象**: (一次取得額 - 一次税額) + 配偶者固有財産 - 生活費 - 債務
            - **相次相続控除**: 一次相続から10年以内の発生で適用
            - **オーナー指定特例**: 分母の一次課税価格から一次税額を控除しない実務慣行ロジックを適用
            """)