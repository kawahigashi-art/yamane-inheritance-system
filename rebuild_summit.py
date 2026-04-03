# =========================================================
# ファイル名: rebuild_summit.py
# 開発責任: 擬似・オーナー監査官 (System-Core v31.17)
# 統括監視: 野党チーム（中略・削除・省略の絶対禁止監視）
# 
# 【修正・更新内容】
# 1. 聖典完全遵守: 過去に中略されていた全ロジック（贈与・特例・遺留分）を完全復元。
# 2. 状態保護: ステート・ポリスの指針に基づき、ウィジェットIDの重複を排除。
# 3. 野党指摘反映: 二次相続シミュレーション（0-100%）を完全可視化し、中略を永久追放。
# =========================================================

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from decimal import Decimal, ROUND_HALF_UP
import streamlit.components.v1 as components

# --- 0. セキュリティ・ページ設定 ---
def check_password():
    if "password_correct" not in st.session_state:
        st.set_page_config(page_title="山根会計 専売システム", layout="wide")
        st.markdown("""
            <style>
            .main { background-color: #f5f5f5; }
            .stButton>button { background-color: #1f2c4d; color: #c5a059; border-radius: 5px; }
            h1 { color: #1f2c4d; border-bottom: 3px solid #c5a059; }
            </style>
        """, unsafe_allow_html=True)
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

# --- 1. 超精密計算エンジン (Decimal実装) ---
class SupremeTaxEngine:
    @staticmethod
    def d(value):
        return Decimal(str(value))

    @staticmethod
    def round_down(val):
        return val.to_integral_value(rounding=ROUND_HALF_UP) # 相続税法上の端数処理

    @staticmethod
    def get_tax_rate(taxable_amount):
        # 日本の相続税率（2024年現在）
        if taxable_amount <= 10000000: return 0.10, 0
        elif taxable_amount <= 30000000: return 0.15, 500000
        elif taxable_amount <= 50000000: return 0.20, 2000000
        elif taxable_amount <= 100000000: return 0.30, 7000000
        elif taxable_amount <= 200000000: return 0.40, 17000000
        elif taxable_amount <= 300000000: return 0.45, 27000000
        elif taxable_amount <= 600000000: return 0.50, 42000000
        else: return 0.55, 72000000

    @classmethod
    def calculate_inheritance_tax(cls, total_assets, has_spouse, num_children, gift_added=0, spouse_share_ratio=0.5):
        d = cls.d
        num_heirs = (1 if has_spouse else 0) + num_children
        if num_heirs <= 0: return d(0), d(0), d(0)

        # 基礎控除
        basic_deduction = d(30000000) + d(60000000) * d(num_heirs) / d(10) # 3000万 + 600万×法定相続人数
        
        # 課税遺産総額
        taxable_total = max(d(0), d(total_assets) + d(gift_added) - basic_deduction)
        if taxable_total == 0: return d(0), d(0), d(0)

        # 法定相続分での仮の税額合算
        total_tax = d(0)
        shares = []
        if has_spouse:
            spouse_legal_share = d(0.5) if num_children > 0 else d(1.0)
            shares.append(spouse_legal_share)
            for _ in range(num_children):
                shares.append(d(0.5) / d(num_children))
        else:
            for _ in range(num_children):
                shares.append(d(1.0) / d(num_children))

        for share in shares:
            amount = taxable_total * share
            rate, deduction = cls.get_tax_rate(amount)
            total_tax += amount * d(rate) - d(deduction)

        # 実際の配分（配偶者控除適用前）
        spouse_tax_before = total_tax * d(spouse_share_ratio) if has_spouse else d(0)
        
        # 配偶者控除（1.6億または法定相続分まで非課税）
        if has_spouse:
            spouse_actual_assets = taxable_total * d(spouse_share_ratio)
            legal_share_amount = taxable_total * (d(0.5) if num_children > 0 else d(1.0))
            exemption_limit = max(d(160000000), legal_share_amount)
            
            if spouse_actual_assets <= exemption_limit:
                spouse_tax_final = d(0)
            else:
                spouse_tax_final = spouse_tax_before * (spouse_actual_assets - exemption_limit) / spouse_actual_assets
        else:
            spouse_tax_final = d(0)

        child_tax_total = total_tax - spouse_tax_before
        return total_tax, spouse_tax_final, child_tax_total

# --- 2. メインアプリケーション ---
def main():
    if not check_password(): return

    st.markdown(f"""
        <div style="text-align: center; padding: 10px; background: linear-gradient(to right, #1f2c4d, #c5a059); border-radius: 10px;">
            <h1 style="color: white; border: none; margin: 0;">SUMMIT v31.17 PRO</h1>
            <p style="color: #f5f5f5; margin: 5px;">山根会計 資産税実務特化型シミュレーター</p>
        </div>
    """, unsafe_allow_html=True)

    # 印刷用JavaScript
    st.sidebar.markdown("---")
    if st.sidebar.button("🖨️ この画面をPDF保存/印刷"):
        components.html("<script>window.print();</script>", height=0)

    # タブ構成
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "💎 財産・基本設定", 
        "🎁 贈与・特例適用", 
        "📊 一次相続計算", 
        "🔄 二次相続最適化", 
        "⚠️ 遺留分監査"
    ])

    with tab1:
        st.subheader("基本情報入力")
        col1, col2 = st.columns(2)
        with col1:
            total_assets = st.number_input("現預金・有価証券 (円)", value=100000000, step=1000000)
            real_estate = st.number_input("不動産評価額 (円)", value=50000000, step=1000000)
        with col2:
            has_spouse = st.checkbox("配偶者あり", value=True)
            num_children = st.number_input("子供の人数", min_value=0, max_value=10, value=2)

    with tab2:
        st.subheader("生前贈与・小規模宅地特例")
        col3, col4 = st.columns(2)
        with col3:
            gift_added = st.number_input("生前贈与加算 (3〜7年内分) (円)", value=0, step=100000)
            seisan_kaikei = st.number_input("相続時精算課税 累計額 (円)", value=0, step=100000)
        with col4:
            st.info("小規模宅地等の特例")
            is_tokurei = st.checkbox("特例を適用する")
            tokurei_val = st.number_input("特例による減額幅 (円)", value=0 if not is_tokurei else 20000000)

    net_assets = total_assets + real_estate - tokurei_val
    total_gift = gift_added + seisan_kaikei

    with tab3:
        st.subheader("一次相続税額の詳細")
        spouse_ratio = st.slider("配偶者の取得割合 (%)", 0, 100, 50) / 100
        
        t_tax, s_tax, c_tax = SupremeTaxEngine.calculate_inheritance_tax(net_assets, has_spouse, num_children, total_gift, spouse_ratio)
        
        c1, c2, c3 = st.columns(3)
        c1.metric("総税額 (仮算定)", f"{int(t_tax):,}円")
        c2.metric("配偶者納税額", f"{int(s_tax):,}円")
        c3.metric("子供納税合計", f"{int(c_tax):,}円")
        
        st.markdown(f"""
            <div style="background-color: #1f2c4d; color: white; padding: 15px; border-radius: 5px; border-left: 10px solid #c5a059;">
                <b>監査報告:</b> 一次相続における納税合計は <b>{int(s_tax + c_tax):,}円</b> です。
            </div>
        """, unsafe_allow_html=True)

    with tab4:
        st.subheader("二次相続を含めたトータルコスト最適化分析")
        if not has_spouse:
            st.warning("配偶者がいない場合、二次相続シミュレーションは不要です。")
        else:
            sim_results = []
            for r in range(0, 101, 10):
                ratio = Decimal(r) / Decimal(100)
                # 一次
                _, t1_s, t1_c = SupremeTaxEngine.calculate_inheritance_tax(net_assets, True, num_children, total_gift, float(ratio))
                # 二次 (配偶者の固有財産は0と仮定)
                secondary_base = net_assets * ratio
                _, _, t2_c = SupremeTaxEngine.calculate_inheritance_tax(secondary_base, False, num_children, 0, 0)
                
                sim_results.append({
                    "配分(%)": f"{r}%",
                    "一次相続税額": int(t1_s + t1_c),
                    "二次相続税額": int(t2_c),
                    "合計納税額": int(t1_s + t1_c + t2_c)
                })
            
            df_sim = pd.DataFrame(sim_results)
            
            fig = go.Figure()
            fig.add_trace(go.Bar(x=df_sim['配分(%)'], y=df_sim['一次相続税額'], name='一次相続税', marker_color='#1f2c4d'))
            fig.add_trace(go.Bar(x=df_sim['配分(%)'], y=df_sim['二次相続税額'], name='二次相続税', marker_color='#c5a059'))
            fig.add_trace(go.Scatter(x=df_sim['配分(%)'], y=df_sim['合計納税額'], name='合計', line=dict(color='#a61d24', width=4)))
            fig.update_layout(barmode='stack', title="配偶者取得割合別の税額推移（全パターン）")
            st.plotly_chart(fig, use_container_width=True)
            
            st.markdown("### 数値データ詳細（監査用全件出力）")
            st.table(df_sim.style.format("{:,}円", subset=["一次相続税額", "二次相続税額", "合計納税額"]))

    with tab5:
        st.subheader("⚠️ 遺留分侵害リスク監査")
        st.write("法定相続分に基づき、各相続人の遺留分（最低保障額）を算出します。")
        
        iryu_total = net_assets / 2
        iryu_data = []
        if has_spouse:
            s_share = Decimal(0.5) if num_children > 0 else Decimal(1.0)
            iryu_data.append({"相続人": "配偶者", "遺留分": f"{int(iryu_total * s_share):,}円"})
        
        if num_children > 0:
            c_share = (Decimal(0.5) / Decimal(num_children)) if has_spouse else (Decimal(1.0) / Decimal(num_children))
            for i in range(num_children):
                iryu_data.append({"相続人": f"子供 {i+1}", "遺留分": f"{int(iryu_total * c_share):,}円"})
        
        st.table(iryu_data)
        st.warning("※遺言等でこの金額を下回る配分を指定した場合、遺留分侵害額請求の対象となる可能性があります。")

if __name__ == "__main__":
    main()
