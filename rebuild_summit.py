# =========================================================
# ファイル名: rebuild_summit.py
# 開発責任: 擬似・オーナー監査官 (System-Core v31.10)
# 統括監視: 新・副議長（野党チーム指摘反映・聖典遵守）
# 
# 【更新内容】
# 1. 野党チーム指摘反映: 精算課税・生前贈与の加算に関する計算根拠（エビデンス）の詳細化
# 2. 実務説明責任: 兄弟姉妹相続時の20%加算および代襲相続に関する注釈を強化
# 3. 聖典遵守: 既存の全計算ロジック、タブ構成、UI要素を削除・省略せず完全継承
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
            per_h = h_total_ratio / d(len(children))
            for h in heirs_info:
                if h['type'] in ["子", "孫（養子含む）"]: shares.append(per_h)
                else: shares.append(d(0))
        elif has_parent:
            s_ratio = d("0.6666666666666667") if has_spouse else d(0)
            h_total_ratio = d("0.3333333333333333") if has_spouse else d("1.0")
            parents = [h for h in heirs_info if h['type'] == "親"]
            per_h = h_total_ratio / d(len(parents))
            for h in heirs_info:
                if h['type'] == "親": shares.append(per_h)
                else: shares.append(d(0))
        elif has_sibling:
            s_ratio = d("0.75") if has_spouse else d(0)
            h_total_ratio = d("0.25") if has_spouse else d("1.0")
            weight_sum = d(0)
            for h in heirs_info:
                if h['type'] == "兄弟姉妹（全血）": weight_sum += d(1)
                elif h['type'] == "兄弟姉妹（半血）": weight_sum += d("0.5")
            unit_share = h_total_ratio / weight_sum if weight_sum > 0 else d(0)
            for h in heirs_info:
                if h['type'] == "兄弟姉妹（全血）": shares.append(unit_share)
                elif h['type'] == "兄弟姉妹（半血）": shares.append(unit_share * d("0.5"))
                else: shares.append(d(0))
        else:
            s_ratio = d(1) if has_spouse else d(0)
            shares = [d(0)] * len(heirs_info)
        return s_ratio, shares

    @staticmethod
    def get_tax(taxable_amt, has_spouse, heirs_info):
        d = SupremeLegacyEngine.to_d
        if taxable_amt <= 0: return d(0)
        s_ratio, h_shares = SupremeLegacyEngine.get_legal_shares(has_spouse, heirs_info)
        total_tax = d(0)
        if has_spouse: total_tax += SupremeLegacyEngine.bracket_calc(taxable_amt * s_ratio)
        for share in h_shares:
            if share > 0: total_tax += SupremeLegacyEngine.bracket_calc(taxable_amt * share)
        return total_tax.quantize(d(1), ROUND_HALF_UP)

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

# --- 2. メインUI ---
if check_password():
    st.set_page_config(page_title="SUMMIT v31.10 PRO", layout="wide")
    st.sidebar.markdown("### 🏢 山根会計 専売システム")
    st.sidebar.info("ログイン: 川東")
    
    with st.sidebar.expander("🏛️ デジタル戦略会議 構成員"):
        st.caption("議長: 擬似・オーナー監査官")
        st.caption("野党: リスク・批判担当 (Active)")
        st.caption("推進: システム/実務")
        st.caption("装飾: エグゼクティブ・UI")

    tabs = st.tabs(["👥 1.基本構成", "💰 2.一次財産詳細", "📑 3.一次相続明細", "📑 4.二次相続明細", "⏳ 5.二次推移予測", "📊 6.精密分析結果"])
    d = SupremeLegacyEngine.to_d

    with tabs[0]:
        st.header("相続関係の設定")
        c1, c2 = st.columns(2)
        heir_count = c1.number_input("相続人の人数（配偶者除く）", 1, 10, 2, key="in_child")
        has_spouse = c2.checkbox("配偶者は健在", value=True, key="in_spouse")
        heirs_info = []
        for i in range(heir_count):
            h_type = st.selectbox(f"相続人 {i+1} の続柄", ["子", "孫（養子含む）", "親", "兄弟姉妹（全血）", "兄弟姉妹（半血）", "その他"], key=f"rel_{i}")
            heirs_info.append({"type": h_type})

    with tabs[1]:
        st.header("一次相続：財産入力")
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            st.subheader("🏗️ 不動産")
            v_home = st.number_input("特定居住用：評価額", value=32781936, key="v_home")
            a_home = st.number_input("特定居住用：面積(㎡)", value=330, key="a_home")
            v_biz = st.number_input("特定事業用：評価額", value=0, key="v_biz")
            a_biz = st.number_input("特定事業用：面積(㎡)", value=400, key="a_biz")
            v_rent = st.number_input("貸付事業用：評価額", value=0, key="v_rent")
            a_rent = st.number_input("貸付事業用：面積(㎡)", value=200, key="a_rent")
            v_build = st.number_input("建物評価", value=1700044, key="v_build")
            v_land_others = st.number_input("その他の土地", value=0, key="v_land_others")
        with col_b:
            st.subheader("💵 金融・贈与財産")
            v_stock = st.number_input("有価証券", value=45132788, key="v_stock")
            v_cash = st.number_input("現預金", value=45573502, key="v_cash")
            v_ins = st.number_input("生命保険金", value=3651514, key="v_ins")
            v_others = st.number_input("その他", value=1662687, key="v_others")
            v_gift_3y = st.number_input("相続前贈与（3〜7年以内）", value=0, key="v_gift_3y")
            v_gift_tax_free = st.number_input("相続時精算課税適用財産", value=0, key="v_gift_tax_free")
        with col_c:
            st.subheader("📉 債務・葬式")
            v_debt = st.number_input("債務", value=322179, key="v_debt")
            v_funeral = st.number_input("葬式費用", value=41401, key="v_funeral")

    # 共通ロジック
    st_count = heir_count + (1 if has_spouse else 0)
    a_lim, b_lim, c_lim = d(330), d(400), d(200)
    a_app = min(d(a_home), a_lim); b_app = min(d(a_biz), b_lim)
    used_ratio = (a_app / a_lim) + (b_app / b_lim)
    c_app = min(d(a_rent), c_lim * max(d(0), d(1) - used_ratio))
    red_home = (d(v_home) / d(max(1, a_home))) * a_app * d("0.8")
    red_biz = (d(v_biz) / d(max(1, a_biz))) * b_app * d("0.8")
    red_rent = (d(v_rent) / d(max(1, a_rent))) * c_app * d("0.5")
    land_final = d(v_home) + d(v_biz) + d(v_rent) + d(v_land_others) - (red_home + red_biz + red_rent)
    ins_deduct = min(d(v_ins), d(5000000) * d(st_count))
    pure_assets = land_final + d(v_build) + d(v_stock) + d(v_cash) + d(v_ins) + d(v_others)
    tax_price = pure_assets - ins_deduct - d(v_debt) - d(v_funeral) + d(v_gift_3y) + d(v_gift_tax_free)
    basic_1 = d(30000000) + (d(6000000) * d(st_count))
    taxable_1 = max(d(0), tax_price - basic_1)
    total_tax_1 = SupremeLegacyEngine.get_tax(taxable_1, has_spouse, heirs_info)

    with tabs[2]:
        st.header("一次相続：計算明細")
        st.table(pd.DataFrame([
            ["1", "正味取得財産", f"{int(pure_assets - ins_deduct - v_debt - v_funeral):,}", ""],
            ["2", "生前贈与等加算", f"{int(v_gift_3y + v_gift_tax_free):,}", "野党注記：改正法要確認"],
            ["3", "課税価格合計", f"{int(tax_price):,}", ""],
            ["4", "基礎控除", f"△{int(basic_1):,}", f"相続人{st_count}名"],
            ["5", "相続税の総額", f"{int(total_tax_1):,}", "1円単位精密計算"],
        ], columns=["No", "項目", "金額", "備考"]))

    with tabs[3]:
        st.header("二次相続：計算明細予測")
        share_s = d("0.5"); s_acq_base = tax_price * share_s
        s_limit = max(d(160000000), taxable_1 * share_s)
        s_tax_1 = d(0) if s_acq_base <= s_limit else (total_tax_1 * share_s * d("0.5"))
        s_acq_net = s_acq_base - s_tax_1
        s_own = d(st.session_state.get("in_s_own", 50000000))
        s_gift_2 = d(st.session_state.get("in_s_gift", 0))
        s_spend = d(st.session_state.get("in_s_spend", 5000000)) * d(st.session_state.get("in_interval", 10))
        s_debt_2 = d(st.session_state.get("in_debt2", 5000000))
        tax_price_2 = max(d(0), s_acq_net + s_own + s_gift_2 - s_spend - s_debt_2)
        children_only = [h for h in heirs_info if h['type'] in ["子", "孫（養子含む）"]]
        c_count_2 = len(children_only) if children_only else heir_count
        basic_2 = d(30000000) + (d(6000000) * d(c_count_2))
        total_tax_2 = SupremeLegacyEngine.get_tax(max(d(0), tax_price_2 - basic_2), False, children_only if children_only else heirs_info)
        st.table(pd.DataFrame([
            ["1", "一次からの承継", f"{int(s_acq_net):,}", ""],
            ["2", "固有財産・消費影響", f"{int(s_own + s_gift_2 - s_spend):,}", ""],
            ["3", "二次相続税総額", f"{int(total_tax_2):,}", ""],
        ], columns=["No", "項目", "金額", "備考"]))

    with tabs[4]:
        st.header("二次推移パラメータ設定")
        c_p1, c_p2 = st.columns(2)
        c_p1.number_input("配偶者の固有財産", value=50000000, key="in_s_own")
        c_p1.number_input("生前贈与累計", value=0, key="in_s_gift")
        c_p1.slider("想定期間(年)", 0, 20, 10, key="in_interval")
        c_p2.number_input("年間生活費", value=5000000, key="in_s_spend")
        c_p2.number_input("二次債務・葬式", value=5000000, key="in_debt2")

    with tabs[5]:
        st.header("納税コスト最適化分析")
        res = []
        for r_val in range(0, 101, 10):
            r = d(r_val) / d(100)
            acq_s = tax_price * r
            lim_s = max(d(160000000), taxable_1 * r)
            t_s1 = d(0) if acq_s <= lim_s else (total_tax_1 * r * d("0.5"))
            t_others1 = d(0); _, h_shares = SupremeLegacyEngine.get_legal_shares(has_spouse, heirs_info)
            for i, h in enumerate(heirs_info):
                surcharge = d("1.2") if h['type'] not in ["子", "親"] else d("1.0")
                t_others1 += (total_tax_1 * h_shares[i] * surcharge)
            net_s2 = max(d(0), acq_s - t_s1 + d(s_own) + d(s_gift_2) - s_spend - d(s_debt_2))
            t2 = SupremeLegacyEngine.get_tax(max(d(0), net_s2 - basic_2), False, children_only if children_only else heirs_info)
            res.append({"配分(%)": r_val, "一次税額": int(t_s1 + t_others1), "二次税額": int(t2), "合計": int(t_s1 + t_others1 + t2)})
        df_res = pd.DataFrame(res)
        st.plotly_chart(go.Figure(data=[go.Bar(x=df_res['配分(%)'], y=df_res['一次税額'], name="一次"), go.Bar(x=df_res['配分(%)'], y=df_res['二次税額'], name="二次"), go.Scatter(x=df_res['配分(%)'], y=df_res['合計'], name="合計")]), use_container_width=True)

        st.divider()
        st.markdown("""
        <div style="background-color: #0a192f; color: #d4af37; padding: 30px; border-radius: 12px; border: 3px double #d4af37;">
            <h3 style="text-align: center; text-decoration: underline;">🏛️ v31.10 野党チーム監査：実務エビデンス</h3>
            <p>🔴 <b>【警告：リスク・批判担当】</b>：贈与加算額は額面通りに入力されているが、実務上は贈与税額控除の精査が必須である。また、令和6年以降の生前贈与における「持ち戻し期間7年」および「100万円控除」の適用は、入力値側で調整すること。本システムは入力値を絶対として計算する。</p>
            <p>⚠️ <b>【法的整合性】</b>：兄弟姉妹が相続人の場合、本シミュレーションは「代襲相続人（甥・姪）」を含め、一親等の血族（子・親）以外には一律で20%加算を適用している。これは相続税法第18条に準拠する。</p>
            <p>⚖️ <b>【実務上の説明責任】</b>：本結果を顧客に提示する際は、あくまで「入力された財産評価に基づく試算」であることを強調し、資産税実務担当による個別具体的な税務判断を仰ぐこと。</p>
        </div>
        """, unsafe_allow_html=True)

st.sidebar.success("✅ System-Core v31.10 野党監査・合意完了")
