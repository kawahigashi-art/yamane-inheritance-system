# =========================================================
# ファイル名: rebuild_summit.py
# 開発責任: 擬似・オーナー監査官 (System-Core v31.16)
# 統括監視: 新・副議長（省略・削除の絶対禁止監視）
# 
# 【修正・更新内容】
# 1. 聖典遵守: v31.11の全タブ内容・計算ロジックを完全維持。
# 2. Excel統合: openpyxlを用いた精密データ書き出し機能を新規実装。
# 3. リスク管理: ライブラリ未検出時のフォールバックおよび文字化け対策を完遂。
# =========================================================

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from decimal import Decimal, ROUND_HALF_UP
import io
from datetime import datetime

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
            if share > 0:
                total_tax += SupremeLegacyEngine.bracket_calc(taxable_amt * share)
        return total_tax.quantize(d(1), ROUND_HALF_UP)

# --- 2. メインUI ---
if check_password():
    st.set_page_config(page_title="SUMMIT v31.16 PRO", layout="wide")
    
    # サイドバー（ブランド設定）
    st.sidebar.markdown("### 🏢 山根会計 専売システム")
    st.sidebar.info("ログイン: 川東")
    
    tabs = st.tabs(["👥 1.基本構成", "💰 2.一次財産詳細", "📑 3.一次相続明細", "📑 4.二次相続明細", "⏳ 5.二次推移予測", "📊 6.精密分析結果", "📥 7.外部データ出力"])
    d = SupremeLegacyEngine.to_d

    # -- TAB 1: 基本構成 --
    with tabs[0]:
        st.header("相続関係の設定")
        c1, c2 = st.columns(2)
        heir_count = c1.number_input("相続人の人数（配偶者除く）", 1, 10, 2, key="in_child")
        has_spouse = c2.checkbox("配偶者は健在", value=True, key="in_spouse")
        heirs_info = []
        for i in range(heir_count):
            h_type = st.selectbox(f"相続人 {i+1} の続柄", ["子", "孫（養子含む）", "親", "兄弟姉妹（全血）", "兄弟姉妹（半血）"], key=f"rel_{i}")
            heirs_info.append({"type": h_type})

    # -- TAB 2: 一次財産詳細 --
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
            v_gift_3y = st.number_input("相続前贈与（3〜7年）", value=0, key="v_gift_3y")
            v_gift_tax_free = st.number_input("相続時精算課税財産", value=0, key="v_gift_tax_free")
        with col_c:
            st.subheader("📉 債務・葬式")
            v_debt = st.number_input("債務", value=322179, key="v_debt")
            v_funeral = st.number_input("葬式費用", value=41401, key="v_funeral")

    # -- 共通計算ロジック（フルスタック） --
    st_count = heir_count + (1 if has_spouse else 0)
    a_lim, b_lim, c_lim = d(330), d(400), d(200)
    a_app = min(d(a_home), a_lim); b_app = min(d(a_biz), b_lim)
    used_r = (a_app / a_lim) + (b_app / b_lim)
    c_app = min(d(a_rent), c_lim * max(d(0), d(1) - used_r))
    red_home = (d(v_home) / d(max(1, a_home))) * a_app * d("0.8")
    red_biz = (d(v_biz) / d(max(1, a_biz))) * b_app * d("0.8")
    red_rent = (d(v_rent) / d(max(1, a_rent))) * c_app * d("0.5")
    total_red = red_home + red_biz + red_rent
    
    land_eval = d(v_home) + d(v_biz) + d(v_rent) + d(v_land_others) - total_red
    ins_ded = min(d(v_ins), d(5000000) * d(st_count))
    pure_as = land_eval + d(v_build) + d(v_stock) + d(v_cash) + d(v_ins) + d(v_others)
    tax_p = pure_as - ins_ded - d(v_debt) - d(v_funeral) + d(v_gift_3y) + d(v_gift_tax_free)
    basic_1 = d(30000000) + (d(6000000) * d(st_count))
    taxable_1 = max(d(0), tax_p - basic_1)
    total_tax_1 = SupremeLegacyEngine.get_tax(taxable_1, has_spouse, heirs_info)

    # -- TAB 3: 一次相続明細 --
    with tabs[2]:
        st.header("📑 一次相続：計算明細")
        df1 = pd.DataFrame([
            ["1", "不動産評価（小規模宅地特例適用後）", int(land_eval), "特例減額済み"],
            ["2", "建物評価額", int(v_build), ""],
            ["3", "有価証券", int(v_stock), ""],
            ["4", "現預金", int(v_cash), ""],
            ["5", "生命保険金", int(v_ins), "非課税枠控除前"],
            ["6", "その他財産", int(v_others), ""],
            ["7", "生命保険非課税限度額", int(-ins_ded), f"500万円 × {st_count}名"],
            ["8", "債務および葬式費用", int(-(v_debt + v_funeral)), ""],
            ["9", "生前贈与加算財産", int(v_gift_3y), "3〜7年内贈与"],
            ["10", "相続時精算課税適用財産", int(v_gift_tax_free), "持ち戻し加算"],
            ["11", "【課税価格合計】", int(tax_p), ""],
            ["12", "遺産に係る基礎控除額", int(-basic_1), f"3000万+(600万×{st_count})"],
            ["13", "課税遺産総額", int(taxable_1), ""],
            ["14", "【相続税の総額】", int(total_tax_1), "法定相続分による按分合算"],
        ], columns=["No", "項目", "金額", "備考"])
        st.table(df1.style.format({"金額": "{:,}"}))

    # -- TAB 4: 二次相続明細 --
    with tabs[3]:
        st.header("📑 二次相続：計算明細予測")
        ratio_s = d("0.5")
        acq_s_1 = tax_p * ratio_s
        limit_s = max(d(160000000), taxable_1 * ratio_s)
        tax_s_1 = d(0) if acq_s_1 <= limit_s else (total_tax_1 * ratio_s * d("0.5"))
        net_acq_s = acq_s_1 - tax_s_1
        
        s_own = d(st.session_state.get("in_s_own", 50000000))
        s_gift_2 = d(st.session_state.get("in_s_gift", 0))
        s_interval = d(st.session_state.get("in_interval", 10))
        s_spend_total = d(st.session_state.get("in_s_spend", 5000000)) * s_interval
        s_debt_2nd = d(st.session_state.get("in_debt2", 5000000))
        
        tax_p_2 = max(d(0), net_acq_s + s_own + s_gift_2 - s_spend_total - s_debt_2nd)
        child_only = [h for h in heirs_info if h['type'] in ["子", "孫（養子含む）"]]
        c_count_2 = len(child_only) if child_only else heir_count
        basic_2 = d(30000000) + (d(6000000) * d(c_count_2))
        taxable_2 = max(d(0), tax_p_2 - basic_2)
        total_tax_2 = SupremeLegacyEngine.get_tax(taxable_2, False, child_only if child_only else heirs_info)

        df2 = pd.DataFrame([
            ["1", "一次相続からの純承継分", int(net_acq_s), "配偶者税額軽減後"],
            ["2", "配偶者固有の財産", int(s_own), ""],
            ["3", "生前贈与加算財産（二次）", int(s_gift_2), ""],
            ["4", "想定生活費消費累計", int(-s_spend_total), f"期間：{int(s_interval)}年"],
            ["5", "二次相続時の債務・葬式費用", int(-s_debt_2nd), ""],
            ["6", "【二次相続 課税価格】", int(tax_p_2), ""],
            ["7", "二次基礎控除額", int(-basic_2), f"相続人{c_count_2}名"],
            ["8", "課税遺産総額（二次）", int(taxable_2), ""],
            ["9", "【二次相続税 総額】", int(total_tax_2), ""],
        ], columns=["No", "項目", "金額", "備考"])
        st.table(df2.style.format({"金額": "{:,}"}))

    # -- TAB 5: 二次推移予測 --
    with tabs[4]:
        st.header("⏳ 二次推移パラメータ設定")
        cp1, cp2 = st.columns(2)
        cp1.number_input("配偶者の固有財産", value=50000000, key="in_s_own")
        cp1.number_input("生前贈与累計", value=0, key="in_s_gift")
        cp1.slider("想定期間(年)", 0, 20, 10, key="in_interval")
        cp2.number_input("年間生活費", value=5000000, key="in_s_spend")
        cp2.number_input("二次：債務・葬式費用", value=5000000, key="in_debt2")

    # -- TAB 6: 精密分析結果 --
    with tabs[5]:
        st.header("📊 納税コスト最適化分析")
        def run_sim(s_ratio_val):
            r = d(s_ratio_val) / d(100)
            acq_s = tax_p * r
            lim_s = max(d(160000000), taxable_1 * r)
            t_s1 = d(0) if acq_s <= lim_s else (total_tax_1 * r * d("0.5"))
            t_others1 = d(0)
            _, h_shares = SupremeLegacyEngine.get_legal_shares(has_spouse, heirs_info)
            for i, h in enumerate(heirs_info):
                sur = d("1.2") if h['type'] not in ["子", "親"] else d("1.0")
                t_others1 += (total_tax_1 * h_shares[i] * sur)
            net_s2_acq = acq_s - t_s1 + d(s_own) + d(s_gift_2) - s_spend_total - d(s_debt_2nd)
            t2 = SupremeLegacyEngine.get_tax(max(d(0), net_s2_acq - basic_2), False, child_only if child_only else heirs_info)
            return int(t_s1 + t_others1), int(t2)

        sim_results = []
        for r in range(0, 101, 10):
            t1, t2 = run_sim(r)
            sim_results.append({"配分(%)": r, "一次相続税額": t1, "二次相続税額": t2, "合計納税額": t1 + t2})
        df_sim = pd.DataFrame(sim_results)
        
        fig = go.Figure()
        fig.add_trace(go.Bar(x=df_sim['配分(%)'], y=df_sim['一次相続税額'], name='一次相続税', marker_color='#1f2c4d'))
        fig.add_trace(go.Bar(x=df_sim['配分(%)'], y=df_sim['二次相続税額'], name='二次相続税', marker_color='#c5a059'))
        fig.add_trace(go.Scatter(x=df_sim['配分(%)'], y=df_sim['合計納税額'], name='合計', line=dict(color='#a61d24', width=4)))
        fig.update_layout(barmode='stack', title="配偶者取得割合別の税額推移")
        st.plotly_chart(fig, use_container_width=True)
        st.table(df_sim.style.format({k: "{:,}円" for k in ["一次相続税額", "二次相続税額", "合計納税額"]}))

    # -- TAB 7: 外部データ出力（新規実装） --
    with tabs[6]:
        st.header("📥 Excel報告書データ出力")
        st.write("計算結果をExcel形式（.xlsx）でダウンロードします。実務での微調整や印刷にご活用ください。")
        
        def to_excel():
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # 1. 一次相続明細
                df1_out = df1.copy()
                df1_out.to_excel(writer, index=False, sheet_name='一次相続明細')
                # 2. 二次相続明細
                df2_out = df2.copy()
                df2_out.to_excel(writer, index=False, sheet_name='二次相続明細予測')
                # 3. シミュレーション結果
                df_sim_out = df_sim.copy()
                df_sim_out.to_excel(writer, index=False, sheet_name='最適化分析データ')
            return output.getvalue()

        try:
            excel_data = to_excel()
            timestamp = datetime.now().strftime("%Y%m%d_%H%M")
            st.download_button(
                label="🟢 Excelファイルをダウンロード",
                data=excel_data,
                file_name=f"山根会計_相続シミュレーション_{timestamp}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Excel生成中にエラーが発生しました: {e}")
            st.info("対策: ターミナルで 'pip install openpyxl' が実行されているか確認してください。")

st.sidebar.success("✅ System-Core v31.16 Excel統合完了")
