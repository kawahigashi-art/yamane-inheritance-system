# =========================================================
# ファイル名: rebuild_summit.py
# 開発責任: 擬似・オーナー監査官 (System-Core v31.13)
# 統括監視: ステート・ポリス（状態維持・印刷制御監視）
# 
# 【修正・更新内容】
# 1. 印刷制御の抜本改善: 印刷時に不要なUI要素を完全に排除し、入力画面のみを抽出。
# 2. スクリーンショット・エミュレーション: 各タブごとに最適化された印刷ビューを実現。
# 3. 聖典遵守: 既存の計算エンジン、UI構造、ビジネスロジックを完全維持。
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

# --- 0.1 印刷専用CSS（スクリーンショット要領の実現） ---
def inject_print_css():
    st.markdown("""
        <style>
        /* 画面表示用スタイル */
        .print-only { display: none; }
        
        /* 印刷実行時の制御（ブラウザの印刷機能に介入） */
        @media print {
            /* サイドバー、ヘッダー、ボタン類を完全に消去 */
            section[data-testid="stSidebar"], 
            header, 
            .stButton, 
            div[data-testid="stToolbar"],
            footer {
                display: none !important;
            }
            
            /* メインコンテンツの余白調整 */
            .main .block-container {
                padding: 0 !important;
                margin: 0 !important;
            }
            
            /* タブの境界線を消し、中身だけを抽出 */
            div[data-testid="stExpander"] {
                border: none !important;
            }
            
            /* 背景を白、文字を黒に固定 */
            body {
                background-color: white !important;
                color: black !important;
            }
        }
        </style>
    """, unsafe_allow_html=True)

# --- 0.2 印刷実行コンポーネント ---
def add_print_button(tab_name):
    """
    JavaScriptを用いてブラウザの印刷ダイアログを起動。
    iframe内から親ウィンドウのprintを叩く。
    """
    html_code = f"""
        <div style="display: flex; justify-content: flex-end; margin-bottom: 10px;">
            <button onclick="window.parent.print()" style="
                background-color: #1f2c4d; 
                color: #c5a059; 
                border: 2px solid #c5a059; 
                padding: 8px 16px; 
                border-radius: 4px; 
                cursor: pointer;
                font-weight: bold;
                font-family: 'sans-serif';
            ">
                🖨️ {tab_name} を印刷(PDF)
            </button>
        </div>
    """
    components.html(html_code, height=50)

# --- 1. 超精密計算エンジン（既存ロジック完全維持） ---
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
    st.set_page_config(page_title="SUMMIT v31.13 PRO", layout="wide")
    inject_print_css()

    st.sidebar.markdown("###  🏢  山根会計 専売システム")
    st.sidebar.info("ログイン: 川東")
    st.sidebar.markdown("---")
    st.sidebar.write("【印刷時の注意】")
    st.sidebar.caption("印刷ボタンを押すと、現在表示されているタブの内容がPDF/印刷対象となります。背景グラフィックの設定をオンにして印刷してください。")

    tabs = st.tabs([" 👥  1.基本構成", " 💰  2.一次財産詳細", " 📑  3.一次相続明細", " 📑  4.二次相続明細", " ⏳  5.二次推移予測", " 📊  6.精密分析結果"])
    d = SupremeLegacyEngine.to_d
    
    # -- TAB 1: 基本構成 --
    with tabs[0]:
        add_print_button("1.基本構成")
        st.subheader("相続関係の設定")
        c1, c2 = st.columns(2)
        heir_count = c1.number_input("相続人の人数（配偶者除く）", 1, 10, 2, key="in_child")
        has_spouse = c2.checkbox("配偶者は健在", value=True, key="in_spouse")
        heirs_info = []
        for i in range(heir_count):
            h_type = st.selectbox(f"相続人 {i+1} の続柄", ["子", "孫（養子含む）", "親", "兄弟姉妹（全血）", "兄弟姉妹（半血）"], key=f"rel_{i}")
            heirs_info.append({"type": h_type})

    # -- TAB 2: 一次財産詳細 --
    with tabs[1]:
        add_print_button("2.一次財産詳細")
        st.subheader("一次相続：財産入力")
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            st.write("#### 🏗️ 不動産")
            v_home = st.number_input("特定居住用：評価額", value=32781936, key="v_home")
            a_home = st.number_input("特定居住用：面積(㎡)", value=330, key="a_home")
            v_biz = st.number_input("特定事業用：評価額", value=0, key="v_biz")
            a_biz = st.number_input("特定事業用：面積(㎡)", value=400, key="a_biz")
            v_rent = st.number_input("貸付事業用：評価額", value=0, key="v_rent")
            a_rent = st.number_input("貸付事業用：面積(㎡)", value=200, key="a_rent")
            v_build = st.number_input("建物評価", value=1700044, key="v_build")
            v_land_others = st.number_input("その他の土地", value=0, key="v_land_others")
        with col_b:
            st.write("#### 💵 金融・贈与財産")
            v_stock = st.number_input("有価証券", value=45132788, key="v_stock")
            v_cash = st.number_input("現預金", value=45573502, key="v_cash")
            v_ins = st.number_input("生命保険金", value=3651514, key="v_ins")
            v_others = st.number_input("その他", value=1662687, key="v_others")
            v_gift_3y = st.number_input("相続前贈与（3〜7年）", value=0, key="v_gift_3y")
            v_gift_tax_free = st.number_input("相続時精算課税財産", value=0, key="v_gift_tax_free")
        with col_c:
            st.write("#### 📉 債務・葬式")
            v_debt = st.number_input("債務", value=322179, key="v_debt")
            v_funeral = st.number_input("葬式費用", value=41401, key="v_funeral")

        # -- 計算ロジック --
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
        add_print_button("3.一次相続明細")
        st.subheader("一次相続：計算明細")
        df1 = pd.DataFrame([
            ["1", "不動産評価（特例適用後）", f"{int(land_eval):,}", "小規模宅地特例減額反映済み"],
            ["2", "建物評価額", f"{int(v_build):,}", ""],
            ["3", "有価証券", f"{int(v_stock):,}", ""],
            ["4", "現預金", f"{int(v_cash):,}", ""],
            ["5", "生命保険金", f"{int(v_ins):,}", ""],
            ["6", "その他財産", f"{int(v_others):,}", ""],
            ["7", "生命保険非課税限度額", f"△{int(ins_ded):,}", f"500万円 × {st_count}名"],
            ["8", "債務および葬式費用", f"△{int(v_debt + v_funeral):,}", ""],
            ["9", "生前贈与加算財産", f"{int(v_gift_3y):,}", ""],
            ["10", "相続時精算課税適用財産", f"{int(v_gift_tax_free):,}", ""],
            ["11", "【課税価格合計】", f"{int(tax_p):,}", ""],
            ["12", "遺産に係る基礎控除額", f"△{int(basic_1):,}", f"3000万+(600万×{st_count})"],
            ["13", "課税遺産総額", f"{int(taxable_1):,}", ""],
            ["14", "【相続税の総額】", f"{int(total_tax_1):,}", "法定相続分による合算"],
        ], columns=["No", "項目", "金額", "備考"])
        st.table(df1)

    # -- TAB 4: 二次相続明細 --
    with tabs[3]:
        add_print_button("4.二次相続明細")
        st.subheader("二次相続：計算明細予測")
        ratio_s = d("0.5")
        acq_s_1 = tax_p * ratio_s
        limit_s = max(d(160000000), taxable_1 * ratio_s)
        tax_s_1 = d(0) if acq_s_1 <= limit_s else (total_tax_1 * ratio_s * d("0.5"))
        net_acq_s = acq_s_1 - tax_s_1
        
        s_own = d(st.session_state.get("in_s_own", 50000000))
        s_gift_2 = d(st.session_state.get("in_s_gift", 0))
        s_spend_total = d(st.session_state.get("in_s_spend", 5000000)) * d(st.session_state.get("in_interval", 10))
        s_debt_2nd = d(st.session_state.get("in_debt2", 5000000))
        tax_p_2 = max(d(0), net_acq_s + s_own + s_gift_2 - s_spend_total - s_debt_2nd)
        
        child_only = [h for h in heirs_info if h['type'] in ["子", "孫（養子含む）"]]
        c_count_2 = len(child_only) if child_only else heir_count
        basic_2 = d(30000000) + (d(6000000) * d(c_count_2))
        taxable_2 = max(d(0), tax_p_2 - basic_2)
        total_tax_2 = SupremeLegacyEngine.get_tax(taxable_2, False, child_only if child_only else heirs_info)
        
        df2 = pd.DataFrame([
            ["1", "一次相続からの純承継分", f"{int(net_acq_s):,}", "配偶者軽減適用後"],
            ["2", "配偶者固有の財産", f"{int(s_own):,}", ""],
            ["3", "生前贈与加算財産（二次）", f"{int(s_gift_2):,}", ""],
            ["4", "想定生活費消費累計", f"△{int(s_spend_total):,}", f"期間：{st.session_state.get('in_interval', 10)}年"],
            ["5", "二次相続時の債務・葬式費用", f"△{int(s_debt_2nd):,}", ""],
            ["6", "【二次相続 課税価格】", f"{int(tax_p_2):,}", ""],
            ["7", "二次基礎控除額", f"△{int(basic_2):,}", f"相続人{c_count_2}名"],
            ["8", "課税遺産総額（二次）", f"{int(taxable_2):,}", ""],
            ["9", "【二次相続税 総額】", f"{int(total_tax_2):,}", ""],
        ], columns=["No", "項目", "金額", "備考"])
        st.table(df2)

    # -- TAB 5: 二次推移予測 --
    with tabs[4]:
        add_print_button("5.二次推移予測")
        st.subheader("二次推移パラメータ設定")
        cp1, cp2 = st.columns(2)
        cp1.number_input("配偶者の固有財産", value=50000000, key="in_s_own")
        cp1.number_input("生前贈与累計", value=0, key="in_s_gift")
        cp1.slider("想定期間(年)", 0, 20, 10, key="in_interval")
        cp2.number_input("年間生活費支出", value=5000000, key="in_s_spend")
        cp2.number_input("二次債務・葬式費用", value=5000000, key="in_debt2")

    # -- TAB 6: 精密分析結果 --
    with tabs[5]:
        add_print_button("6.精密分析結果")
        st.subheader("配偶者取得割合別の税額推移分析")
        sim_results = []
        for i in range(0, 101, 10):
            ratio = d(i) / d(100)
            acq_s = tax_p * ratio
            limit_s = max(d(160000000), taxable_1 * ratio)
            t_s_1 = d(0) if acq_s <= limit_s else (total_tax_1 * ratio * d("0.5"))
            t1 = total_tax_1 - (total_tax_1 * ratio) + t_s_1
            net_s = acq_s - t_s_1
            tp2 = max(d(0), net_s + s_own + s_gift_2 - s_spend_total - s_debt_2nd)
            taxable_2_sim = max(d(0), tp2 - basic_2)
            t2 = SupremeLegacyEngine.get_tax(taxable_2_sim, False, child_only if child_only else heirs_info)
            sim_results.append({
                "配分(%)": f"{i}%",
                "一次相続税額": int(t1),
                "二次相続税額": int(t2),
                "合計納税額": int(t1 + t2)
            })
        
        df_sim = pd.DataFrame(sim_results)
        fig = go.Figure()
        fig.add_trace(go.Bar(x=df_sim['配分(%)'], y=df_sim['一次相続税額'], name='一次相続税', marker_color='#1f2c4d'))
        fig.add_trace(go.Bar(x=df_sim['配分(%)'], y=df_sim['二次相続税額'], name='二次相続税', marker_color='#c5a059'))
        fig.add_trace(go.Scatter(x=df_sim['配分(%)'], y=df_sim['合計納税額'], name='合計', line=dict(color='#a61d24', width=4)))
        fig.update_layout(barmode='stack', title="配偶者取得割合別の税額推移", xaxis_title="配偶者の取得割合(%)", yaxis_title="税額(円)")
        st.plotly_chart(fig, use_container_width=True)
        
        st.table(df_sim.style.format({
            "一次相続税額": "{:,}円",
            "二次相続税額": "{:,}円",
            "合計納税額": "{:,}円"
        }))

        st.divider()
        st.markdown("""
        <div style="background-color: #f9f9f9; border: 1px solid #c5a059; padding: 20px; border-radius: 5px;">
            <p style="color: #1f2c4d; font-weight: bold;">【山根会計 監査証跡】</p>
            <p style="font-size: 0.85em; color: #333;">本結果は、川東（ログインユーザー）による設定パラメータに基づき出力されました。
            印刷時には、表示中の各タブ項目がそのままキャプチャされます。</p>
        </div>
        """, unsafe_allow_html=True)
