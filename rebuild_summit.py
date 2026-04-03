# =========================================================
# ファイル名: rebuild_summit.py
# 開発責任: 擬似・オーナー監査官 (System-Core v31.15)
# 統括監視: ステート・ポリス（削除・省略・中略の絶対禁止監視）
# 
# 【修正・更新内容】
# 1. 印刷機能の完全統合: 全てのタブに独立した「このページを印刷」ボタンを実装。
# 2. 聖典の完全遵守: 基盤コードにある計算、表示、遺留分、監査証跡を1行も削らず保持。
# 3. 印刷用CSS最適化: 印刷時にサイドバーやボタン自体が映り込まないよう制御。
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

# --- 0.1 印刷専用スタイルシート (印刷時に不要なUIを隠す) ---
def inject_print_css():
    st.markdown("""
        <style>
        @media print {
            /* サイドバー、ヘッダー、ツールバー、ボタンを非表示 */
            section[data-testid="stSidebar"], 
            header, 
            .stButton, 
            div[data-testid="stToolbar"],
            footer {
                display: none !important;
            }
            /* 余白調整 */
            .main .block-container { 
                padding: 0 !important; 
                margin: 0 !important; 
            }
            /* グラフやテーブルが途切れないように設定 */
            .stPlotlyChart { page-break-inside: avoid; }
            table { page-break-inside: auto; }
            tr { page-break-inside: avoid; page-break-after: auto; }
        }
        /* 画面上での印刷ボタンのデザイン */
        .print-btn-container {
            display: flex;
            justify-content: flex-end;
            margin-bottom: 20px;
        }
        </style>
    """, unsafe_allow_html=True)

# --- 0.2 タブ別印刷実行コンポーネント ---
def add_print_button(tab_name):
    # JavaScriptによる印刷実行
    html_code = f"""
        <div class="print-btn-container">
            <button onclick="window.parent.print()" style="
                background-color: #1f2c4d; 
                color: #c5a059; 
                border: 2px solid #c5a059; 
                padding: 10px 20px; 
                border-radius: 5px; 
                cursor: pointer;
                font-weight: bold;
                font-family: 'MS Gothic', sans-serif;
            ">
                🖨️ 「{tab_name}」を印刷 / PDF保存
            </button>
        </div>
    """
    components.html(html_code, height=60)

# --- 1. 計算エンジン（基盤コード完全維持） ---
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
    st.set_page_config(page_title="SUMMIT v31.15 PRO", layout="wide")
    inject_print_css()

    st.sidebar.markdown("###  🏢  山根会計 専売システム")
    st.sidebar.info("ログイン: 川東")

    tabs = st.tabs([" 👥  1.基本構成", " 💰  2.一次財産詳細", " 📑  3.一次相続明細", " 📑  4.二次相続明細", " ⏳  5.二次推移予測", " 📊  6.精密分析結果"])
    d = SupremeLegacyEngine.to_d
    
    # -- TAB 1: 基本構成 --
    with tabs[0]:
        add_print_button("1. 基本構成")
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
        add_print_button("2. 一次財産詳細")
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
        with col_b:
            st.write("#### 💵 金融・贈与財産")
            v_stock = st.number_input("有価証券", value=45132788, key="v_stock")
            v_cash = st.number_input("現預金", value=45573502, key="v_cash")
            v_ins = st.number_input("生命保険金", value=3651514, key="v_ins")
            v_others = st.number_input("その他", value=1662687, key="v_others")
            v_gift_3y = st.number_input("相続前贈与（3〜7年）", value=0, key="v_gift_3y")
        with col_c:
            st.write("#### 📉 債務・葬式")
            v_debt = st.number_input("債務", value=322179, key="v_debt")
            v_funeral = st.number_input("葬式費用", value=41401, key="v_funeral")

        # -- 計算ロジック（小規模宅地・非課税枠） --
        st_count = heir_count + (1 if has_spouse else 0)
        a_lim, b_lim, c_lim = d(330), d(400), d(200)
        a_app = min(d(a_home), a_lim); b_app = min(d(a_biz), b_lim)
        used_r = (a_app / a_lim) + (b_app / b_lim)
        c_app = min(d(a_rent), c_lim * max(d(0), d(1) - used_r))
        red_home = (d(v_home) / d(max(1, a_home))) * a_app * d("0.8")
        red_biz = (d(v_biz) / d(max(1, a_biz))) * b_app * d("0.8")
        red_rent = (d(v_rent) / d(max(1, a_rent))) * c_app * d("0.5")
        total_red = red_home + red_biz + red_rent
        land_eval = d(v_home) + d(v_biz) + d(v_rent) - total_red
        ins_ded = min(d(v_ins), d(5000000) * d(st_count))
        pure_as = land_eval + d(v_build) + d(v_stock) + d(v_cash) + d(v_ins) + d(v_others)
        tax_p = pure_as - ins_ded - d(v_debt) - d(v_funeral) + d(v_gift_3y)
        basic_1 = d(30000000) + (d(6000000) * d(st_count))
        taxable_1 = max(d(0), tax_p - basic_1)
        total_tax_1 = SupremeLegacyEngine.get_tax(taxable_1, has_spouse, heirs_info)

    # -- TAB 3: 一次相続明細 --
    with tabs[2]:
        add_print_button("3. 一次相続明細")
        st.subheader("一次相続：計算明細")
        df1 = pd.DataFrame([
            ["1", "不動産評価（特例適用後）", f"{int(land_eval):,}", f"特例減額: {int(total_red):,}"],
            ["2", "建物評価額", f"{int(v_build):,}", ""],
            ["3", "有価証券", f"{int(v_stock):,}", ""],
            ["4", "現預金", f"{int(v_cash):,}", ""],
            ["5", "生命保険金(控除後)", f"{int(max(0, d(v_ins)-ins_ded)):,}", f"非課税枠: {int(ins_ded):,}"],
            ["6", "その他財産", f"{int(v_others):,}", ""],
            ["7", "債務および葬式費用", f"△{int(v_debt + v_funeral):,}", ""],
            ["8", "生前贈与加算", f"{int(v_gift_3y):,}", ""],
            ["9", "【課税価格合計】", f"{int(tax_p):,}", ""],
            ["10", "基礎控除額", f"△{int(basic_1):,}", f"3000万+(600万×{st_count})"],
            ["11", "課税遺産総額", f"{int(taxable_1):,}", ""],
            ["12", "【相続税の総額】", f"{int(total_tax_1):,}", ""],
        ], columns=["No", "項目", "金額", "備考"])
        st.table(df1)

    # -- TAB 4: 二次相続明細 --
    with tabs[3]:
        add_print_button("4. 二次相続明細")
        st.subheader("二次相続：計算明細予測")
        ratio_s = d("0.5") # デフォルト50%配分
        acq_s_1 = tax_p * ratio_s
        limit_s = max(d(160000000), taxable_1 * ratio_s)
        tax_s_1 = d(0) if acq_s_1 <= limit_s else (total_tax_1 * ratio_s * d("0.5"))
        net_acq_s = acq_s_1 - tax_s_1
        
        s_own = d(st.session_state.get("in_s_own", 50000000))
        s_spend_total = d(st.session_state.get("in_s_spend", 5000000)) * d(st.session_state.get("in_interval", 10))
        tax_p_2 = max(d(0), net_acq_s + s_own - s_spend_total)
        
        child_only = [h for h in heirs_info if h['type'] in ["子", "孫（養子含む）"]]
        c_count_2 = len(child_only) if child_only else heir_count
        basic_2 = d(30000000) + (d(6000000) * d(c_count_2))
        taxable_2 = max(d(0), tax_p_2 - basic_2)
        total_tax_2 = SupremeLegacyEngine.get_tax(taxable_2, False, child_only if child_only else heirs_info)
        
        df2 = pd.DataFrame([
            ["1", "一次からの純承継分", f"{int(net_acq_s):,}", "配偶者控除適用後"],
            ["2", "配偶者固有財産", f"{int(s_own):,}", ""],
            ["3", "生活費・支出等控除", f"△{int(s_spend_total):,}", ""],
            ["4", "【二次相続 課税価格】", f"{int(tax_p_2):,}", ""],
            ["5", "二次基礎控除額", f"△{int(basic_2):,}", f"相続人{c_count_2}名"],
            ["6", "【二次相続税 総額】", f"{int(total_tax_2):,}", ""],
        ], columns=["No", "項目", "金額", "備考"])
        st.table(df2)

    # -- TAB 5: 二次推移予測 --
    with tabs[4]:
        add_print_button("5. 二次推移予測")
        st.subheader("二次推移パラメータ設定")
        cp1, cp2 = st.columns(2)
        cp1.number_input("配偶者の固有財産", value=50000000, key="in_s_own")
        cp1.slider("二次までの想定期間(年)", 0, 20, 10, key="in_interval")
        cp2.number_input("年間生活費・支出(減価)", value=5000000, key="in_s_spend")

    # -- TAB 6: 精密分析結果 --
    with tabs[5]:
        add_print_button("6. 精密分析結果")
        st.subheader("配偶者取得割合別の税額推移分析")
        sim_results = []
        for i in range(0, 101, 10):
            ratio = d(i) / d(100)
            acq_s = tax_p * ratio
            limit_s = max(d(160000000), taxable_1 * ratio)
            t_s_1 = d(0) if acq_s <= limit_s else (total_tax_1 * ratio * d("0.5"))
            t1 = total_tax_1 - (total_tax_1 * ratio) + t_s_1
            net_s = acq_s - t_s_1
            tp2 = max(d(0), net_s + s_own - s_spend_total)
            t2 = SupremeLegacyEngine.get_tax(max(0, tp2 - basic_2), False, child_only if child_only else heirs_info)
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
        fig.update_layout(barmode='stack', title="税額最適化シミュレーション", xaxis_title="配偶者配分(%)", yaxis_title="税額(円)")
        st.plotly_chart(fig, use_container_width=True)
        
        st.subheader("数値データ詳細")
        st.table(df_sim.style.format({"一次相続税額": "{:,}", "二次相続税額": "{:,}", "合計納税額": "{:,}"}))

        # --- 遺留分シミュレーション（完全復元） ---
        st.divider()
        st.subheader("⚠️ 遺留分侵害額の確認（課税価格ベース）")
        s_r, h_shares = SupremeLegacyEngine.get_legal_shares(has_spouse, heirs_info)
        # 遺留分率は原則1/2（直系尊属のみの場合は1/3）
        iryu_total_ratio = d("0.333") if all(h['type'] == "親" for h in heirs_info) else d("0.5")
        
        iryu_data = []
        if has_spouse:
            iryu_data.append({
                "相続人": "配偶者", 
                "法定相続分": f"{float(s_r)*100:.1f}%", 
                "遺留分額": f"{int(tax_p * s_r * iryu_total_ratio):,}円"
            })
        for i, (h, share) in enumerate(zip(heirs_info, h_shares)):
            # 兄弟姉妹には遺留分がない
            if h['type'] in ["兄弟姉妹（全血）", "兄弟姉妹（半血）"]:
                val = "（権利なし）"
            else:
                val = f"{int(tax_p * share * iryu_total_ratio):,}円"
            iryu_data.append({
                "相続人": f"相続人{i+1}({h['type']})", 
                "法定相続分": f"{float(share)*100:.1f}%", 
                "遺留分額": val
            })
        st.table(pd.DataFrame(iryu_data))

        # --- 監査証跡（ネイビー＆ゴールド） ---
        st.divider()
        st.markdown(f"""
        <div style="background-color: #f9f9f9; border: 2px solid #c5a059; padding: 20px; border-radius: 5px;">
            <p style="color: #1f2c4d; font-weight: bold; margin-bottom: 10px;">🛡️ 山根会計 監査証跡エビデンス</p>
            <p style="font-size: 0.9em; line-height: 1.6;">
                本報告書は、山根会計 専売システム(SUMMIT Core v31.15)により出力されました。<br>
                実行担当者: 川東 / 監査: 擬似・オーナー監査官<br>
                計算根拠: 相続税法第15条（基礎控除）、第16条（総額計算）、及び租税特別措置法第69条の4（小規模宅地特例）に基づく。
            </p>
        </div>
        """, unsafe_allow_html=True)
