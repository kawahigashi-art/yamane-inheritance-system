from __future__ import annotations

# =========================================================
# 0. Imports / Page Config
# =========================================================
from dataclasses import dataclass
from datetime import date, timedelta
from decimal import Decimal, ROUND_HALF_UP
from io import BytesIO
from numbers import Number
from typing import Any

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

st.set_page_config(page_title="SUMMIT v31.16 PRO", layout="wide")

# =========================================================
# 1. Constants
# =========================================================
APP_TITLE = "山根会計 専売システム"
APP_LOGIN_USER_LABEL = "ログイン: 川東"
APP_PASSWORD = "yamane777"  # TODO: st.secrets へ移行
EXCEL_FILE_NAME = "相続シミュレーション.xlsx"
EXCEL_TITLE = "山根会計 相続税シミュレーション資料"

COLOR_NAVY = "1f2c4d"
COLOR_GOLD = "c5a059"
COLOR_RED = "a61d24"

SMALL_SCALE_HOME_LIMIT = Decimal("330")
SMALL_SCALE_BUSINESS_LIMIT = Decimal("400")
SMALL_SCALE_RENT_LIMIT = Decimal("200")
SMALL_SCALE_HOME_RATE = Decimal("0.8")
SMALL_SCALE_BUSINESS_RATE = Decimal("0.8")
SMALL_SCALE_RENT_RATE = Decimal("0.5")

SPOUSE_TAX_REDUCTION_REFERENCE = Decimal("160000000")
LIFE_INSURANCE_EXEMPT_PER_HEIR = Decimal("5000000")
BASIC_DEDUCTION_BASE = Decimal("30000000")
BASIC_DEDUCTION_PER_HEIR = Decimal("6000000")
PERCENT_DENOMINATOR = Decimal("100")
TWO_TENTHS_SURTAX_RATE = Decimal("0.2")
MAX_INSURANCE_RECIPIENT_ROWS = 5
MAX_GIFT_RECORD_ROWS = 10
ANNUAL_GIFT_LOOKBACK_YEARS = 7  # 概算。制度経過措置の精密判定は今後拡張
SEISAN_ANNUAL_BASIC_EXEMPTION = Decimal("1100000")
GIFT_TYPE_ANNUAL = "暦年課税"
GIFT_TYPE_SEISAN = "相続時精算課税"
GIFT_TYPE_OPTIONS = [GIFT_TYPE_ANNUAL, GIFT_TYPE_SEISAN]

SMALL_SCALE_STATUS_APPLICABLE = "適用可"
SMALL_SCALE_STATUS_REVIEW = "要確認"
SMALL_SCALE_STATUS_NOT_APPLICABLE = "適用不可"

LAND_CATEGORY_HOME = "特定居住用"
LAND_CATEGORY_BUSINESS = "特定事業用"
LAND_CATEGORY_RENTAL = "貸付事業用"

HEIR_TYPE_CHILD = "子"
HEIR_TYPE_GRANDCHILD = "孫（養子含む）"
HEIR_TYPE_PARENT = "親"
HEIR_TYPE_FULL_SIBLING = "兄弟姉妹（全血）"
HEIR_TYPE_HALF_SIBLING = "兄弟姉妹（半血）"

HEIR_TYPE_OPTIONS = [
    HEIR_TYPE_CHILD,
    HEIR_TYPE_GRANDCHILD,
    HEIR_TYPE_PARENT,
    HEIR_TYPE_FULL_SIBLING,
    HEIR_TYPE_HALF_SIBLING,
]

TAB_LABELS = [
    " 👥  1.基本構成",
    " 💰  2.一次財産詳細",
    " 📑  3.一次相続明細",
    " 📑  4.二次相続明細",
    " ⏳  5.二次推移予測",
    " 📊  6.精密分析結果",
]

# =========================================================
# 2. Data Models
# =========================================================
@dataclass
class PrimaryInputs:
    heir_count: int
    has_spouse: bool
    heirs_info: list[dict[str, str]]
    date_of_death: date
    v_home: int
    a_home: int
    v_biz: int
    a_biz: int
    v_rent: int
    a_rent: int
    small_scale_inputs: dict[str, SmallScaleInput]
    v_build: int
    v_stock: int
    v_cash: int
    v_ins: int
    insurance_entries: list[InsuranceRecipientInput]
    v_others: int
    v_debt: int
    v_funeral: int
    gift_records: list[GiftRecord]


@dataclass
class SecondaryInputs:
    spouse_acquisition_pct: int
    s_own: int
    annual_spend: int
    interval_years: int
    use_individual_allocations: bool
    actual_acquisition_inputs: list[int]


@dataclass
class InsuranceRecipientInput:
    recipient_name: str
    recipient_type: str
    amount: int
    is_statutory_heir: bool
    is_two_tenths_target: bool


@dataclass
class GiftRecord:
    gift_date: date
    recipient_name: str
    recipient_type: str
    amount: int
    tax_type: str


@dataclass
class GiftComputationRecord:
    gift_date: date
    recipient_name: str
    recipient_type: str
    tax_type: str
    amount: Decimal
    calendar_year: int
    is_addback_target: bool
    addback_amount: Decimal
    reason: str


@dataclass
class SmallScaleInput:
    category: str
    acquirer_name: str
    apply_special_rule: bool
    is_spouse_acquirer: bool = False
    is_cohabiting_heir: bool = False
    is_no_house_heir: bool = False
    will_hold_until_filing: bool = False
    will_reside_until_filing: bool = False
    is_business_successor: bool = False
    will_continue_business: bool = False
    will_continue_rental: bool = False


@dataclass
class SmallScaleLandRecord:
    category: str
    land_name: str
    acquirer_name: str
    status: str
    reason: str
    original_value: Decimal
    area_sqm: Decimal
    eligible_area_sqm: Decimal
    reduction_rate: Decimal
    reduction_amount: Decimal


@dataclass
class HeirTaxRecord:
    name: str
    heir_type: str
    legal_share: Decimal
    actual_share: Decimal
    input_acquisition_amount: Decimal
    normalized_acquisition_amount: Decimal
    insurance_gross: Decimal
    insurance_nontaxable: Decimal
    insurance_taxable: Decimal
    annual_gift_addback: Decimal
    seisan_gift_addback: Decimal
    base_taxable_price: Decimal
    taxable_price: Decimal
    preliminary_tax: Decimal
    two_tenths_surtax: Decimal
    spouse_credit: Decimal
    final_tax: Decimal
    is_two_tenths_target: bool


@dataclass
class PrimaryResult:
    st_count: int
    land_eval: Decimal
    total_red: Decimal
    small_scale_records: list[SmallScaleLandRecord]
    ins_ded: Decimal
    pure_as: Decimal
    tax_p: Decimal
    basic_1: Decimal
    taxable_1: Decimal
    total_tax_1: Decimal
    spouse_legal_share: Decimal
    heir_legal_shares: list[Decimal]
    spouse_actual_share: Decimal
    spouse_actual_taxable_price: Decimal
    spouse_tax_limit: Decimal
    total_final_tax: Decimal
    total_insurance_gross: Decimal
    total_insurance_nontaxable: Decimal
    total_insurance_taxable: Decimal
    total_annual_gift_addback: Decimal
    total_seisan_gift_addback: Decimal
    gift_detail_records: list[GiftComputationRecord]
    heir_tax_records: list[HeirTaxRecord]


@dataclass
class SecondaryResult:
    ratio_s: Decimal
    acq_s_1: Decimal
    limit_s: Decimal
    tax_s_1: Decimal
    net_acq_s: Decimal
    s_own: Decimal
    s_spend_total: Decimal
    tax_p_2: Decimal
    c_count_2: int
    basic_2: Decimal
    taxable_2: Decimal
    total_tax_2: Decimal
    child_only: list[dict[str, str]]


# =========================================================
# 3. Utility Functions
# =========================================================
def to_d(val: Any) -> Decimal:
    return Decimal(str(val))


def quantize_yen(val: Decimal) -> Decimal:
    return val.quantize(Decimal("1"), ROUND_HALF_UP)


def fmt_int(val: int | Decimal) -> str:
    return f"{int(val):,}"


def fmt_pct(val: Decimal) -> str:
    return f"{(val * PERCENT_DENOMINATOR):.1f}%"


def build_heir_labels(has_spouse: bool, heirs_info: list[dict[str, str]]) -> list[tuple[str, str]]:
    labels: list[tuple[str, str]] = []
    if has_spouse:
        labels.append(("配偶者", "配偶者"))
    for idx, heir in enumerate(heirs_info, start=1):
        labels.append((f"相続人{idx}", heir["type"]))
    return labels


def build_recipient_options(has_spouse: bool, heirs_info: list[dict[str, Any]]) -> list[tuple[str, str, bool, bool]]:
    options: list[tuple[str, str, bool, bool]] = []
    if has_spouse:
        options.append(("配偶者", "配偶者", True, False))
    for idx, heir in enumerate(heirs_info, start=1):
        label = f"相続人{idx}"
        heir_type = heir["type"]
        is_substitute = bool(heir.get("is_substitute", False))
        is_statutory_heir = True
        is_two_tenths_target = is_two_tenths_surtax_target(heir_type, is_substitute)
        options.append((label, heir_type, is_statutory_heir, is_two_tenths_target))
    return options


def normalize_ratio(total: Decimal) -> Decimal:
    return Decimal("0") if total <= 0 else total


def build_gift_recipient_options(has_spouse: bool, heirs_info: list[dict[str, Any]]) -> list[tuple[str, str]]:
    options: list[tuple[str, str]] = []
    if has_spouse:
        options.append(("配偶者", "配偶者"))
    for idx, heir in enumerate(heirs_info, start=1):
        options.append((f"相続人{idx}", heir["type"]))
    return options


def calculate_annual_gift_addback(
    gift_records: list[GiftRecord],
    date_of_death: date,
    labels: list[tuple[str, str]],
) -> tuple[list[Decimal], list[GiftComputationRecord]]:
    recipient_map = {label: idx for idx, (label, _) in enumerate(labels)}
    addbacks = [Decimal("0")] * len(labels)
    detail_records: list[GiftComputationRecord] = []
    threshold_date = date_of_death - timedelta(days=365 * ANNUAL_GIFT_LOOKBACK_YEARS)

    for gift in gift_records:
        if gift.tax_type != GIFT_TYPE_ANNUAL:
            continue
        amount = to_d(max(0, gift.amount))
        is_target = threshold_date <= gift.gift_date <= date_of_death and amount > 0
        addback_amount = amount if is_target else Decimal("0")
        reason = "加算対象期間内" if is_target else "加算対象期間外または相続開始日後"
        if is_target and gift.recipient_name in recipient_map:
            addbacks[recipient_map[gift.recipient_name]] += addback_amount
        detail_records.append(
            GiftComputationRecord(
                gift_date=gift.gift_date,
                recipient_name=gift.recipient_name,
                recipient_type=gift.recipient_type,
                tax_type=gift.tax_type,
                amount=amount,
                calendar_year=gift.gift_date.year,
                is_addback_target=is_target,
                addback_amount=addback_amount,
                reason=reason,
            )
        )
    return addbacks, detail_records


def calculate_seisan_addback(
    gift_records: list[GiftRecord],
    date_of_death: date,
    labels: list[tuple[str, str]],
) -> tuple[list[Decimal], list[GiftComputationRecord]]:
    recipient_map = {label: idx for idx, (label, _) in enumerate(labels)}
    addbacks = [Decimal("0")] * len(labels)
    detail_records: list[GiftComputationRecord] = []
    grouped: dict[tuple[str, int], list[GiftRecord]] = {}

    for gift in gift_records:
        if gift.tax_type != GIFT_TYPE_SEISAN:
            continue
        if gift.gift_date > date_of_death:
            detail_records.append(
                GiftComputationRecord(
                    gift_date=gift.gift_date,
                    recipient_name=gift.recipient_name,
                    recipient_type=gift.recipient_type,
                    tax_type=gift.tax_type,
                    amount=to_d(max(0, gift.amount)),
                    calendar_year=gift.gift_date.year,
                    is_addback_target=False,
                    addback_amount=Decimal("0"),
                    reason="相続開始日後のため対象外",
                )
            )
            continue
        grouped.setdefault((gift.recipient_name, gift.gift_date.year), []).append(gift)

    for (recipient_name, calendar_year), records in grouped.items():
        total_amount = sum((to_d(max(0, record.amount)) for record in records), Decimal("0"))
        total_addback = max(Decimal("0"), total_amount - SEISAN_ANNUAL_BASIC_EXEMPTION)
        running_allocated = Decimal("0")
        for idx, record in enumerate(records, start=1):
            amount = to_d(max(0, record.amount))
            if idx == len(records):
                addback_amount = max(Decimal("0"), total_addback - running_allocated)
            elif total_amount <= 0 or total_addback <= 0:
                addback_amount = Decimal("0")
            else:
                addback_amount = quantize_yen(total_addback * amount / total_amount)
                running_allocated += addback_amount
            if recipient_name in recipient_map:
                addbacks[recipient_map[recipient_name]] += addback_amount
            reason = "年110万円控除後の戻し額" if total_addback > 0 else "年110万円控除内"
            detail_records.append(
                GiftComputationRecord(
                    gift_date=record.gift_date,
                    recipient_name=record.recipient_name,
                    recipient_type=record.recipient_type,
                    tax_type=record.tax_type,
                    amount=amount,
                    calendar_year=calendar_year,
                    is_addback_target=total_addback > 0,
                    addback_amount=addback_amount,
                    reason=reason,
                )
            )
    return addbacks, detail_records


def calculate_gift_addbacks(
    gift_records: list[GiftRecord],
    date_of_death: date,
    labels: list[tuple[str, str]],
) -> tuple[list[Decimal], list[Decimal], list[GiftComputationRecord]]:
    annual_addbacks, annual_details = calculate_annual_gift_addback(gift_records, date_of_death, labels)
    seisan_addbacks, seisan_details = calculate_seisan_addback(gift_records, date_of_death, labels)
    detail_records = sorted(annual_details + seisan_details, key=lambda x: (x.gift_date, x.recipient_name, x.tax_type))
    return annual_addbacks, seisan_addbacks, detail_records


def is_two_tenths_surtax_target(heir_type: str, is_substitute: bool = False) -> bool:
    if heir_type == "配偶者":
        return False
    if heir_type == HEIR_TYPE_GRANDCHILD:
        return not is_substitute
    if heir_type in [HEIR_TYPE_FULL_SIBLING, HEIR_TYPE_HALF_SIBLING]:
        return True
    return False


def authenticate_user() -> bool:
    if st.session_state.get("password_correct"):
        return True

    st.title(f" 🔐  {APP_TITLE}")
    pwd = st.text_input("アクセスパスワード", type="password")
    if st.button("ログイン"):
        if pwd == APP_PASSWORD:
            st.session_state["password_correct"] = True
            st.rerun()
        else:
            st.error("パスワードが正しくありません。")
    return False


def inject_print_css() -> None:
    st.markdown(
        """
        <style>
        @media print {
            section[data-testid="stSidebar"], header, .stButton, div[data-testid="stToolbar"], footer {
                display: none !important;
            }
            .main .block-container { padding: 0 !important; margin: 0 !important; }
        }
        .print-btn-container { display: flex; justify-content: flex-end; margin-bottom: 20px; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def add_print_button(tab_name: str) -> None:
    html_code = f"""
        <div class="print-btn-container">
            <button onclick="window.parent.print()" style="
                background-color: #{COLOR_NAVY}; color: #{COLOR_GOLD}; border: 2px solid #{COLOR_GOLD};
                padding: 10px 20px; border-radius: 5px; cursor: pointer; font-weight: bold;
            ">
                🖨️ 「{tab_name}」を印刷 / PDF保存
            </button>
        </div>
    """
    components.html(html_code, height=60)


# =========================================================
# 4. Tax Logic
# =========================================================
class SupremeLegacyEngine:
    @staticmethod
    def get_legal_shares(has_spouse: bool, heirs_info: list[dict[str, str]]) -> tuple[Decimal, list[Decimal]]:
        shares: list[Decimal] = []
        has_child = any(h["type"] in [HEIR_TYPE_CHILD, HEIR_TYPE_GRANDCHILD] for h in heirs_info)
        has_parent = any(h["type"] == HEIR_TYPE_PARENT for h in heirs_info) if not has_child else False
        has_sibling = (
            any(h["type"] in [HEIR_TYPE_FULL_SIBLING, HEIR_TYPE_HALF_SIBLING] for h in heirs_info)
            if not (has_child or has_parent)
            else False
        )

        if has_child:
            s_ratio = Decimal("0.5") if has_spouse else Decimal("0")
            h_total_ratio = Decimal("0.5") if has_spouse else Decimal("1.0")
            children = [h for h in heirs_info if h["type"] in [HEIR_TYPE_CHILD, HEIR_TYPE_GRANDCHILD]]
            per_h = h_total_ratio / to_d(len(children))
            for h in heirs_info:
                shares.append(per_h if h["type"] in [HEIR_TYPE_CHILD, HEIR_TYPE_GRANDCHILD] else Decimal("0"))
        elif has_parent:
            s_ratio = Decimal("0.6666666666666667") if has_spouse else Decimal("0")
            h_total_ratio = Decimal("0.3333333333333333") if has_spouse else Decimal("1.0")
            parents = [h for h in heirs_info if h["type"] == HEIR_TYPE_PARENT]
            per_h = h_total_ratio / to_d(len(parents))
            for h in heirs_info:
                shares.append(per_h if h["type"] == HEIR_TYPE_PARENT else Decimal("0"))
        elif has_sibling:
            s_ratio = Decimal("0.75") if has_spouse else Decimal("0")
            h_total_ratio = Decimal("0.25") if has_spouse else Decimal("1.0")
            weight_sum = Decimal("0")
            for h in heirs_info:
                if h["type"] == HEIR_TYPE_FULL_SIBLING:
                    weight_sum += Decimal("1")
                elif h["type"] == HEIR_TYPE_HALF_SIBLING:
                    weight_sum += Decimal("0.5")
            unit_share = h_total_ratio / weight_sum if weight_sum > 0 else Decimal("0")
            for h in heirs_info:
                if h["type"] == HEIR_TYPE_FULL_SIBLING:
                    shares.append(unit_share)
                elif h["type"] == HEIR_TYPE_HALF_SIBLING:
                    shares.append(unit_share * Decimal("0.5"))
                else:
                    shares.append(Decimal("0"))
        else:
            s_ratio = Decimal("1") if has_spouse else Decimal("0")
            shares = [Decimal("0")] * len(heirs_info)
        return s_ratio, shares

    @staticmethod
    def bracket_calc(amount: Decimal) -> Decimal:
        if amount <= 10000000:
            return amount * Decimal("0.10")
        if amount <= 30000000:
            return amount * Decimal("0.15") - Decimal("500000")
        if amount <= 50000000:
            return amount * Decimal("0.20") - Decimal("2000000")
        if amount <= 100000000:
            return amount * Decimal("0.30") - Decimal("7000000")
        if amount <= 200000000:
            return amount * Decimal("0.40") - Decimal("17000000")
        if amount <= 300000000:
            return amount * Decimal("0.45") - Decimal("27000000")
        if amount <= 600000000:
            return amount * Decimal("0.50") - Decimal("42000000")
        return amount * Decimal("0.55") - Decimal("72000000")

    @staticmethod
    def get_tax(taxable_amt: Decimal, has_spouse: bool, heirs_info: list[dict[str, str]]) -> Decimal:
        if taxable_amt <= 0:
            return Decimal("0")
        s_ratio, h_shares = SupremeLegacyEngine.get_legal_shares(has_spouse, heirs_info)
        total_tax = Decimal("0")
        if has_spouse:
            total_tax += SupremeLegacyEngine.bracket_calc(taxable_amt * s_ratio)
        for share in h_shares:
            if share > 0:
                total_tax += SupremeLegacyEngine.bracket_calc(taxable_amt * share)
        return quantize_yen(total_tax)


def allocate_actual_shares(has_spouse: bool, heirs_info: list[dict[str, str]], spouse_acquisition_pct: int) -> list[Decimal]:
    spouse_share = Decimal("0")
    if has_spouse:
        spouse_share = to_d(spouse_acquisition_pct) / PERCENT_DENOMINATOR
        spouse_share = max(Decimal("0"), min(Decimal("1"), spouse_share))

    non_spouse_count = len(heirs_info)
    if not has_spouse:
        spouse_share = Decimal("0")
    if non_spouse_count <= 0:
        return [Decimal("1")] if has_spouse else []

    remaining = Decimal("1") - spouse_share
    per_non_spouse = remaining / to_d(non_spouse_count)
    shares = [spouse_share] if has_spouse else []
    shares.extend([per_non_spouse] * non_spouse_count)
    return shares


def normalize_amounts_to_total(total_amount: Decimal, desired_amounts: list[Decimal], fallback_shares: list[Decimal]) -> list[Decimal]:
    if total_amount <= 0:
        return [Decimal("0")] * len(fallback_shares)

    sanitized = [max(Decimal("0"), amount) for amount in desired_amounts[: len(fallback_shares)]]
    while len(sanitized) < len(fallback_shares):
        sanitized.append(Decimal("0"))

    desired_total = sum(sanitized, Decimal("0"))
    normalized: list[Decimal] = []

    if desired_total <= 0:
        running_total = Decimal("0")
        for idx, share in enumerate(fallback_shares, start=1):
            if idx == len(fallback_shares):
                amount = total_amount - running_total
            else:
                amount = quantize_yen(total_amount * share)
                running_total += amount
            normalized.append(max(Decimal("0"), amount))
        return normalized

    running_total = Decimal("0")
    for idx, amount in enumerate(sanitized, start=1):
        if idx == len(sanitized):
            normalized_amount = total_amount - running_total
        else:
            normalized_amount = quantize_yen(total_amount * amount / desired_total)
            running_total += normalized_amount
        normalized.append(max(Decimal("0"), normalized_amount))
    return normalized


def normalize_actual_acquisition_plan(
    total_taxable_price: Decimal,
    desired_amounts: list[int],
    fallback_shares: list[Decimal],
) -> tuple[list[Decimal], list[Decimal]]:
    normalized_amounts = normalize_amounts_to_total(
        total_amount=total_taxable_price,
        desired_amounts=[to_d(max(0, amount)) for amount in desired_amounts],
        fallback_shares=fallback_shares,
    )
    if total_taxable_price <= 0:
        return [Decimal("0")] * len(fallback_shares), normalized_amounts
    actual_shares = [amount / total_taxable_price for amount in normalized_amounts]
    return actual_shares, normalized_amounts



def allocate_taxable_prices(total_taxable_price: Decimal, actual_shares: list[Decimal]) -> list[Decimal]:
    return [quantize_yen(total_taxable_price * share) for share in actual_shares]


def allocate_preliminary_taxes(total_tax: Decimal, taxable_prices: list[Decimal], total_taxable_price: Decimal) -> list[Decimal]:
    if total_tax <= 0 or total_taxable_price <= 0:
        return [Decimal("0")] * len(taxable_prices)
    return [quantize_yen(total_tax * taxable / total_taxable_price) for taxable in taxable_prices]


def normalize_insurance_entries(total_insurance: int, entries: list[InsuranceRecipientInput]) -> list[InsuranceRecipientInput]:
    valid_entries = [entry for entry in entries if entry.amount > 0]
    if total_insurance <= 0 or not valid_entries:
        return []
    entry_total = sum(entry.amount for entry in valid_entries)
    if entry_total <= 0:
        return []
    normalized: list[InsuranceRecipientInput] = []
    cumulative = 0
    for idx, entry in enumerate(valid_entries, start=1):
        if idx == len(valid_entries):
            normalized_amount = total_insurance - cumulative
        else:
            normalized_amount = int(round(total_insurance * entry.amount / entry_total))
            cumulative += normalized_amount
        normalized.append(
            InsuranceRecipientInput(
                recipient_name=entry.recipient_name,
                recipient_type=entry.recipient_type,
                amount=max(0, normalized_amount),
                is_statutory_heir=entry.is_statutory_heir,
                is_two_tenths_target=entry.is_two_tenths_target,
            )
        )
    return normalized


def allocate_insurance_by_recipient(
    total_insurance: int,
    entries: list[InsuranceRecipientInput],
    labels: list[tuple[str, str]],
    statutory_heir_count: int,
) -> tuple[list[Decimal], list[Decimal], list[Decimal]]:
    grosses = [Decimal("0")] * len(labels)
    nontaxables = [Decimal("0")] * len(labels)
    taxables = [Decimal("0")] * len(labels)
    normalized_entries = normalize_insurance_entries(total_insurance, entries)
    if total_insurance <= 0:
        return grosses, nontaxables, taxables
    if not normalized_entries and labels:
        fallback_label, fallback_type = labels[0]
        normalized_entries = [
            InsuranceRecipientInput(
                recipient_name=fallback_label,
                recipient_type=fallback_type,
                amount=total_insurance,
                is_statutory_heir=True,
                is_two_tenths_target=is_two_tenths_surtax_target(fallback_type),
            )
        ]

    label_to_index = {label: idx for idx, (label, _) in enumerate(labels)}
    for entry in normalized_entries:
        idx = label_to_index.get(entry.recipient_name)
        if idx is not None:
            grosses[idx] += to_d(entry.amount)

    insurance_limit = LIFE_INSURANCE_EXEMPT_PER_HEIR * to_d(statutory_heir_count)
    eligible_indices = [label_to_index[entry.recipient_name] for entry in normalized_entries if entry.is_statutory_heir and entry.recipient_name in label_to_index]
    eligible_total = sum((grosses[idx] for idx in eligible_indices), Decimal("0"))

    if insurance_limit > 0 and eligible_total > 0:
        remaining_limit = insurance_limit
        for position, idx in enumerate(eligible_indices, start=1):
            gross = grosses[idx]
            if position == len(eligible_indices):
                exempt_amount = min(gross, remaining_limit)
            else:
                exempt_amount = min(gross, quantize_yen(insurance_limit * gross / eligible_total))
                remaining_limit -= exempt_amount
            nontaxables[idx] += exempt_amount

    for idx, gross in enumerate(grosses):
        taxables[idx] = max(Decimal("0"), gross - nontaxables[idx])
    return grosses, nontaxables, taxables


def apply_two_tenths_surtax(preliminary_taxes: list[Decimal], two_tenths_targets: list[bool]) -> tuple[list[Decimal], list[Decimal]]:
    surtax_amounts: list[Decimal] = []
    adjusted_taxes: list[Decimal] = []
    for tax, is_target in zip(preliminary_taxes, two_tenths_targets):
        surtax = quantize_yen(tax * TWO_TENTHS_SURTAX_RATE) if is_target and tax > 0 else Decimal("0")
        surtax_amounts.append(surtax)
        adjusted_taxes.append(tax + surtax)
    return surtax_amounts, adjusted_taxes


# =========================================================
# 5. Special Rule Logic
# =========================================================
def determine_small_scale_land_eligibility(rule: SmallScaleInput) -> tuple[str, str]:
    """小規模宅地等の特例の概算判定を行う。
    今回は保守的運用として、要件が不足する場合は「要確認」または「適用不可」とする。
    """
    if not rule.apply_special_rule:
        return SMALL_SCALE_STATUS_NOT_APPLICABLE, "特例適用を選択していません"

    if rule.category == LAND_CATEGORY_HOME:
        if rule.is_spouse_acquirer:
            return SMALL_SCALE_STATUS_APPLICABLE, "配偶者取得として判定"
        if rule.is_cohabiting_heir and rule.will_hold_until_filing and rule.will_reside_until_filing:
            return SMALL_SCALE_STATUS_APPLICABLE, "同居親族・継続保有・継続居住を充足"
        if rule.is_no_house_heir and rule.will_hold_until_filing:
            return SMALL_SCALE_STATUS_APPLICABLE, "家なき子・継続保有を充足"
        if rule.is_cohabiting_heir or rule.is_no_house_heir:
            return SMALL_SCALE_STATUS_REVIEW, "居住継続または保有継続の確認が未了"
        return SMALL_SCALE_STATUS_NOT_APPLICABLE, "居住用の主要要件を満たしていません"

    if rule.category == LAND_CATEGORY_BUSINESS:
        if rule.is_business_successor and rule.will_continue_business and rule.will_hold_until_filing:
            return SMALL_SCALE_STATUS_APPLICABLE, "事業承継・継続事業・継続保有を充足"
        if rule.is_business_successor:
            return SMALL_SCALE_STATUS_REVIEW, "継続事業または継続保有の確認が未了"
        return SMALL_SCALE_STATUS_NOT_APPLICABLE, "事業承継者要件を満たしていません"

    if rule.category == LAND_CATEGORY_RENTAL:
        if rule.will_continue_rental and rule.will_hold_until_filing:
            return SMALL_SCALE_STATUS_APPLICABLE, "貸付継続・継続保有を充足"
        if rule.will_continue_rental or rule.will_hold_until_filing:
            return SMALL_SCALE_STATUS_REVIEW, "貸付継続または継続保有の確認が未了"
        return SMALL_SCALE_STATUS_NOT_APPLICABLE, "貸付事業継続要件を満たしていません"

    return SMALL_SCALE_STATUS_REVIEW, "用途判定が未確定です"


def calc_small_scale_land_reduction(value: int, area: int, category: str, status: str) -> tuple[Decimal, Decimal, Decimal]:
    area_d = to_d(max(area, 0))
    value_d = to_d(max(value, 0))
    if value_d <= 0 or area_d <= 0 or status != SMALL_SCALE_STATUS_APPLICABLE:
        return Decimal("0"), Decimal("0"), Decimal("0")

    if category == LAND_CATEGORY_HOME:
        eligible_area = min(area_d, SMALL_SCALE_HOME_LIMIT)
        rate = SMALL_SCALE_HOME_RATE
    elif category == LAND_CATEGORY_BUSINESS:
        eligible_area = min(area_d, SMALL_SCALE_BUSINESS_LIMIT)
        rate = SMALL_SCALE_BUSINESS_RATE
    else:
        eligible_area = min(area_d, SMALL_SCALE_RENT_LIMIT)
        rate = SMALL_SCALE_RENT_RATE

    reduction = quantize_yen((value_d / area_d) * eligible_area * rate)
    return eligible_area, rate, reduction


def calculate_small_scale_reduction(inputs: PrimaryInputs) -> tuple[Decimal, Decimal, list[SmallScaleLandRecord]]:
    records: list[SmallScaleLandRecord] = []
    total_reduction = Decimal("0")

    land_specs = [
        (LAND_CATEGORY_HOME, "特定居住用宅地", inputs.v_home, inputs.a_home),
        (LAND_CATEGORY_BUSINESS, "特定事業用宅地", inputs.v_biz, inputs.a_biz),
        (LAND_CATEGORY_RENTAL, "貸付事業用宅地", inputs.v_rent, inputs.a_rent),
    ]

    for category, land_name, value, area in land_specs:
        rule = inputs.small_scale_inputs.get(category)
        if rule is None:
            rule = SmallScaleInput(category=category, acquirer_name="未設定", apply_special_rule=False)
        status, reason = determine_small_scale_land_eligibility(rule)
        eligible_area, rate, reduction = calc_small_scale_land_reduction(value, area, category, status)
        total_reduction += reduction
        records.append(
            SmallScaleLandRecord(
                category=category,
                land_name=land_name,
                acquirer_name=rule.acquirer_name,
                status=status,
                reason=reason,
                original_value=to_d(value),
                area_sqm=to_d(area),
                eligible_area_sqm=eligible_area,
                reduction_rate=rate,
                reduction_amount=reduction,
            )
        )

    land_eval = to_d(inputs.v_home) + to_d(inputs.v_biz) + to_d(inputs.v_rent) - total_reduction
    return quantize_yen(total_reduction), quantize_yen(land_eval), records


def calculate_life_insurance_deduction(v_ins: int, heir_count: int) -> Decimal:
    return min(to_d(v_ins), LIFE_INSURANCE_EXEMPT_PER_HEIR * to_d(heir_count))


def apply_spouse_tax_credit(
    has_spouse: bool,
    total_taxable_price: Decimal,
    spouse_legal_share: Decimal,
    taxable_prices: list[Decimal],
    preliminary_taxes: list[Decimal],
) -> tuple[list[Decimal], Decimal, Decimal]:
    final_taxes = list(preliminary_taxes)
    if not has_spouse or not taxable_prices:
        return final_taxes, Decimal("0"), Decimal("0")

    spouse_actual_taxable = taxable_prices[0]
    spouse_tax_limit = max(SPOUSE_TAX_REDUCTION_REFERENCE, total_taxable_price * spouse_legal_share)
    if spouse_actual_taxable <= 0:
        return final_taxes, spouse_tax_limit, Decimal("0")

    if spouse_actual_taxable <= spouse_tax_limit:
        spouse_credit = final_taxes[0]
        final_taxes[0] = Decimal("0")
        return final_taxes, spouse_tax_limit, spouse_credit

    credit_ratio = spouse_tax_limit / spouse_actual_taxable
    spouse_credit = quantize_yen(preliminary_taxes[0] * credit_ratio)
    final_taxes[0] = max(Decimal("0"), preliminary_taxes[0] - spouse_credit)
    return final_taxes, spouse_tax_limit, spouse_credit


def build_heir_tax_records(
    labels: list[tuple[str, str]],
    legal_shares: list[Decimal],
    actual_shares: list[Decimal],
    input_acquisition_amounts: list[Decimal],
    normalized_acquisition_amounts: list[Decimal],
    insurance_grosses: list[Decimal],
    insurance_nontaxables: list[Decimal],
    insurance_taxables: list[Decimal],
    annual_gift_addbacks: list[Decimal],
    seisan_gift_addbacks: list[Decimal],
    base_taxable_prices: list[Decimal],
    taxable_prices: list[Decimal],
    preliminary_taxes: list[Decimal],
    surtax_amounts: list[Decimal],
    final_taxes: list[Decimal],
    spouse_credit: Decimal,
    two_tenths_targets: list[bool],
) -> list[HeirTaxRecord]:
    records: list[HeirTaxRecord] = []
    for idx, (label, heir_type) in enumerate(labels):
        current_spouse_credit = spouse_credit if idx == 0 and heir_type == "配偶者" else Decimal("0")
        records.append(
            HeirTaxRecord(
                name=label,
                heir_type=heir_type,
                legal_share=legal_shares[idx] if idx < len(legal_shares) else Decimal("0"),
                actual_share=actual_shares[idx] if idx < len(actual_shares) else Decimal("0"),
                input_acquisition_amount=input_acquisition_amounts[idx] if idx < len(input_acquisition_amounts) else Decimal("0"),
                normalized_acquisition_amount=normalized_acquisition_amounts[idx] if idx < len(normalized_acquisition_amounts) else Decimal("0"),
                insurance_gross=insurance_grosses[idx] if idx < len(insurance_grosses) else Decimal("0"),
                insurance_nontaxable=insurance_nontaxables[idx] if idx < len(insurance_nontaxables) else Decimal("0"),
                insurance_taxable=insurance_taxables[idx] if idx < len(insurance_taxables) else Decimal("0"),
                annual_gift_addback=annual_gift_addbacks[idx] if idx < len(annual_gift_addbacks) else Decimal("0"),
                seisan_gift_addback=seisan_gift_addbacks[idx] if idx < len(seisan_gift_addbacks) else Decimal("0"),
                base_taxable_price=base_taxable_prices[idx] if idx < len(base_taxable_prices) else Decimal("0"),
                taxable_price=taxable_prices[idx] if idx < len(taxable_prices) else Decimal("0"),
                preliminary_tax=preliminary_taxes[idx] if idx < len(preliminary_taxes) else Decimal("0"),
                two_tenths_surtax=surtax_amounts[idx] if idx < len(surtax_amounts) else Decimal("0"),
                spouse_credit=current_spouse_credit,
                final_tax=final_taxes[idx] if idx < len(final_taxes) else Decimal("0"),
                is_two_tenths_target=two_tenths_targets[idx] if idx < len(two_tenths_targets) else False,
            )
        )
    return records


def build_iryubun_reference(primary_inputs: PrimaryInputs, primary_result: PrimaryResult) -> pd.DataFrame:
    iryu_total_ratio = Decimal("0.333") if all(h["type"] == HEIR_TYPE_PARENT for h in primary_inputs.heirs_info) else Decimal("0.5")
    rows: list[dict[str, str]] = []
    if primary_inputs.has_spouse:
        rows.append(
            {
                "相続人": "配偶者",
                "法定相続分": f"{float(primary_result.spouse_legal_share) * 100:.1f}%",
                "遺留分額": f"{fmt_int(primary_result.tax_p * primary_result.spouse_legal_share * iryu_total_ratio)}円",
            }
        )
    for idx, (heir, share) in enumerate(zip(primary_inputs.heirs_info, primary_result.heir_legal_shares), start=1):
        if heir["type"] in [HEIR_TYPE_FULL_SIBLING, HEIR_TYPE_HALF_SIBLING]:
            iryubun_value = "（権利なし）"
        else:
            iryubun_value = f"{fmt_int(primary_result.tax_p * share * iryu_total_ratio)}円"
        rows.append(
            {
                "相続人": f"相続人{idx}({heir['type']})",
                "法定相続分": f"{float(share) * 100:.1f}%",
                "遺留分額": iryubun_value,
            }
        )
    return pd.DataFrame(rows)


# =========================================================
# 6. Simulation Logic
# =========================================================
def calculate_primary_inheritance(inputs: PrimaryInputs, secondary_inputs: SecondaryInputs) -> PrimaryResult:
    st_count = inputs.heir_count + (1 if inputs.has_spouse else 0)
    total_red, land_eval, small_scale_records = calculate_small_scale_reduction(inputs)
    ins_ded = calculate_life_insurance_deduction(inputs.v_ins, st_count)
    pure_as = land_eval + to_d(inputs.v_build) + to_d(inputs.v_stock) + to_d(inputs.v_cash) + to_d(inputs.v_ins) + to_d(inputs.v_others)
    basic_1 = BASIC_DEDUCTION_BASE + (BASIC_DEDUCTION_PER_HEIR * to_d(st_count))
    spouse_legal_share, heir_legal_shares = SupremeLegacyEngine.get_legal_shares(inputs.has_spouse, inputs.heirs_info)

    fallback_shares = allocate_actual_shares(inputs.has_spouse, inputs.heirs_info, secondary_inputs.spouse_acquisition_pct)
    labels = build_heir_labels(inputs.has_spouse, inputs.heirs_info)
    input_acquisition_amounts = [to_d(max(0, amount)) for amount in secondary_inputs.actual_acquisition_inputs]
    if not secondary_inputs.use_individual_allocations:
        input_acquisition_amounts = []
    annual_gift_addbacks, seisan_gift_addbacks, gift_detail_records = calculate_gift_addbacks(
        inputs.gift_records,
        inputs.date_of_death,
        labels,
    )
    total_annual_gift_addback = sum(annual_gift_addbacks, Decimal("0"))
    total_seisan_gift_addback = sum(seisan_gift_addbacks, Decimal("0"))
    tax_p = pure_as - ins_ded - to_d(inputs.v_debt) - to_d(inputs.v_funeral) + total_annual_gift_addback + total_seisan_gift_addback
    taxable_1 = max(Decimal("0"), tax_p - basic_1)
    total_tax_1 = SupremeLegacyEngine.get_tax(taxable_1, inputs.has_spouse, inputs.heirs_info)

    actual_shares, normalized_acquisition_amounts = normalize_actual_acquisition_plan(
        total_taxable_price=tax_p,
        desired_amounts=[int(amount) for amount in input_acquisition_amounts],
        fallback_shares=fallback_shares,
    )
    legal_shares = ([spouse_legal_share] if inputs.has_spouse else []) + heir_legal_shares
    recipient_options = build_recipient_options(inputs.has_spouse, inputs.heirs_info)

    insurance_grosses, insurance_nontaxables, insurance_taxables = allocate_insurance_by_recipient(
        total_insurance=inputs.v_ins,
        entries=inputs.insurance_entries,
        labels=labels,
        statutory_heir_count=st_count,
    )
    total_insurance_taxable = sum(insurance_taxables, Decimal("0"))
    total_gift_addbacks = total_annual_gift_addback + total_seisan_gift_addback
    base_taxable_pool = max(Decimal("0"), tax_p - total_insurance_taxable - total_gift_addbacks)
    desired_base_amounts = [
        max(Decimal("0"), normalized_acquisition_amounts[idx] - insurance_taxables[idx] - annual_gift_addbacks[idx] - seisan_gift_addbacks[idx])
        for idx in range(len(actual_shares))
    ]
    base_taxable_prices = normalize_amounts_to_total(
        total_amount=base_taxable_pool,
        desired_amounts=desired_base_amounts,
        fallback_shares=actual_shares,
    )
    taxable_prices = [
        quantize_yen(base_taxable_prices[idx] + insurance_taxables[idx] + annual_gift_addbacks[idx] + seisan_gift_addbacks[idx])
        for idx in range(len(actual_shares))
    ]
    preliminary_taxes = allocate_preliminary_taxes(total_tax_1, taxable_prices, tax_p)
    two_tenths_targets = [option[3] for option in recipient_options]
    surtax_amounts, adjusted_taxes = apply_two_tenths_surtax(preliminary_taxes, two_tenths_targets)
    final_taxes, spouse_tax_limit, spouse_credit = apply_spouse_tax_credit(
        inputs.has_spouse,
        tax_p,
        spouse_legal_share,
        taxable_prices,
        adjusted_taxes,
    )
    heir_tax_records = build_heir_tax_records(
        labels=labels,
        legal_shares=legal_shares,
        actual_shares=actual_shares,
        input_acquisition_amounts=input_acquisition_amounts if input_acquisition_amounts else normalized_acquisition_amounts,
        normalized_acquisition_amounts=normalized_acquisition_amounts,
        insurance_grosses=insurance_grosses,
        insurance_nontaxables=insurance_nontaxables,
        insurance_taxables=insurance_taxables,
        annual_gift_addbacks=annual_gift_addbacks,
        seisan_gift_addbacks=seisan_gift_addbacks,
        base_taxable_prices=base_taxable_prices,
        taxable_prices=taxable_prices,
        preliminary_taxes=preliminary_taxes,
        surtax_amounts=surtax_amounts,
        final_taxes=final_taxes,
        spouse_credit=spouse_credit,
        two_tenths_targets=two_tenths_targets,
    )

    spouse_actual_share = actual_shares[0] if inputs.has_spouse and actual_shares else Decimal("0")
    spouse_actual_taxable_price = taxable_prices[0] if inputs.has_spouse and taxable_prices else Decimal("0")
    total_final_tax = quantize_yen(sum(final_taxes, Decimal("0")))

    return PrimaryResult(
        st_count=st_count,
        land_eval=land_eval,
        total_red=total_red,
        small_scale_records=small_scale_records,
        ins_ded=ins_ded,
        pure_as=pure_as,
        tax_p=tax_p,
        basic_1=basic_1,
        taxable_1=taxable_1,
        total_tax_1=total_tax_1,
        spouse_legal_share=spouse_legal_share,
        heir_legal_shares=heir_legal_shares,
        spouse_actual_share=spouse_actual_share,
        spouse_actual_taxable_price=spouse_actual_taxable_price,
        spouse_tax_limit=quantize_yen(spouse_tax_limit),
        total_final_tax=total_final_tax,
        total_insurance_gross=sum(insurance_grosses, Decimal("0")),
        total_insurance_nontaxable=sum(insurance_nontaxables, Decimal("0")),
        total_insurance_taxable=sum(insurance_taxables, Decimal("0")),
        total_annual_gift_addback=quantize_yen(total_annual_gift_addback),
        total_seisan_gift_addback=quantize_yen(total_seisan_gift_addback),
        gift_detail_records=gift_detail_records,
        heir_tax_records=heir_tax_records,
    )


def calculate_secondary_inheritance(
    primary_inputs: PrimaryInputs,
    primary_result: PrimaryResult,
    secondary_inputs: SecondaryInputs,
) -> SecondaryResult:
    ratio_s = primary_result.spouse_actual_share if primary_inputs.has_spouse else Decimal("0")
    acq_s_1 = primary_result.spouse_actual_taxable_price
    limit_s = primary_result.spouse_tax_limit
    tax_s_1 = Decimal("0")
    if primary_inputs.has_spouse and primary_result.heir_tax_records:
        tax_s_1 = primary_result.heir_tax_records[0].final_tax
    net_acq_s = acq_s_1 - tax_s_1

    s_own = to_d(secondary_inputs.s_own)
    s_spend_total = to_d(secondary_inputs.annual_spend) * to_d(secondary_inputs.interval_years)
    tax_p_2 = max(Decimal("0"), net_acq_s + s_own - s_spend_total)

    child_only = [h for h in primary_inputs.heirs_info if h["type"] in [HEIR_TYPE_CHILD, HEIR_TYPE_GRANDCHILD]]
    c_count_2 = len(child_only) if child_only else primary_inputs.heir_count
    basic_2 = BASIC_DEDUCTION_BASE + (BASIC_DEDUCTION_PER_HEIR * to_d(c_count_2))
    taxable_2 = max(Decimal("0"), tax_p_2 - basic_2)
    total_tax_2 = SupremeLegacyEngine.get_tax(taxable_2, False, child_only if child_only else primary_inputs.heirs_info)

    return SecondaryResult(
        ratio_s=ratio_s,
        acq_s_1=quantize_yen(acq_s_1),
        limit_s=quantize_yen(limit_s),
        tax_s_1=quantize_yen(tax_s_1),
        net_acq_s=quantize_yen(net_acq_s),
        s_own=s_own,
        s_spend_total=s_spend_total,
        tax_p_2=quantize_yen(tax_p_2),
        c_count_2=c_count_2,
        basic_2=basic_2,
        taxable_2=taxable_2,
        total_tax_2=quantize_yen(total_tax_2),
        child_only=child_only,
    )


def build_small_scale_detail_df(result: PrimaryResult) -> pd.DataFrame:
    rows = []
    for record in result.small_scale_records:
        rows.append({
            "区分": record.category,
            "対象宅地": record.land_name,
            "取得者": record.acquirer_name,
            "判定": record.status,
            "判定理由": record.reason,
            "評価額": int(record.original_value),
            "地積(㎡)": float(record.area_sqm),
            "減額対象面積(㎡)": float(record.eligible_area_sqm),
            "減額率": f"{int(record.reduction_rate * 100)}%" if record.reduction_rate > 0 else "0%",
            "減額額": int(record.reduction_amount),
        })
    return pd.DataFrame(rows)


def build_primary_detail_df(inputs: PrimaryInputs, result: PrimaryResult) -> pd.DataFrame:
    return pd.DataFrame(
        [
            ["1", "不動産評価（小宅反映後）", fmt_int(result.land_eval), f"小宅減額: {fmt_int(result.total_red)}"],
            ["2", "建物・金融・その他合計", fmt_int(to_d(inputs.v_build) + to_d(inputs.v_stock) + to_d(inputs.v_cash) + to_d(inputs.v_others)), ""],
            ["3", "生命保険金(受取人別・控除後)", fmt_int(result.total_insurance_taxable), f"総額: {fmt_int(result.total_insurance_gross)} / 非課税: {fmt_int(result.total_insurance_nontaxable)}"],
            ["4", "債務および葬式費用", f"△{fmt_int(inputs.v_debt + inputs.v_funeral)}", ""],
            ["5", "生前贈与加算（贈与日ベース）", fmt_int(result.total_annual_gift_addback), f"明細件数: {len(inputs.gift_records)}"],
            ["6", "相続時精算課税贈与（110万円控除後）", fmt_int(result.total_seisan_gift_addback), "明細台帳ベース"],
            ["7", "【課税価格合計】", fmt_int(result.tax_p), ""],
            ["8", "基礎控除額", f"△{fmt_int(result.basic_1)}", f"相続人{result.st_count}名"],
            ["9", "課税遺産総額", fmt_int(result.taxable_1), ""],
            ["10", "【相続税の総額】", fmt_int(result.total_tax_1), "法定相続分ベース"],
            ["11", "配偶者税額軽減後の納付税額合計", fmt_int(result.total_final_tax), "概算"],
        ],
        columns=["No", "項目", "金額", "備考"],
    )


def build_primary_heir_tax_df(result: PrimaryResult) -> pd.DataFrame:
    rows = []
    for record in result.heir_tax_records:
        rows.append(
            {
                "相続人": f"{record.name}({record.heir_type})",
                "法定相続分": fmt_pct(record.legal_share),
                "実際取得割合": fmt_pct(record.actual_share),
                "実取得入力額": fmt_int(record.input_acquisition_amount),
                "正規化後取得額": fmt_int(record.normalized_acquisition_amount),
                "保険金総額": fmt_int(record.insurance_gross),
                "保険非課税": fmt_int(record.insurance_nontaxable),
                "保険課税対象": fmt_int(record.insurance_taxable),
                "暦年課税加算": fmt_int(record.annual_gift_addback),
                "精算課税戻し": fmt_int(record.seisan_gift_addback),
                "保険・贈与除外後課税価格": fmt_int(record.base_taxable_price),
                "各人別課税価格": fmt_int(record.taxable_price),
                "按分前税額": fmt_int(record.preliminary_tax),
                "2割加算": fmt_int(record.two_tenths_surtax),
                "配偶者軽減": fmt_int(record.spouse_credit),
                "納付税額": fmt_int(record.final_tax),
                "2割加算対象": "対象" if record.is_two_tenths_target else "",
            }
        )
    return pd.DataFrame(rows)




def build_gift_detail_df(result: PrimaryResult) -> pd.DataFrame:
    rows = []
    for record in result.gift_detail_records:
        rows.append(
            {
                "贈与日": record.gift_date.isoformat(),
                "受贈者": f"{record.recipient_name}({record.recipient_type})",
                "課税方式": record.tax_type,
                "贈与額": int(record.amount),
                "年分": record.calendar_year,
                "加算対象": "対象" if record.is_addback_target else "",
                "相続戻し対象額": int(record.addback_amount),
                "判定理由": record.reason,
            }
        )
    return pd.DataFrame(rows)

def build_secondary_detail_df(result: SecondaryResult) -> pd.DataFrame:
    return pd.DataFrame(
        [
            ["1", "一次からの純承継分", fmt_int(result.net_acq_s), f"配偶者取得{int(result.ratio_s * 100)}%時"],
            ["2", "配偶者固有財産", fmt_int(result.s_own), ""],
            ["3", "生活費・支出等控除", f"△{fmt_int(result.s_spend_total)}", ""],
            ["4", "【二次相続 課税価格】", fmt_int(result.tax_p_2), ""],
            ["5", "二次基礎控除額", f"△{fmt_int(result.basic_2)}", f"相続人{result.c_count_2}名"],
            ["6", "【二次相続税 総額】", fmt_int(result.total_tax_2), "概算"],
        ],
        columns=["No", "項目", "金額", "備考"],
    )


def build_simulation_df(primary_inputs: PrimaryInputs, primary_result: PrimaryResult, secondary_inputs: SecondaryInputs) -> pd.DataFrame:
    sim_results: list[dict[str, Any]] = []
    heirs_for_second = [h for h in primary_inputs.heirs_info if h["type"] in [HEIR_TYPE_CHILD, HEIR_TYPE_GRANDCHILD]]
    heirs_for_second = heirs_for_second if heirs_for_second else primary_inputs.heirs_info
    ratio_candidates = range(0, 101, 10) if primary_inputs.has_spouse else [0]

    for i in ratio_candidates:
        sim_secondary_inputs = SecondaryInputs(
            spouse_acquisition_pct=i,
            s_own=secondary_inputs.s_own,
            annual_spend=secondary_inputs.annual_spend,
            interval_years=secondary_inputs.interval_years,
            use_individual_allocations=secondary_inputs.use_individual_allocations,
            actual_acquisition_inputs=build_simulation_allocation_inputs(
                total_taxable_price=primary_result.tax_p,
                current_inputs=secondary_inputs.actual_acquisition_inputs,
                has_spouse=primary_inputs.has_spouse,
                heirs_info=primary_inputs.heirs_info,
                spouse_acquisition_pct=i,
            ) if secondary_inputs.use_individual_allocations else [],
        )
        sim_primary = calculate_primary_inheritance(primary_inputs, sim_secondary_inputs)
        spouse_tax = sim_primary.heir_tax_records[0].final_tax if primary_inputs.has_spouse and sim_primary.heir_tax_records else Decimal("0")
        spouse_net = sim_primary.spouse_actual_taxable_price - spouse_tax
        tp2 = max(Decimal("0"), spouse_net + to_d(secondary_inputs.s_own) - (to_d(secondary_inputs.annual_spend) * to_d(secondary_inputs.interval_years)))
        c_count_2 = len(heirs_for_second)
        basic_2 = BASIC_DEDUCTION_BASE + (BASIC_DEDUCTION_PER_HEIR * to_d(c_count_2))
        t2 = SupremeLegacyEngine.get_tax(max(Decimal("0"), tp2 - basic_2), False, heirs_for_second)
        sim_results.append(
            {
                "配分(%)": f"{i}%",
                "一次相続税額": int(sim_primary.total_final_tax),
                "二次相続税額": int(quantize_yen(t2)),
                "合計納税額": int(quantize_yen(sim_primary.total_final_tax + t2)),
            }
        )
    return pd.DataFrame(sim_results)


# =========================================================
# 7. Output Logic
# =========================================================
def create_excel_file(df1: pd.DataFrame, df_heirs: pd.DataFrame, df_small: pd.DataFrame, df_gifts: pd.DataFrame, df2: pd.DataFrame, df_sim: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df1.to_excel(writer, sheet_name="一次相続", index=False)
        df_heirs.to_excel(writer, sheet_name="各人別税額", index=False)
        df_small.to_excel(writer, sheet_name="小宅判定", index=False)
        df_gifts.to_excel(writer, sheet_name="贈与台帳", index=False)
        df2.to_excel(writer, sheet_name="二次相続", index=False)
        df_sim.to_excel(writer, sheet_name="シミュレーション", index=False)

    output.seek(0)
    wb = load_workbook(output)

    header_fill = PatternFill(start_color=COLOR_NAVY, end_color=COLOR_NAVY, fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    title_font = Font(size=14, bold=True)
    center_align = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        ws.insert_rows(1)
        ws["A1"] = EXCEL_TITLE
        ws["A1"].font = title_font

        for cell in ws[2]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.border = thin_border
                if isinstance(cell.value, Number):
                    cell.number_format = "#,##0"

        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value is not None:
                    max_length = max(max_length, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = max_length + 2

    final_output = BytesIO()
    wb.save(final_output)
    return final_output.getvalue()


def build_simulation_figure(df_sim: pd.DataFrame) -> go.Figure:
    fig = go.Figure()
    fig.add_trace(go.Bar(x=df_sim["配分(%)"], y=df_sim["一次相続税額"], name="一次相続税", marker_color=f"#{COLOR_NAVY}"))
    fig.add_trace(go.Bar(x=df_sim["配分(%)"], y=df_sim["二次相続税額"], name="二次相続税", marker_color=f"#{COLOR_GOLD}"))
    fig.add_trace(go.Scatter(x=df_sim["配分(%)"], y=df_sim["合計納税額"], name="合計", line=dict(color=f"#{COLOR_RED}", width=4)))
    fig.update_layout(barmode="stack", title="税額最適化シミュレーション", xaxis_title="配偶者配分(%)", yaxis_title="税額(円)")
    return fig


def render_audit_evidence() -> None:
    st.markdown(
        f"""
        <div style="background-color: #f9f9f9; border: 2px solid #{COLOR_GOLD}; padding: 20px; border-radius: 5px;">
            <p style="color: #{COLOR_NAVY}; font-weight: bold; margin-bottom: 10px;">🛡️ 山根会計 監査証跡エビデンス (v31.16)</p>
            <p style="font-size: 0.9em; line-height: 1.6;">担当: 川東 / 本版は概算試算ロジックです。<br>
            各人別課税価格・各人別税額・配偶者税額軽減の土台に加え、生命保険金非課税の受取人別管理、2割加算、小規模宅地等の要件判定付き概算ロジックを反映。さらに一次相続の実取得額を相続人ごとに入力できるUIを追加し、入力差額は内部で比率正規化する設計です。</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


# =========================================================
# 8. UI Rendering
# =========================================================
def render_sidebar() -> None:
    st.sidebar.markdown(f"###  🏢  {APP_TITLE}")
    st.sidebar.info(APP_LOGIN_USER_LABEL)


def render_tab_basic() -> tuple[int, bool, list[dict[str, Any]]]:
    add_print_button("1. 基本構成")
    st.subheader("相続関係の設定")
    c1, c2 = st.columns(2)
    heir_count = c1.number_input("相続人の人数（配偶者除く）", min_value=1, max_value=10, value=2, key="in_child")
    has_spouse = c2.checkbox("配偶者は健在", value=True, key="in_spouse")
    heirs_info: list[dict[str, Any]] = []
    for i in range(heir_count):
        h_type = st.selectbox(f"相続人 {i + 1} の続柄", HEIR_TYPE_OPTIONS, key=f"rel_{i}")
        is_substitute = False
        if h_type == HEIR_TYPE_GRANDCHILD:
            is_substitute = st.checkbox(f"相続人 {i + 1} は代襲相続人", value=False, key=f"substitute_{i}")
        heirs_info.append({"type": h_type, "is_substitute": is_substitute})
    return heir_count, has_spouse, heirs_info


def render_small_scale_input_section(category: str, title: str, has_spouse: bool, heirs_info: list[dict[str, Any]]) -> SmallScaleInput:
    st.write(f"##### {title}：小規模宅地等の特例判定")
    options = build_heir_labels(has_spouse, heirs_info)
    option_labels = [label for label, _ in options] if options else ["未設定"]
    apply_special_rule = st.checkbox(f"{title}で小宅特例を検討する", value=(category == LAND_CATEGORY_HOME), key=f"apply_small_{category}")
    acquirer_name = st.selectbox(f"{title}の取得者", option_labels, key=f"acquirer_{category}")

    if category == LAND_CATEGORY_HOME:
        is_spouse_acquirer = st.checkbox("配偶者が取得", value=(acquirer_name == "配偶者"), key=f"home_spouse_{category}")
        is_cohabiting_heir = st.checkbox("同居親族が取得", value=False, key=f"home_cohab_{category}")
        is_no_house_heir = st.checkbox("家なき子要件に該当", value=False, key=f"home_nohouse_{category}")
        will_hold_until_filing = st.checkbox("申告期限まで保有する", value=True, key=f"home_hold_{category}")
        will_reside_until_filing = st.checkbox("申告期限まで居住継続する", value=True, key=f"home_live_{category}")
        return SmallScaleInput(category=category, acquirer_name=acquirer_name, apply_special_rule=apply_special_rule, is_spouse_acquirer=is_spouse_acquirer, is_cohabiting_heir=is_cohabiting_heir, is_no_house_heir=is_no_house_heir, will_hold_until_filing=will_hold_until_filing, will_reside_until_filing=will_reside_until_filing)

    if category == LAND_CATEGORY_BUSINESS:
        is_business_successor = st.checkbox("取得者は事業承継者", value=False, key=f"biz_successor_{category}")
        will_continue_business = st.checkbox("申告期限まで事業継続", value=False, key=f"biz_continue_{category}")
        will_hold_until_filing = st.checkbox("申告期限まで保有継続", value=False, key=f"biz_hold_{category}")
        return SmallScaleInput(category=category, acquirer_name=acquirer_name, apply_special_rule=apply_special_rule, is_business_successor=is_business_successor, will_continue_business=will_continue_business, will_hold_until_filing=will_hold_until_filing)

    will_continue_rental = st.checkbox("申告期限まで貸付継続", value=False, key=f"rent_continue_{category}")
    will_hold_until_filing = st.checkbox("申告期限まで保有継続", value=False, key=f"rent_hold_{category}")
    return SmallScaleInput(category=category, acquirer_name=acquirer_name, apply_special_rule=apply_special_rule, will_continue_rental=will_continue_rental, will_hold_until_filing=will_hold_until_filing)


def render_tab_primary_inputs(heir_count: int, has_spouse: bool, heirs_info: list[dict[str, Any]]) -> PrimaryInputs:
    add_print_button("2. 一次財産詳細")
    st.subheader("一次相続：財産・贈与入力")
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        st.write("#### 🏗️ 不動産")
        v_home = st.number_input("特定居住用：評価額", min_value=0, value=32781936, key="v_home")
        a_home = st.number_input("特定居住用：面積(㎡)", min_value=0, value=330, key="a_home")
        home_rule = render_small_scale_input_section(LAND_CATEGORY_HOME, "特定居住用", has_spouse, heirs_info)
        st.divider()
        v_biz = st.number_input("特定事業用：評価額", min_value=0, value=0, key="v_biz")
        a_biz = st.number_input("特定事業用：面積(㎡)", min_value=0, value=400, key="a_biz")
        biz_rule = render_small_scale_input_section(LAND_CATEGORY_BUSINESS, "特定事業用", has_spouse, heirs_info)
        st.divider()
        v_rent = st.number_input("貸付事業用：評価額", min_value=0, value=0, key="v_rent")
        a_rent = st.number_input("貸付事業用：面積(㎡)", min_value=0, value=200, key="a_rent")
        rent_rule = render_small_scale_input_section(LAND_CATEGORY_RENTAL, "貸付事業用", has_spouse, heirs_info)
        v_build = st.number_input("建物評価", min_value=0, value=1700044, key="v_build")
    with col_b:
        st.write("#### 💵 金融財産")
        v_stock = st.number_input("有価証券", min_value=0, value=45132788, key="v_stock")
        v_cash = st.number_input("現預金", min_value=0, value=45573502, key="v_cash")
        v_ins = st.number_input("生命保険金（総額）", min_value=0, value=3651514, key="v_ins")
        recipient_options = build_recipient_options(has_spouse, heirs_info)
        recipient_labels = [f"{label}（{heir_type}）" for label, heir_type, _, _ in recipient_options]
        recipient_map = {f"{label}（{heir_type}）": (label, heir_type, is_statutory_heir, is_two_tenths_target) for label, heir_type, is_statutory_heir, is_two_tenths_target in recipient_options}
        insurance_entry_count = st.number_input("生命保険金の受取人明細数", min_value=0, max_value=MAX_INSURANCE_RECIPIENT_ROWS, value=min(1, len(recipient_labels)), key="insurance_entry_count")
        insurance_entries: list[InsuranceRecipientInput] = []
        for idx in range(int(insurance_entry_count)):
            recipient_key = st.selectbox(f"保険金受取人 {idx + 1}", recipient_labels, key=f"insurance_recipient_{idx}")
            amount = st.number_input(f"保険金受取額 {idx + 1}", min_value=0, value=v_ins if idx == 0 else 0, key=f"insurance_amount_{idx}")
            label, heir_type, is_statutory_heir, is_two_tenths_target = recipient_map[recipient_key]
            insurance_entries.append(InsuranceRecipientInput(recipient_name=label, recipient_type=heir_type, amount=amount, is_statutory_heir=is_statutory_heir, is_two_tenths_target=is_two_tenths_target))
        if insurance_entries:
            entered_total = sum(entry.amount for entry in insurance_entries)
            st.caption(f"※受取人明細合計: {entered_total:,}円 / 総額入力: {v_ins:,}円。差異がある場合は内部で総額に合わせて按分補正します。")
        v_others = st.number_input("その他", min_value=0, value=1662687, key="v_others")
        st.write("#### 📉 債務・葬式")
        v_debt = st.number_input("債務", min_value=0, value=322179, key="v_debt")
        v_funeral = st.number_input("葬式費用", min_value=0, value=41401, key="v_funeral")
    with col_c:
        st.write("#### 🎁 贈与台帳（明細管理）")
        date_of_death = st.date_input("相続開始日", value=date.today(), key="date_of_death")
        gift_recipient_options = build_gift_recipient_options(has_spouse, heirs_info)
        gift_recipient_labels = [f"{label}（{heir_type}）" for label, heir_type in gift_recipient_options]
        gift_recipient_map = {f"{label}（{heir_type}）": (label, heir_type) for label, heir_type in gift_recipient_options}
        gift_record_count = st.number_input("贈与明細件数", min_value=0, max_value=MAX_GIFT_RECORD_ROWS, value=0, key="gift_record_count")
        gift_records: list[GiftRecord] = []
        for idx in range(int(gift_record_count)):
            st.markdown(f"**贈与明細 {idx + 1}**")
            gift_date_value = st.date_input(f"贈与日 {idx + 1}", value=date.today(), key=f"gift_date_{idx}")
            recipient_key = st.selectbox(f"受贈者 {idx + 1}", gift_recipient_labels, key=f"gift_recipient_{idx}")
            gift_amount = st.number_input(f"贈与額 {idx + 1}", min_value=0, value=0, key=f"gift_amount_{idx}")
            gift_tax_type = st.selectbox(f"課税方式 {idx + 1}", GIFT_TYPE_OPTIONS, key=f"gift_tax_type_{idx}")
            recipient_name, recipient_type = gift_recipient_map[recipient_key]
            gift_records.append(
                GiftRecord(
                    gift_date=gift_date_value,
                    recipient_name=recipient_name,
                    recipient_type=recipient_type,
                    amount=gift_amount,
                    tax_type=gift_tax_type,
                )
            )
        if gift_records:
            annual_total = sum(record.amount for record in gift_records if record.tax_type == GIFT_TYPE_ANNUAL)
            seisan_total = sum(record.amount for record in gift_records if record.tax_type == GIFT_TYPE_SEISAN)
            st.caption(f"暦年課税入力合計: {annual_total:,}円 / 相続時精算課税入力合計: {seisan_total:,}円")
        else:
            st.caption("贈与明細がない場合は0件のままで構いません。")

    return PrimaryInputs(
        heir_count=heir_count,
        has_spouse=has_spouse,
        heirs_info=heirs_info,
        date_of_death=date_of_death,
        v_home=v_home,
        a_home=a_home,
        v_biz=v_biz,
        a_biz=a_biz,
        v_rent=v_rent,
        a_rent=a_rent,
        small_scale_inputs={
            LAND_CATEGORY_HOME: home_rule,
            LAND_CATEGORY_BUSINESS: biz_rule,
            LAND_CATEGORY_RENTAL: rent_rule,
        },
        v_build=v_build,
        v_stock=v_stock,
        v_cash=v_cash,
        v_ins=v_ins,
        insurance_entries=insurance_entries,
        v_others=v_others,
        v_debt=v_debt,
        v_funeral=v_funeral,
        gift_records=gift_records,
    )


def render_tab_primary_detail(df1: pd.DataFrame, df_heirs: pd.DataFrame, df_small: pd.DataFrame, df_gifts: pd.DataFrame) -> None:
    add_print_button("3. 一次相続明細")
    st.subheader("一次相続：計算明細")
    st.table(df1)
    st.divider()
    st.subheader("小規模宅地等の特例 判定結果")
    st.dataframe(df_small, use_container_width=True)
    st.divider()
    st.subheader("贈与加算・相続時精算課税 明細")
    if df_gifts.empty:
        st.caption("贈与明細はありません。")
    else:
        st.dataframe(df_gifts, use_container_width=True)
    st.divider()
    st.subheader("各人別課税価格・各人別税額（概算）")
    st.dataframe(df_heirs, use_container_width=True)


def render_tab_secondary_detail(df2: pd.DataFrame) -> None:
    add_print_button("4. 二次相続明細")
    st.subheader("二次相続：計算明細予測")
    st.table(df2)


def estimate_total_taxable_price_reference(primary_inputs: PrimaryInputs) -> Decimal:
    total_red, land_eval, _ = calculate_small_scale_reduction(primary_inputs)
    st_count = primary_inputs.heir_count + (1 if primary_inputs.has_spouse else 0)
    ins_ded = calculate_life_insurance_deduction(primary_inputs.v_ins, st_count)
    pure_as = land_eval + to_d(primary_inputs.v_build) + to_d(primary_inputs.v_stock) + to_d(primary_inputs.v_cash) + to_d(primary_inputs.v_ins) + to_d(primary_inputs.v_others)
    labels = build_heir_labels(primary_inputs.has_spouse, primary_inputs.heirs_info)
    annual_addbacks, seisan_addbacks, _ = calculate_gift_addbacks(primary_inputs.gift_records, primary_inputs.date_of_death, labels)
    tax_p = pure_as - ins_ded - to_d(primary_inputs.v_debt) - to_d(primary_inputs.v_funeral) + sum(annual_addbacks, Decimal("0")) + sum(seisan_addbacks, Decimal("0"))
    return quantize_yen(max(Decimal("0"), tax_p))


def build_default_acquisition_input_amounts(total_taxable_price: Decimal, has_spouse: bool, heirs_info: list[dict[str, Any]]) -> list[int]:
    fallback_shares = allocate_actual_shares(has_spouse, heirs_info, 50 if has_spouse else 0)
    defaults = normalize_amounts_to_total(total_taxable_price, [], fallback_shares)
    return [int(amount) for amount in defaults]


def build_simulation_allocation_inputs(
    total_taxable_price: Decimal,
    current_inputs: list[int],
    has_spouse: bool,
    heirs_info: list[dict[str, Any]],
    spouse_acquisition_pct: int,
) -> list[int]:
    fallback_shares = allocate_actual_shares(has_spouse, heirs_info, spouse_acquisition_pct)
    if not has_spouse:
        normalized = normalize_amounts_to_total(total_taxable_price, [to_d(max(0, amount)) for amount in current_inputs], fallback_shares)
        return [int(amount) for amount in normalized]

    non_spouse_count = len(heirs_info)
    spouse_amount = quantize_yen(total_taxable_price * to_d(spouse_acquisition_pct) / PERCENT_DENOMINATOR)
    remaining_amount = max(Decimal("0"), total_taxable_price - spouse_amount)

    desired_non_spouse = [to_d(max(0, amount)) for amount in current_inputs[1 : 1 + non_spouse_count]]
    fallback_non_spouse_shares = fallback_shares[1:] if len(fallback_shares) > 1 else [Decimal("0")] * non_spouse_count
    normalized_non_spouse = normalize_amounts_to_total(remaining_amount, desired_non_spouse, fallback_non_spouse_shares)

    combined = [spouse_amount] + normalized_non_spouse
    return [int(amount) for amount in combined]




def render_tab_secondary_parameters(has_spouse: bool, heirs_info: list[dict[str, Any]], estimated_tax_p: Decimal) -> SecondaryInputs:
    add_print_button("5. 二次推移予測")
    st.subheader("二次推移パラメータ設定")
    cp1, cp2 = st.columns(2)
    spouse_acquisition_pct = 0
    if has_spouse:
        spouse_acquisition_pct = cp1.slider("一次相続における配偶者の取得割合(%)", min_value=0, max_value=100, value=50, key="in_spouse_ratio")
    else:
        cp1.caption("配偶者がいないため、配偶者取得割合は0%で固定です。")

    use_individual_allocations = st.checkbox("一次相続の実取得額を相続人ごとに入力する（推奨）", value=True, key="use_individual_allocations")
    labels = build_heir_labels(has_spouse, heirs_info)
    default_amounts = build_default_acquisition_input_amounts(estimated_tax_p, has_spouse, heirs_info)
    actual_acquisition_inputs: list[int] = []

    if use_individual_allocations:
        st.caption(f"参考：一次相続の課税価格合計 {fmt_int(estimated_tax_p)}円 を基準に、各人の実取得額（概算）を入力してください。合計が一致しない場合は内部で比率按分します。")
        for idx, (label, heir_type) in enumerate(labels):
            key = f"actual_acq_{idx}"
            if key not in st.session_state:
                st.session_state[key] = default_amounts[idx] if idx < len(default_amounts) else 0
            amount = st.number_input(
                f"{label}（{heir_type}）の実取得額",
                min_value=0,
                value=int(st.session_state[key]),
                key=key,
            )
            actual_acquisition_inputs.append(int(amount))
        entered_total = sum(actual_acquisition_inputs)
        discrepancy = int(estimated_tax_p) - entered_total
        st.caption(f"入力合計: {entered_total:,}円 / 参考課税価格合計: {int(estimated_tax_p):,}円 / 差額: {discrepancy:,}円")
    else:
        actual_acquisition_inputs = default_amounts

    s_own = cp1.number_input("配偶者の固有財産", min_value=0, value=50000000, key="in_s_own")
    interval_years = cp1.slider("二次までの想定期間(年)", min_value=0, max_value=20, value=10, key="in_interval")
    annual_spend = cp2.number_input("年間生活費・支出(減価)", min_value=0, value=5000000, key="in_s_spend")
    return SecondaryInputs(
        spouse_acquisition_pct=spouse_acquisition_pct,
        s_own=s_own,
        annual_spend=annual_spend,
        interval_years=interval_years,
        use_individual_allocations=use_individual_allocations,
        actual_acquisition_inputs=actual_acquisition_inputs,
    )


def render_tab_analysis(df_sim: pd.DataFrame, iryu_df: pd.DataFrame, df1: pd.DataFrame, df_heirs: pd.DataFrame, df_small: pd.DataFrame, df_gifts: pd.DataFrame, df2: pd.DataFrame) -> None:
    add_print_button("6. 精密分析結果")
    st.subheader("配偶者取得割合別の税額推移分析")
    st.plotly_chart(build_simulation_figure(df_sim), use_container_width=True)
    st.dataframe(
        df_sim.style.format({"一次相続税額": "{:,}", "二次相続税額": "{:,}", "合計納税額": "{:,}"}),
        use_container_width=True,
    )

    st.divider()
    st.subheader("⚠️ 遺留分侵害額の参考表示")
    st.table(iryu_df)

    st.divider()
    render_audit_evidence()

    st.divider()
    st.subheader("📥 Excel出力")
    try:
        excel_data = create_excel_file(df1, df_heirs, df_small, df_gifts, df2, df_sim)
        st.download_button(
            label="📊 Excelファイルをダウンロード",
            data=excel_data,
            file_name=EXCEL_FILE_NAME,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as exc:
        st.error(f"Excel出力エラー: {exc}")


# =========================================================
# 9. main()
# =========================================================
def main() -> None:
    inject_print_css()
    if not authenticate_user():
        return

    render_sidebar()
    tabs = st.tabs(TAB_LABELS)

    with tabs[0]:
        heir_count, has_spouse, heirs_info = render_tab_basic()

    with tabs[1]:
        primary_inputs = render_tab_primary_inputs(heir_count, has_spouse, heirs_info)

    estimated_tax_p = estimate_total_taxable_price_reference(primary_inputs)

    with tabs[4]:
        secondary_inputs = render_tab_secondary_parameters(has_spouse, heirs_info, estimated_tax_p)

    primary_result = calculate_primary_inheritance(primary_inputs, secondary_inputs)
    secondary_result = calculate_secondary_inheritance(primary_inputs, primary_result, secondary_inputs)
    df1 = build_primary_detail_df(primary_inputs, primary_result)
    df_heirs = build_primary_heir_tax_df(primary_result)
    df_small = build_small_scale_detail_df(primary_result)
    df_gifts = build_gift_detail_df(primary_result)
    df2 = build_secondary_detail_df(secondary_result)
    df_sim = build_simulation_df(primary_inputs, primary_result, secondary_inputs)
    iryu_df = build_iryubun_reference(primary_inputs, primary_result)

    with tabs[2]:
        render_tab_primary_detail(df1, df_heirs, df_small, df_gifts)

    with tabs[3]:
        render_tab_secondary_detail(df2)

    with tabs[5]:
        render_tab_analysis(df_sim, iryu_df, df1, df_heirs, df_small, df_gifts, df2)


if __name__ == "__main__":
    main()
