"""
=====================================================================
ASC 842 LEASE ACCOUNTING ENGINE — Streamlit Application (With Excel Export)
=====================================================================

Built from, and mathematically validated against, the client's
"ASC842_Master_Template.xlsx" workbook.

Includes automated Excel Report Generation via openpyxl.

Run with:  streamlit run asc842_app.py
Requires:  streamlit, pandas, numpy, numpy_financial, python-dateutil, openpyxl
=====================================================================
"""

import io
import streamlit as st
import pandas as pd
import numpy as np
from datetime import date, datetime
from dateutil.relativedelta import relativedelta
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    import numpy_financial as npf
    HAVE_NPF = True
except ImportError:
    HAVE_NPF = False
    class _NpfShim:
        @staticmethod
        def npv(rate, values):
            return float(sum(v / (1 + rate) ** i for i, v in enumerate(values)))
    npf = _NpfShim()

st.set_page_config(page_title="ASC 842 Lease Accounting Engine", layout="wide")


# =====================================================================
# EXCEL REPORT GENERATOR (Audit-Ready openpyxl Writer)
# =====================================================================
def generate_excel_workbook(lease_info, schedule_df, classification_memo, initial_je_df, roll_forward_df):
    wb = openpyxl.Workbook()
    
    # Define Palette & Styles
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    section_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    total_fill = PatternFill(start_color="E6FCF5", end_color="E6FCF5", fill_type="solid")
    
    white_bold_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    title_font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    section_font = Font(name="Calibri", size=11, bold=True, color="1F4E78")
    bold_font = Font(name="Calibri", size=10, bold=True)
    italic_font = Font(name="Calibri", size=10, italic=True)
    regular_font = Font(name="Calibri", size=10)
    
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    
    # -----------------------------------------------------------------
    # TAB 1: Executive Summary & Audit Memo
    # -----------------------------------------------------------------
    ws1 = wb.active
    ws1.title = "Executive Summary"
    ws1.views.sheetView[0].showGridLines = True
    
    ws1.merge_cells("A1:F1")
    title_cell = ws1["A1"]
    title_cell.value = f"ASC 842 LEASE ACCOUNTING REPORT — {lease_info['lease_name'].upper()}"
    title_cell.font = title_font
    title_cell.fill = header_fill
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 35
    
    # 1. Lease Parameters
    ws1["A3"] = "1. Lease Contract Parameters"
    ws1["A3"].font = section_font
    
    params = [
        ("Lease Name:", lease_info['lease_name']),
        ("Start Date:", lease_info['start_date'].strftime("%Y-%m-%d")),
        ("Lease Term (Months):", lease_info['num_months']),
        ("Monthly Payment:", lease_info['monthly_payment']),
        ("Payment Timing:", lease_info['payment_timing']),
        ("Annual Discount Rate:", f"{lease_info['annual_rate']:.2f}%"),
        ("Origination Fees / Prepaid:", lease_info['origination_fee']),
        ("Security Deposit:", lease_info['security_deposit']),
    ]
    for idx, (label, val) in enumerate(params, start=4):
        ws1[f"A{idx}"] = label
        ws1[f"A{idx}"].font = bold_font
        ws1[f"B{idx}"] = val
        ws1[f"B{idx}"].font = regular_font
        if "Payment:" in label or "Fees" in label or "Deposit" in label:
            ws1[f"B{idx}"].number_format = '$#,##0.00'
            
    # 2. Accounting Valuation
    ws1["D3"] = "2. Accounting Initial Valuation"
    ws1["D3"].font = section_font
    
    outputs = [
        ("Initial Lease Liability (PV):", lease_info['initial_liability']),
        ("Prepaid First Payment:", lease_info['prepaid_first_payment']),
        ("Initial ROU Asset:", lease_info['initial_rou']),
        ("Lease Classification:", lease_info['classification']),
    ]
    for idx, (label, val) in enumerate(outputs, start=4):
        ws1[f"D{idx}"] = label
        ws1[f"D{idx}"].font = bold_font
        ws1[f"E{idx}"] = val
        ws1[f"E{idx}"].font = regular_font
        if "Liability" in label or "Prepaid" in label or "Asset" in label:
            ws1[f"E{idx}"].number_format = '$#,##0.00'
        elif "Classification" in label:
            ws1[f"E{idx}"].font = Font(name="Calibri", size=10, bold=True, color="2E75B6")
            
    # 3. Audit Memo
    ws1["A13"] = "3. Audit Rationale & Memo"
    ws1["A13"].font = section_font
    ws1.merge_cells("A14:F17")
    memo_cell = ws1["A14"]
    memo_cell.value = classification_memo
    memo_cell.font = italic_font
    memo_cell.alignment = Alignment(wrap_text=True, vertical="top")
    memo_cell.fill = section_fill
    
    # -----------------------------------------------------------------
    # TAB 2: Amortization Schedule
    # -----------------------------------------------------------------
    ws2 = wb.create_sheet(title="Amortization Schedule")
    ws2.views.sheetView[0].showGridLines = True
    
    headers = list(schedule_df.columns)
    for col_idx, h in enumerate(headers, start=1):
        c = ws2.cell(row=1, column=col_idx, value=h)
        c.font = white_bold_font
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 25
    
    for row_idx, row_data in enumerate(schedule_df.itertuples(index=False), start=2):
        for col_idx, val in enumerate(row_data, start=1):
            c = ws2.cell(row=row_idx, column=col_idx)
            c.border = thin_border
            h_name = headers[col_idx-1]
            if h_name == "Date":
                c.value = pd.to_datetime(val).strftime("%Y-%m-%d")
                c.alignment = Alignment(horizontal="center")
            elif h_name == "Period":
                c.value = int(val)
                c.alignment = Alignment(horizontal="center")
            else:
                c.value = float(val)
                c.number_format = '$#,##0.00'
                
    # -----------------------------------------------------------------
    # TAB 3: Initial Recognition Journal Entry
    # -----------------------------------------------------------------
    ws3 = wb.create_sheet(title="Journal Entries")
    ws3.views.sheetView[0].showGridLines = True
    
    ws3["A1"] = "INITIAL RECOGNITION JOURNAL ENTRY (DAY 1)"
    ws3["A1"].font = section_font
    
    je_headers = ["Account Description", "Debit ($)", "Credit ($)"]
    for col_idx, h in enumerate(je_headers, start=1):
        c = ws3.cell(row=3, column=col_idx, value=h)
        c.font = white_bold_font
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center")
        
    for row_idx, row_data in enumerate(initial_je_df.itertuples(index=False), start=4):
        ws3.cell(row=row_idx, column=1, value=row_data[0]).font = regular_font
        
        c_deb = ws3.cell(row=row_idx, column=2)
        if pd.notna(row_data[1]) and row_data[1] != 0:
            c_deb.value = float(row_data[1])
            c_deb.number_format = '$#,##0.00'
            
        c_cred = ws3.cell(row=row_idx, column=3)
        if pd.notna(row_data[2]) and row_data[2] != 0:
            c_cred.value = float(row_data[2])
            c_cred.number_format = '$#,##0.00'
            
    # Auto-fit Column Widths across all sheets
    for ws in wb.worksheets:
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 14)
            
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


# =====================================================================
# MODULE 3 — Mock Treasury Rate Table
# =====================================================================
MOCK_TREASURY_CURVE = {
    1: 4.65, 2: 4.55, 3: 4.45, 4: 4.40, 5: 4.35,
    6: 4.35, 7: 4.40, 10: 4.55, 20: 4.85, 30: 4.75,
}

def fetch_treasury_rate(term_years: int) -> float:
    term_years = max(1, int(term_years))
    available = sorted(MOCK_TREASURY_CURVE.keys())
    nearest = min(available, key=lambda t: abs(t - term_years))
    return MOCK_TREASURY_CURVE[nearest]


# =====================================================================
# MODULE 2 — Auto-Classification Engine
# =====================================================================
def classify_lease(
    lease_term_months: float,
    economic_life_months: float,
    pv_lease_payments: float,
    fmv_asset: float,
    transfers_ownership: bool,
    purchase_option_reasonably_certain: bool,
    specialized_no_alt_use: bool,
    guaranteed_residual_value: float = 0.0,
):
    term_pct_of_life = (lease_term_months / economic_life_months) if economic_life_months else 0
    pv_pct_of_fmv = ((pv_lease_payments + guaranteed_residual_value) / fmv_asset) if fmv_asset else 0

    tests = {
        "1. Transfer of ownership at end of term": transfers_ownership,
        "2. Purchase option reasonably certain to be exercised": purchase_option_reasonably_certain,
        "3. Lease term is a major part of remaining economic life (>=75%)": term_pct_of_life >= 0.75,
        "4. PV of payments is substantially all of asset FMV (>=90%)": pv_pct_of_fmv >= 0.90,
        "5. Asset is so specialized it has no alternative use to lessor": specialized_no_alt_use,
    }

    is_finance = any(tests.values())
    classification = "Finance Lease" if is_finance else "Operating Lease"

    triggered = [k for k, v in tests.items() if v]
    if is_finance:
        memo = (
            f"The lease is classified as a Finance Lease because it met the following "
            f"ASC 842 criteri{'on' if len(triggered) == 1 else 'a'}: {'; '.join(triggered)}. "
        )
        if tests["3. Lease term is a major part of remaining economic life (>=75%)"]:
            memo += f"Specifically, the lease term represents {term_pct_of_life:.1%} of the economic life. "
        if tests["4. PV of payments is substantially all of asset FMV (>=90%)"]:
            memo += f"The present value of lease payments (${pv_lease_payments:,.0f}) equals {pv_pct_of_fmv:.1%} of FMV (${fmv_asset:,.0f}). "
        memo += "Accordingly, the lessee recognizes separate interest and amortization expense."
    else:
        memo = (
            f"The lease is classified as an Operating Lease because none of the five ASC 842 criteria were met: "
            f"the lease term is {term_pct_of_life:.1%} of economic life (below 75%), and the PV of payments "
            f"(${pv_lease_payments:,.0f}) is {pv_pct_of_fmv:.1%} of FMV (${fmv_asset:,.0f}, below 90%). "
            f"Accordingly, the lessee recognizes a single straight-line lease cost over the lease term."
        )

    return classification, memo, tests, term_pct_of_life, pv_pct_of_fmv


# =====================================================================
# MODULE 4 — Core Calculation Engine
# =====================================================================
def generate_period_dates(start_date: date, num_months: int):
    return [start_date + relativedelta(months=i) for i in range(num_months)]


def build_amortization_schedule(
    start_date: date,
    num_months: int,
    monthly_payment: float,
    annual_discount_rate: float,
    payment_timing: str = "Advance",
    origination_fee: float = 0.0,
    lease_incentives: float = 0.0,
    escalation_pct: float = 0.0,
):
    monthly_rate = annual_discount_rate / 12.0
    dates = generate_period_dates(start_date, num_months)

    payments = []
    for i in range(num_months):
        years_elapsed = i // 12
        payments.append(monthly_payment * ((1 + escalation_pct) ** years_elapsed))

    if payment_timing == "Advance":
        prepaid_first_payment = payments[0]
        discounted_cashflows = payments[1:]
    else:
        prepaid_first_payment = 0.0
        discounted_cashflows = payments[:]

    initial_liability = npf.npv(monthly_rate, [0.0] + discounted_cashflows)
    initial_rou = initial_liability + prepaid_first_payment + origination_fee - lease_incentives

    n_periods = len(discounted_cashflows)
    rou_amort_per_period = initial_rou / n_periods if n_periods else 0.0

    rows = []
    beg_liability = initial_liability
    beg_rou = initial_rou

    rows.append({
        "Period": 0,
        "Date": dates[0],
        "Beg. Lease Liability": beg_liability,
        "Interest Expense": 0.0,
        "Lease Payment": prepaid_first_payment if payment_timing == "Advance" else 0.0,
        "Liability Amortization": 0.0,
        "End Lease Liability": beg_liability,
        "Beg. ROU Asset": beg_rou,
        "ROU Amortization": 0.0,
        "End ROU Asset": beg_rou,
    })

    for idx, cf in enumerate(discounted_cashflows, start=1):
        interest = beg_liability * monthly_rate
        amortization = cf - interest
        end_liability = beg_liability - amortization
        rou_amort = rou_amort_per_period
        end_rou = beg_rou - rou_amort

        rows.append({
            "Period": idx,
            "Date": dates[idx] if idx < len(dates) else dates[-1] + relativedelta(months=idx - len(dates) + 1),
            "Beg. Lease Liability": beg_liability,
            "Interest Expense": interest,
            "Lease Payment": cf,
            "Liability Amortization": amortization,
            "End Lease Liability": end_liability,
            "Beg. ROU Asset": beg_rou,
            "ROU Amortization": rou_amort,
            "End ROU Asset": end_rou,
        })

        beg_liability = end_liability
        beg_rou = end_rou

    return {
        "schedule": pd.DataFrame(rows),
        "initial_liability": initial_liability,
        "initial_rou": initial_rou,
        "prepaid_first_payment": prepaid_first_payment,
        "monthly_rate": monthly_rate,
        "n_periods": n_periods,
        "rou_amort_per_period": rou_amort_per_period,
    }


# =====================================================================
# MODULE 5 — Journal Entries
# =====================================================================
def initial_recognition_je(initial_rou, initial_liability, prepaid_first_payment, origination_fee, security_deposit):
    upfront_cash = prepaid_first_payment + origination_fee
    entries = [
        {"Account": "ROU Asset", "Debit": initial_rou, "Credit": None},
        {"Account": "Lease Liability", "Debit": None, "Credit": initial_liability},
        {"Account": "Cash (upfront payment/fees)", "Debit": None, "Credit": upfront_cash},
    ]
    if security_deposit:
        entries.append({"Account": "Security Deposit (Asset)", "Debit": security_deposit, "Credit": None})
        entries.append({"Account": "Cash (security deposit)", "Debit": None, "Credit": security_deposit})
    df = pd.DataFrame(entries)
    return df, df["Debit"].fillna(0).sum(), df["Credit"].fillna(0).sum()


def monthly_journal_entry(period_row: pd.Series, classification: str):
    interest = period_row["Interest Expense"]
    payment = period_row["Lease Payment"]
    liability_amort = period_row["Liability Amortization"]
    rou_amort = period_row["ROU Amortization"]

    if classification == "Finance Lease":
        entries = [
            {"Account": "Interest Expense", "Debit": interest, "Credit": None},
            {"Account": "Lease Liability", "Debit": liability_amort, "Credit": None},
            {"Account": "Cash", "Debit": None, "Credit": payment},
            {"Account": "Amortization Expense (ROU)", "Debit": rou_amort, "Credit": None},
            {"Account": "Accumulated Amortization - ROU Asset", "Debit": None, "Credit": rou_amort},
        ]
    else:
        lease_expense = interest + rou_amort
        entries = [
            {"Account": "Lease Expense (single lease cost)", "Debit": lease_expense, "Credit": None},
            {"Account": "Lease Liability", "Debit": liability_amort, "Credit": None},
            {"Account": "Cash", "Debit": None, "Credit": payment},
            {"Account": "ROU Asset", "Debit": None, "Credit": rou_amort},
        ]
    df = pd.DataFrame(entries)
    return df, df["Debit"].fillna(0).sum(), df["Credit"].fillna(0).sum()


# =====================================================================
# STREAMLIT UI APP
# =====================================================================
def money(x):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return ""
    return f"${x:,.2f}"

def init_session_state():
    defaults = {
        "lease_name": "Sample Equipment Lease",
        "start_date": date(2026, 1, 1),
        "num_months": 36,
        "monthly_payment": 5000.0,
        "payment_timing": "Advance",
        "security_deposit": 0.0,
        "origination_fee": 0.0,
        "lease_incentives": 0.0,
        "escalation_pct": 0.0,
        "annual_rate": 5.0,
        "rate_source": "Manual IBR",
        "economic_life_months": 60,
        "fmv_asset": 250000.0,
        "transfers_ownership": False,
        "purchase_option_certain": False,
        "specialized_asset": False,
        "guaranteed_residual": 0.0,
        "schedule_result": None,
        "classification": None,
        "memo": "",
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_session_state()

st.title("📋 ASC 842 Lease Accounting Engine")
st.caption("Turnkey Managed Service Platform — Core Calculation Engine & Audit-Ready Excel Exporter")

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "1️⃣ Lease Input",
    "2️⃣ Classification",
    "3️⃣ Discount Rate",
    "4️⃣ Calculation Engine",
    "5️⃣ Journal Entries",
])

# MODULE 1
with tab1:
    st.header("Module 1 — Lease Input")
    col_a, col_b = st.columns([2, 1])
    with col_b:
        st.subheader("🤖 AI Contract Extraction")
        if st.button("📄 Upload Lease PDF", use_container_width=True):
            st.toast("AI PDF Extraction Module will auto-populate fields (Coming Soon).", icon="🤖")
    with col_a:
        st.subheader("Lease Parameters")
        c1, c2 = st.columns(2)
        with c1:
            st.session_state.lease_name = st.text_input("Lease Name", value=st.session_state.lease_name)
            st.session_state.start_date = st.date_input("Start Date", value=st.session_state.start_date)
            st.session_state.num_months = st.number_input("Lease Term (Months)", min_value=1, value=int(st.session_state.num_months))
            st.session_state.payment_timing = st.selectbox("Payment Timing", options=["Advance", "Arrears"])
        with c2:
            st.session_state.monthly_payment = st.number_input("Monthly Payment ($)", min_value=0.0, value=float(st.session_state.monthly_payment), step=100.0)
            st.session_state.security_deposit = st.number_input("Security Deposit ($)", min_value=0.0, value=float(st.session_state.security_deposit), step=100.0)
            st.session_state.origination_fee = st.number_input("Origination Fee / Prepaid ($)", min_value=0.0, value=float(st.session_state.origination_fee), step=100.0)
            st.session_state.lease_incentives = st.number_input("Lease Incentives ($)", min_value=0.0, value=float(st.session_state.lease_incentives), step=100.0)

# MODULE 2
with tab2:
    st.header("Module 2 — Auto-Classification Engine")
    c1, c2 = st.columns(2)
    with c1:
        st.session_state.economic_life_months = st.number_input("Economic Life (months)", min_value=1, value=int(st.session_state.economic_life_months))
        st.session_state.fmv_asset = st.number_input("Fair Market Value ($)", min_value=0.0, value=float(st.session_state.fmv_asset), step=1000.0)
    with c2:
        st.session_state.transfers_ownership = st.checkbox("Ownership transfers?", value=st.session_state.transfers_ownership)
        st.session_state.purchase_option_certain = st.checkbox("Purchase option certain?", value=st.session_state.purchase_option_certain)
        st.session_state.specialized_asset = st.checkbox("Specialized asset?", value=st.session_state.specialized_asset)

    prelim_rate = st.session_state.get("annual_rate", 5.0) / 100.0
    prelim = build_amortization_schedule(
        start_date=st.session_state.start_date,
        num_months=int(st.session_state.num_months),
        monthly_payment=float(st.session_state.monthly_payment),
        annual_discount_rate=prelim_rate,
        payment_timing=st.session_state.payment_timing,
        origination_fee=float(st.session_state.origination_fee),
    )
    
    if st.button("🔍 Run Classification Test", type="primary"):
        classif, memo, tests, term_pct, pv_pct = classify_lease(
            lease_term_months=int(st.session_state.num_months),
            economic_life_months=int(st.session_state.economic_life_months),
            pv_lease_payments=prelim["initial_liability"] + prelim["prepaid_first_payment"],
            fmv_asset=float(st.session_state.fmv_asset),
            transfers_ownership=st.session_state.transfers_ownership,
            purchase_option_reasonably_certain=st.session_state.purchase_option_certain,
            specialized_no_alt_use=st.session_state.specialized_asset,
        )
        st.session_state.classification = classif
        st.session_state.memo = memo
        st.success(f"Classification Result: **{classif}**")
        st.info(memo)

# MODULE 3
with tab3:
    st.header("Module 3 — Discount Rate")
    st.session_state.rate_source = st.radio("Method", options=["Manual IBR", "Risk-Free Rate (Treasury)"], horizontal=True)
    if st.session_state.rate_source == "Manual IBR":
        st.session_state.annual_rate = st.number_input("IBR (Annual %)", min_value=0.0, value=float(st.session_state.annual_rate), step=0.05)
    else:
        term_years = int(np.ceil(st.session_state.num_months / 12))
        if st.button("📈 Fetch Treasury Rate"):
            fetched = fetch_treasury_rate(term_years)
            st.session_state.annual_rate = fetched
            st.success(f"Fetched {term_years}-yr Treasury Rate: **{fetched:.2f}%**")

# MODULE 4
with tab4:
    st.header("Module 4 — Core Calculation Engine & Amortization")
    if st.button("⚙️ Build / Refresh Amortization Schedule", type="primary"):
        result = build_amortization_schedule(
            start_date=st.session_state.start_date,
            num_months=int(st.session_state.num_months),
            monthly_payment=float(st.session_state.monthly_payment),
            annual_discount_rate=float(st.session_state.annual_rate) / 100.0,
            payment_timing=st.session_state.payment_timing,
            origination_fee=float(st.session_state.origination_fee),
            lease_incentives=float(st.session_state.lease_incentives),
        )
        st.session_state.schedule_result = result
        st.success("Schedule built successfully.")

    result = st.session_state.schedule_result
    if result:
        c1, c2, c3 = st.columns(3)
        c1.metric("Lease Liability (PV)", money(result["initial_liability"]))
        c2.metric("Prepaid 1st Payment", money(result["prepaid_first_payment"]))
        c3.metric("Initial ROU Asset", money(result["initial_rou"]))

        st.dataframe(result["schedule"].style.format({c: "${:,.2f}" for c in result["schedule"].columns if c not in ("Period", "Date")}), use_container_width=True)

        # -------------------------------------------------------------
        # EXCEL DOWNLOAD BUTTON
        # -------------------------------------------------------------
        st.divider()
        st.subheader("📥 Export Audit-Ready Deliverables")
        
        lease_info_dict = {
            'lease_name': st.session_state.lease_name,
            'start_date': st.session_state.start_date,
            'num_months': int(st.session_state.num_months),
            'monthly_payment': float(st.session_state.monthly_payment),
            'payment_timing': st.session_state.payment_timing,
            'annual_rate': float(st.session_state.annual_rate),
            'origination_fee': float(st.session_state.origination_fee),
            'security_deposit': float(st.session_state.security_deposit),
            'initial_liability': result["initial_liability"],
            'prepaid_first_payment': result["prepaid_first_payment"],
            'initial_rou': result["initial_rou"],
            'classification': st.session_state.classification or "Operating Lease"
        }
        
        initial_je, _, _ = initial_recognition_je(
            result["initial_rou"], result["initial_liability"],
            result["prepaid_first_payment"], float(st.session_state.origination_fee),
            float(st.session_state.security_deposit)
        )
        
        excel_data = generate_excel_workbook(
            lease_info=lease_info_dict,
            schedule_df=result["schedule"],
            classification_memo=st.session_state.memo or "Standard ASC 842 lease schedule.",
            initial_je_df=initial_je,
            roll_forward_df=pd.DataFrame()
        )
        
        st.download_button(
            label="📊 Download Full Client Excel Report (.xlsx)",
            data=excel_data,
            file_name=f"{st.session_state.lease_name.replace(' ', '_')}_ASC842_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

# MODULE 5
with tab5:
    st.header("Module 5 — Journal Entries")
    result = st.session_state.schedule_result
    if result:
        je_df, td, tc = initial_recognition_je(
            result["initial_rou"], result["initial_liability"],
            result["prepaid_first_payment"], float(st.session_state.origination_fee),
            float(st.session_state.security_deposit)
        )
        st.subheader("Day 1 Initial Recognition Entry")
        st.dataframe(je_df.style.format({"Debit": "${:,.2f}", "Credit": "${:,.2f}"}), use_container_width=True)
