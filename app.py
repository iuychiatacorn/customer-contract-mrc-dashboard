
import io
from pathlib import Path
import pandas as pd
import streamlit as st
import plotly.express as px

def check_password():
    if "password_correct" not in st.session_state:
        st.session_state.password_correct = False

    if st.session_state.password_correct:
        return True

    password = st.text_input("Enter password", type="password")

    if password:
        if password == st.secrets["APP_PASSWORD"]:
            st.session_state.password_correct = True
            st.rerun()
        else:
            st.error("Incorrect password")

    return False

if not check_password():
    st.stop()

st.set_page_config(
    page_title="Customer Contract & MRC Dashboard",
    page_icon="📊",
    layout="wide",
)

DEFAULT_FILE = "Customer Contract and MRC Tracking (1).xlsx"


@st.cache_data(show_spinner=False)
def excel_file_bytes(path: str) -> bytes:
    return Path(path).read_bytes()


@st.cache_data(show_spinner=False)
def load_all_sheets(file_bytes: bytes):
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    sheets = {name: xl.parse(name, header=None) for name in xl.sheet_names}
    return sheets


def _to_datetime(series):
    return pd.to_datetime(series, errors="coerce")


def _to_numeric(series):
    return pd.to_numeric(series, errors="coerce")


def parse_customer_status(df_raw: pd.DataFrame) -> pd.DataFrame:
    headers = df_raw.iloc[0].tolist()
    df = df_raw.iloc[3:].copy()
    df.columns = headers
    df = df[df["Customer Code"].notna()].copy()
    df = df[df["Customer Name"].notna()].copy()
    df = df[~df["Customer Code"].astype(str).str.contains("Tier", na=False)].copy()

    date_cols = [
        "Contract Expiration", "QBR vCIO Generated", "Signed off by C/U",
        "Last BR", "Next BR", "Q1 QBR Closure TARGET"
    ]
    numeric_cols = ["MRR", "Seats"]

    for col in date_cols:
        if col in df.columns:
            df[col] = _to_datetime(df[col])
    for col in numeric_cols:
        if col in df.columns:
            df[col] = _to_numeric(df[col])

    df["Days to Expiration"] = (df["Contract Expiration"] - pd.Timestamp.today().normalize()).dt.days
    df["Expiring in 12 Months"] = df["Days to Expiration"].between(0, 365, inclusive="both")
    df["Expired"] = df["Days to Expiration"] < 0
    return df.reset_index(drop=True)


def parse_project_rate(df_raw: pd.DataFrame) -> pd.DataFrame:
    headers = df_raw.iloc[0].tolist()
    df = df_raw.iloc[2:].copy()
    df.columns = headers
    df = df[df["Customer Code"].notna()].copy()
    df = df[~df["Customer Code"].astype(str).str.contains("Tier", na=False)].copy()
    df["Project Rate"] = _to_numeric(df["Project Rate"])
    return df.reset_index(drop=True)


def parse_mrc_contracted(df_raw: pd.DataFrame) -> pd.DataFrame:
    headers = df_raw.iloc[4].tolist()
    df = df_raw.iloc[5:].copy()
    df.columns = headers
    df = df[df["Customer Code"].notna()].copy()

    date_cols = ["Contract Expiration", "Budget Creation"]
    numeric_cols = ["Tier", "MRR", "Current IT-Services MRC", "Proposed IT-Services", "MRC Difference"]

    for col in date_cols:
        if col in df.columns:
            df[col] = _to_datetime(df[col])
    for col in numeric_cols:
        if col in df.columns:
            df[col] = _to_numeric(df[col])

    if "MRC Difference" in df.columns:
        df["MRC Difference"] = df["Proposed IT-Services"] - df["Current IT-Services MRC"]

    df["Projected Uplift %"] = (
        df["MRC Difference"] / df["Current IT-Services MRC"]
    ).replace([pd.NA, pd.NaT], None)
    return df.reset_index(drop=True)


def parse_true_up(df_raw: pd.DataFrame) -> pd.DataFrame:
    headers = df_raw.iloc[0].tolist()
    df = df_raw.iloc[2:].copy()
    df.columns = headers

    current_client = None
    rows = []
    for _, row in df.iterrows():
        client = row.get("Client")
        desc = row.get("Support Description (Per Unit Per Month)")
        if pd.notna(client):
            current_client = str(client).strip()
        if pd.isna(desc):
            continue

        unit_pricing = _to_numeric(pd.Series([row.get("Unit Pricing")])).iloc[0]
        current_mrc = _to_numeric(pd.Series([row.get("Current MRC")])).iloc[0]
        current_qty = _to_numeric(pd.Series([row.get("Current Qty.")])).iloc[0]
        applied_discounts = _to_numeric(pd.Series([row.get("Applied Discounts")])).iloc[0]
        calc_true_up = _to_numeric(pd.Series([row.get("Current (TrueUP) MRC")])).iloc[0]

        if pd.isna(calc_true_up):
            calc_true_up = (unit_pricing or 0) * (current_qty or 0) - (applied_discounts or 0)

        rows.append(
            {
                "Client": current_client,
                "Support Description": str(desc).strip(),
                "Unit Pricing": unit_pricing,
                "Current MRC": current_mrc,
                "Current Qty": current_qty,
                "Applied Discounts": applied_discounts,
                "True Up MRC": calc_true_up,
            }
        )

    out = pd.DataFrame(rows)
    out = out[out["Client"].notna()].copy()
    return out.reset_index(drop=True)


def parse_hosting(df_raw: pd.DataFrame) -> pd.DataFrame:
    df = df_raw.iloc[2:].copy().reset_index(drop=True)
    df.columns = [
        "Tier", "Customer Code", "Customer Name", "Contract Expiration",
        "MRR", "Account Manager", "Col7", "Col8", "Col9",
        "Renewal Note", "Col11", "Budget Creation"
    ]
    df = df[df["Customer Code"].notna()].copy()
    df = df[~df["Customer Code"].astype(str).isin(["Hosting", "Recently Gone"])].copy()
    df["MRR"] = _to_numeric(df["MRR"])
    df["Contract Expiration"] = _to_datetime(df["Contract Expiration"])
    df["Budget Creation"] = _to_datetime(df["Budget Creation"])
    return df[["Customer Code", "Customer Name", "Contract Expiration", "MRR", "Account Manager", "Budget Creation"]].reset_index(drop=True)


def currency(value):
    if pd.isna(value):
        return "-"
    return f"${value:,.0f}"


def number(value):
    if pd.isna(value):
        return "-"
    return f"{value:,.0f}"


st.title("📊 Customer Contract & MRC Dashboard")
st.caption("Interactive view of customer contracts, MRR, project rates, and true-up planning from the uploaded Excel workbook.")

def get_default_file_path():
    here = Path(__file__).parent
    candidate = here / DEFAULT_FILE
    return candidate if candidate.exists() else None


default_path = get_default_file_path()
uploaded_file = st.sidebar.file_uploader("Upload Excel workbook", type=["xlsx"])

if uploaded_file is not None:
    file_bytes = uploaded_file.getvalue()
elif default_path is not None:
    file_bytes = excel_file_bytes(str(default_path))
    st.sidebar.success(f"Using bundled workbook: {DEFAULT_FILE}")
else:
    st.info("Upload the workbook to get started.")
    st.stop()

sheets = load_all_sheets(file_bytes)

customer_df = parse_customer_status(sheets["Customer status"])
project_df = parse_project_rate(sheets["Project Rate"])
contract_df = parse_mrc_contracted(sheets["MRC Contracted Rate"])
trueup_df = parse_true_up(sheets["MRC - True UP Planning"])
hosting_df = parse_hosting(sheets["Sheet1"])

st.sidebar.header("Filters")
am_options = sorted(x for x in customer_df["Account Manager"].dropna().unique())
tier_options = sorted(x for x in customer_df["Tier"].dropna().unique())

selected_am = st.sidebar.multiselect("Account Manager", am_options, default=am_options)
selected_tier = st.sidebar.multiselect("Tier", tier_options, default=tier_options)
customer_search = st.sidebar.text_input("Customer name contains")

filtered_customers = customer_df.copy()
if selected_am:
    filtered_customers = filtered_customers[filtered_customers["Account Manager"].isin(selected_am)]
if selected_tier:
    filtered_customers = filtered_customers[filtered_customers["Tier"].isin(selected_tier)]
if customer_search:
    filtered_customers = filtered_customers[
        filtered_customers["Customer Name"].astype(str).str.contains(customer_search, case=False, na=False)
    ]

customer_codes = set(filtered_customers["Customer Code"].dropna())
filtered_contracts = contract_df[contract_df["Customer Code"].isin(customer_codes)].copy()
filtered_projects = project_df[project_df["Customer Code"].isin(customer_codes)].copy()
filtered_trueup = trueup_df[trueup_df["Client"].isin(customer_codes)].copy()

total_customers = filtered_customers["Customer Code"].nunique()
total_mrr = filtered_customers["MRR"].sum()
avg_seats = filtered_customers["Seats"].mean()
expiring_12m = filtered_customers["Expiring in 12 Months"].fillna(False).sum()
projected_uplift = filtered_contracts["MRC Difference"].sum()

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Customers", number(total_customers))
k2.metric("Total MRR", currency(total_mrr))
k3.metric("Avg Seats", number(avg_seats))
k4.metric("Renewals in 12 Months", number(expiring_12m))
k5.metric("Projected MRC Uplift", currency(projected_uplift))

tab1, tab2, tab3, tab4, tab5 = st.tabs(
    ["Overview", "Contracts", "Project Rates", "True-Up Planning", "Hosting"]
)

with tab1:
    c1, c2 = st.columns([1, 1])

    with c1:
        mrr_by_am = (
            filtered_customers.groupby("Account Manager", dropna=False)["MRR"]
            .sum()
            .reset_index()
            .sort_values("MRR", ascending=False)
        )
        fig = px.bar(mrr_by_am, x="Account Manager", y="MRR", title="MRR by Account Manager")
        st.plotly_chart(fig, use_container_width=True)

        tier_summary = (
            filtered_customers.groupby("Tier", dropna=False)
            .agg(Customers=("Customer Code", "nunique"), MRR=("MRR", "sum"), Seats=("Seats", "sum"))
            .reset_index()
        )
        st.dataframe(tier_summary, use_container_width=True, hide_index=True)

    with c2:
        renewals = filtered_customers.dropna(subset=["Contract Expiration"]).copy()
        renewals["Renewal Month"] = renewals["Contract Expiration"].dt.to_period("M").astype(str)
        renewals_chart = (
            renewals.groupby("Renewal Month")["Customer Code"].count().reset_index(name="Count")
        )
        if not renewals_chart.empty:
            fig2 = px.line(renewals_chart, x="Renewal Month", y="Count", markers=True, title="Renewal timeline")
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.info("No contract expiration dates available for the current filter.")

        upcoming = (
            filtered_customers.loc[filtered_customers["Expiring in 12 Months"].fillna(False),
                                   ["Customer Code", "Customer Name", "Contract Expiration", "MRR", "Account Manager"]]
            .sort_values("Contract Expiration")
        )
        st.subheader("Upcoming renewals")
        st.dataframe(upcoming, use_container_width=True, hide_index=True)

    st.subheader("Customer detail")
    cols = [
        "Customer Code", "Customer Name", "Tier", "Account Manager", "MRR", "Seats",
        "Contract Expiration", "Contract Type", "CyberSuite?", "Firewall Equipment",
        "Backup (Immutable)", "Expired", "Expiring in 12 Months"
    ]
    existing_cols = [c for c in cols if c in filtered_customers.columns]
    st.dataframe(filtered_customers[existing_cols], use_container_width=True, hide_index=True)

with tab2:
    left, right = st.columns([1, 1])
    with left:
        uplift_chart = filtered_contracts.sort_values("MRC Difference", ascending=False)
        fig3 = px.bar(
            uplift_chart,
            x="Customer Code",
            y="MRC Difference",
            hover_data=["Customer Name"],
            title="Projected MRC Difference by customer",
        )
        st.plotly_chart(fig3, use_container_width=True)
    with right:
        st.subheader("Contract summary")
        summary = pd.DataFrame(
            {
                "Metric": [
                    "Current IT Services MRC",
                    "Proposed IT Services",
                    "Net Difference",
                ],
                "Value": [
                    filtered_contracts["Current IT-Services MRC"].sum(),
                    filtered_contracts["Proposed IT-Services"].sum(),
                    filtered_contracts["MRC Difference"].sum(),
                ],
            }
        )
        summary["Value"] = summary["Value"].map(currency)
        st.dataframe(summary, use_container_width=True, hide_index=True)

    show_cols = [
        "Tier", "Customer Code", "Customer Name", "Account Manager", "Contract Expiration",
        "Current IT-Services MRC", "Proposed IT-Services", "MRC Difference",
        "Alignment report completion date", "Strategy Review", "Strategy Road Map", "Budget Creation"
    ]
    show_cols = [c for c in show_cols if c in filtered_contracts.columns]
    st.dataframe(filtered_contracts[show_cols], use_container_width=True, hide_index=True)

with tab3:
    st.subheader("Project rate tracker")
    st.dataframe(filtered_projects, use_container_width=True, hide_index=True)

    numeric_rates = filtered_projects.dropna(subset=["Project Rate"]).copy()
    if not numeric_rates.empty:
        fig4 = px.histogram(numeric_rates, x="Project Rate", nbins=15, title="Distribution of current project rates")
        st.plotly_chart(fig4, use_container_width=True)

with tab4:
    st.subheader("True-up planning")
    agg_trueup = (
        filtered_trueup.groupby("Client", dropna=False)
        .agg(
            Lines=("Support Description", "count"),
            Total_True_Up_MRC=("True Up MRC", "sum"),
            Total_Current_MRC=("Current MRC", "sum"),
        )
        .reset_index()
        .sort_values("Total_True_Up_MRC", ascending=False)
    )

    if not agg_trueup.empty:
        fig5 = px.bar(agg_trueup, x="Client", y="Total_True_Up_MRC", title="True-up MRC by client")
        st.plotly_chart(fig5, use_container_width=True)

    st.dataframe(filtered_trueup, use_container_width=True, hide_index=True)

with tab5:
    st.subheader("Hosting and recent changes")
    if not hosting_df.empty:
        st.dataframe(hosting_df, use_container_width=True, hide_index=True)
        fig6 = px.bar(
            hosting_df.sort_values("MRR", ascending=False),
            x="Customer Code",
            y="MRR",
            hover_data=["Customer Name"],
            title="Hosting MRR by customer",
        )
        st.plotly_chart(fig6, use_container_width=True)
    else:
        st.info("No hosting data found in Sheet1.")

st.markdown("---")
st.caption("Tip: Push these files to a public GitHub repo, then deploy with Streamlit Community Cloud.")
