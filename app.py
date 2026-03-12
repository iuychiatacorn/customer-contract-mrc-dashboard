import streamlit as st
import pandas as pd

# -------------------------------------------------
# PAGE CONFIG
# -------------------------------------------------
st.set_page_config(
    page_title="Customer Contract Dashboard",
    page_icon="📊",
    layout="wide"
)

# -------------------------------------------------
# PASSWORD PROTECTION
# -------------------------------------------------
def check_password():
    st.markdown(
        """
        <style>
        .center-box {
            text-align: center;
            margin-top: 20vh;
        }

        div[data-testid="stTextInput"] {
            max-width: 350px;
            margin: 0 auto;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    if "password_correct" not in st.session_state:
        st.session_state.password_correct = False

    if st.session_state.password_correct:
        return True

    col1, col2, col3 = st.columns([1, 2, 1])

    with col2:
        st.markdown('<div class="center-box">', unsafe_allow_html=True)
        st.markdown("## 🔐 Enter Password")
        st.caption("Authorized Access Only")

        password = st.text_input("", type="password", placeholder="Enter password")

        if password:
            if password == st.secrets["APP_PASSWORD"]:
                st.session_state.password_correct = True
                st.rerun()
            else:
                st.error("Incorrect password")

        st.markdown("</div>", unsafe_allow_html=True)

    return False


if not check_password():
    st.stop()

# -------------------------------------------------
# FILE PATH
# -------------------------------------------------
FILE_PATH = "Customer Contract and MRC Tracking (1).xlsx"

# -------------------------------------------------
# COLUMN CANDIDATES
# -------------------------------------------------
POSSIBLE_CODE_COLUMNS = ["Customer Code", "CustomerCode", "Cust Code", "Code", "Customer ID"]
POSSIBLE_NAME_COLUMNS = ["Customer Name", "Customer", "Name", "Account Name", "Client Name"]
POSSIBLE_STATUS_COLUMNS = ["Status", "Customer Status", "Contract Status"]
POSSIBLE_AM_COLUMNS = ["AM", "Account Manager", "Owner", "Sales Rep"]
POSSIBLE_CATEGORY_COLUMNS = ["Customer Category", "Category", "Tier", "Service Tier"]
POSSIBLE_MRR_COLUMNS = ["MRR", "MRC", "Monthly Recurring Revenue"]
POSSIBLE_IT_MRC_COLUMNS = ["Current IT-Services MRC", "Current IT Services MRC", "IT Services MRC", "Current MRC"]
POSSIBLE_CONTRACT_EXP_COLUMNS = ["Contract Expiration", "Expiration", "Contract Expiry", "Renewal Date"]

# -------------------------------------------------
# HELPERS
# -------------------------------------------------
def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df = df.dropna(how="all")
    df.columns = [str(col).strip() for col in df.columns]
    return df


def find_first_matching_column(df: pd.DataFrame, candidates: list[str]):
    for col in candidates:
        if col in df.columns:
            return col

    lowered = {str(col).strip().lower(): col for col in df.columns}
    for col in candidates:
        if col.lower() in lowered:
            return lowered[col.lower()]
    return None


def safe_str_series(series: pd.Series) -> pd.Series:
    return series.fillna("").astype(str).str.strip()


def to_numeric_series(series: pd.Series) -> pd.Series:
    cleaned = (
        series.astype(str)
        .str.replace("$", "", regex=False)
        .str.replace(",", "", regex=False)
        .str.strip()
    )
    return pd.to_numeric(cleaned, errors="coerce")


def format_currency(value):
    if pd.isna(value):
        return "$0.00"
    return f"${value:,.2f}"


@st.cache_data
def load_workbook(path: str):
    xls = pd.ExcelFile(path)
    sheets = {}
    for sheet in xls.sheet_names:
        sheets[sheet] = normalize_df(pd.read_excel(path, sheet_name=sheet))
    return sheets


def get_related_rows(sheets: dict[str, pd.DataFrame], customer_code: str):
    related = {}
    for sheet_name, df in sheets.items():
        code_col = find_first_matching_column(df, POSSIBLE_CODE_COLUMNS)
        if code_col:
            matches = df[safe_str_series(df[code_col]) == str(customer_code).strip()]
            if not matches.empty:
                related[sheet_name] = matches
    return related


def filter_customer_status(df: pd.DataFrame) -> pd.DataFrame:
    filtered_df = df.copy()

    code_col = find_first_matching_column(df, POSSIBLE_CODE_COLUMNS)
    name_col = find_first_matching_column(df, POSSIBLE_NAME_COLUMNS)
    status_col = find_first_matching_column(df, POSSIBLE_STATUS_COLUMNS)
    am_col = find_first_matching_column(df, POSSIBLE_AM_COLUMNS)
    category_col = find_first_matching_column(df, POSSIBLE_CATEGORY_COLUMNS)

    st.markdown("### Filters")
    c1, c2, c3, c4 = st.columns(4)

    with c1:
        search_term = st.text_input(
            "Search",
            placeholder="Customer code, name, manager..."
        )

    with c2:
        selected_statuses = []
        if status_col:
            options = sorted([x for x in safe_str_series(df[status_col]).unique() if x])
            selected_statuses = st.multiselect("Status", options)

    with c3:
        selected_ams = []
        if am_col:
            options = sorted([x for x in safe_str_series(df[am_col]).unique() if x])
            selected_ams = st.multiselect("Account Manager", options)

    with c4:
        selected_categories = []
        if category_col:
            options = sorted([x for x in safe_str_series(df[category_col]).unique() if x])
            selected_categories = st.multiselect("Category / Tier", options)

    if status_col and selected_statuses:
        filtered_df = filtered_df[safe_str_series(filtered_df[status_col]).isin(selected_statuses)]

    if am_col and selected_ams:
        filtered_df = filtered_df[safe_str_series(filtered_df[am_col]).isin(selected_ams)]

    if category_col and selected_categories:
        filtered_df = filtered_df[safe_str_series(filtered_df[category_col]).isin(selected_categories)]

    if search_term:
        mask = pd.Series(False, index=filtered_df.index)
        search_cols = [c for c in [code_col, name_col, am_col, status_col, category_col] if c]
        if not search_cols:
            search_cols = filtered_df.columns.tolist()

        for col in search_cols:
            mask = mask | safe_str_series(filtered_df[col]).str.contains(search_term, case=False, na=False)
        filtered_df = filtered_df[mask]

    return filtered_df


# -------------------------------------------------
# LOAD DATA
# -------------------------------------------------
sheets = load_workbook(FILE_PATH)

sheet_names = list(sheets.keys())
ordered_tabs = ["Customer Status"] if "Customer Status" in sheets else []
ordered_tabs += [sheet for sheet in sheet_names if sheet != "Customer Status"]

# -------------------------------------------------
# APP HEADER
# -------------------------------------------------
st.title("📊 Customer Contract Dashboard")
st.caption("Data and metric driven contract / MRC dashboard")

# -------------------------------------------------
# TABS
# -------------------------------------------------
tabs = st.tabs(ordered_tabs)

for i, tab_name in enumerate(ordered_tabs):
    with tabs[i]:
        df = sheets[tab_name]

        # -------------------------------------------------
        # MAIN TAB
        # -------------------------------------------------
        if tab_name == "Customer Status":
            st.subheader("Customer Status")

            code_col = find_first_matching_column(df, POSSIBLE_CODE_COLUMNS)
            name_col = find_first_matching_column(df, POSSIBLE_NAME_COLUMNS)
            status_col = find_first_matching_column(df, POSSIBLE_STATUS_COLUMNS)
            am_col = find_first_matching_column(df, POSSIBLE_AM_COLUMNS)
            category_col = find_first_matching_column(df, POSSIBLE_CATEGORY_COLUMNS)
            mrr_col = find_first_matching_column(df, POSSIBLE_MRR_COLUMNS)
            it_mrc_col = find_first_matching_column(df, POSSIBLE_IT_MRC_COLUMNS)
            contract_exp_col = find_first_matching_column(df, POSSIBLE_CONTRACT_EXP_COLUMNS)

            filtered_df = filter_customer_status(df)

            # Metrics
            total_customers = filtered_df[code_col].nunique() if code_col else len(filtered_df)

            total_mrr = 0.0
            if mrr_col:
                total_mrr = to_numeric_series(filtered_df[mrr_col]).fillna(0).sum()

            total_it_mrc = 0.0
            if it_mrc_col:
                total_it_mrc = to_numeric_series(filtered_df[it_mrc_col]).fillna(0).sum()

            active_customers = 0
            if status_col:
                active_customers = safe_str_series(filtered_df[status_col]).str.lower().isin(
                    ["active", "current", "live", "in service"]
                ).sum()

            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Customers", total_customers)
            k2.metric("Total MRR", format_currency(total_mrr))
            k3.metric("Current IT Services MRC", format_currency(total_it_mrc))
            k4.metric("Active / Current", int(active_customers))

            st.divider()

            # Main table
            st.markdown("### Customer Table")

            preferred_cols = [
                code_col, name_col, category_col, status_col,
                am_col, contract_exp_col, mrr_col, it_mrc_col
            ]
            display_cols = [c for c in preferred_cols if c and c in filtered_df.columns]

            if display_cols:
                table_df = filtered_df[display_cols].copy()
            else:
                table_df = filtered_df.copy()

            st.dataframe(
                table_df,
                use_container_width=True,
                hide_index=True
            )

            # Customer selector
            st.markdown("### Customer Details")

            selected_customer_code = None

            sel1, sel2 = st.columns(2)

            with sel1:
                if code_col:
                    code_options = sorted(filtered_df[code_col].dropna().astype(str).unique().tolist())
                    selected_customer_code = st.selectbox(
                        "Select customer code",
                        options=[""] + code_options,
                        index=0
                    )

            with sel2:
                selected_customer_name = ""
                if name_col:
                    name_options = sorted(filtered_df[name_col].dropna().astype(str).unique().tolist())
                    selected_customer_name = st.selectbox(
                        "Or select customer name",
                        options=[""] + name_options,
                        index=0
                    )
                    if selected_customer_name and code_col:
                        match = filtered_df[safe_str_series(filtered_df[name_col]) == selected_customer_name]
                        if not match.empty:
                            selected_customer_code = str(match.iloc[0][code_col]).strip()

            if selected_customer_code and code_col:
                customer_main = df[safe_str_series(df[code_col]) == str(selected_customer_code).strip()]

                if not customer_main.empty:
                    record = customer_main.iloc[0]

                    d1, d2, d3, d4 = st.columns(4)
                    if code_col:
                        d1.metric("Customer Code", str(record.get(code_col, "")))
                    if name_col:
                        d2.metric("Customer Name", str(record.get(name_col, "")))
                    if status_col:
                        d3.metric("Status", str(record.get(status_col, "")))
                    if category_col:
                        d4.metric("Category / Tier", str(record.get(category_col, "")))

                    d5, d6, d7, d8 = st.columns(4)
                    if am_col:
                        d5.metric("Account Manager", str(record.get(am_col, "")))
                    if contract_exp_col:
                        d6.metric("Contract Expiration", str(record.get(contract_exp_col, "")))
                    if mrr_col:
                        mrr_value = to_numeric_series(pd.Series([record.get(mrr_col, None)])).iloc[0]
                        d7.metric("MRR", format_currency(mrr_value if pd.notna(mrr_value) else 0))
                    if it_mrc_col:
                        it_value = to_numeric_series(pd.Series([record.get(it_mrc_col, None)])).iloc[0]
                        d8.metric("IT Services MRC", format_currency(it_value if pd.notna(it_value) else 0))

                    st.markdown("#### Full Customer Status Record")
                    st.dataframe(customer_main, use_container_width=True, hide_index=True)

                    related = get_related_rows(sheets, selected_customer_code)

                    st.markdown("#### Related Records Across Sheets")
                    for related_sheet, related_df in related.items():
                        with st.expander(f"{related_sheet} ({len(related_df)} row(s))", expanded=(related_sheet == "Customer Status")):
                            st.dataframe(related_df, use_container_width=True, hide_index=True)

            st.download_button(
                label="Download filtered Customer Status CSV",
                data=filtered_df.to_csv(index=False).encode("utf-8"),
                file_name="customer_status_filtered.csv",
                mime="text/csv",
                key="download_customer_status_filtered"
            )

        # -------------------------------------------------
        # OTHER SHEETS
        # -------------------------------------------------
        else:
            st.subheader(tab_name)

            search_term = st.text_input(
                f"Search in {tab_name}",
                placeholder="Filter this sheet...",
                key=f"search_{tab_name}"
            )

            filtered_df = df.copy()

            if search_term:
                mask = pd.Series(False, index=filtered_df.index)
                for col in filtered_df.columns:
                    mask = mask | safe_str_series(filtered_df[col]).str.contains(search_term, case=False, na=False)
                filtered_df = filtered_df[mask]

            # Basic metrics for each sheet
            m1, m2 = st.columns(2)
            m1.metric("Rows", len(filtered_df))
            m2.metric("Columns", len(filtered_df.columns))

            st.dataframe(
                filtered_df,
                use_container_width=True,
                hide_index=True
            )

            st.download_button(
                label=f"Download {tab_name} CSV",
                data=filtered_df.to_csv(index=False).encode("utf-8"),
                file_name=f"{tab_name}.csv",
                mime="text/csv",
                key=f"download_{tab_name}"
            )
