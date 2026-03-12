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
# HELPERS
# -------------------------------------------------
POSSIBLE_CODE_COLUMNS = [
    "Customer Code", "CustomerCode", "Cust Code", "Code", "Customer ID", "CustomerID"
]

POSSIBLE_NAME_COLUMNS = [
    "Customer Name", "Customer", "Name", "Client Name", "Account Name"
]

POSSIBLE_STATUS_COLUMNS = [
    "Status", "Customer Status", "Contract Status"
]

POSSIBLE_AM_COLUMNS = [
    "AM", "Account Manager", "Owner", "Sales Rep"
]


def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df = df.dropna(how="all")
    df.columns = [str(col).strip() for col in df.columns]
    return df


def find_first_matching_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    for col in candidates:
        if col in df.columns:
            return col

    lower_map = {str(col).strip().lower(): col for col in df.columns}
    for col in candidates:
        if col.lower() in lower_map:
            return lower_map[col.lower()]

    return None


def safe_str_series(series: pd.Series) -> pd.Series:
    return series.fillna("").astype(str).str.strip()


@st.cache_data
def load_workbook(path: str) -> dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(path)
    sheets = {}
    for sheet in xls.sheet_names:
        sheets[sheet] = normalize_df(pd.read_excel(path, sheet_name=sheet))
    return sheets


def filter_dataframe_ui(df: pd.DataFrame) -> pd.DataFrame:
    filtered_df = df.copy()

    st.markdown("### Filters")

    text_cols = [c for c in df.columns if df[c].dtype == "object" or str(df[c].dtype).startswith("string")]
    low_cardinality_cols = [c for c in text_cols if df[c].nunique(dropna=True) <= 50]

    cols = st.columns(4)

    # Search box
    with cols[0]:
        search_term = st.text_input("Search any text", placeholder="Customer name, code, notes...")

    # Status filter
    status_col = find_first_matching_column(df, POSSIBLE_STATUS_COLUMNS)
    with cols[1]:
        selected_statuses = []
        if status_col:
            statuses = sorted([x for x in safe_str_series(df[status_col]).unique() if x])
            selected_statuses = st.multiselect("Status", statuses)

    # Account manager filter
    am_col = find_first_matching_column(df, POSSIBLE_AM_COLUMNS)
    with cols[2]:
        selected_ams = []
        if am_col:
            ams = sorted([x for x in safe_str_series(df[am_col]).unique() if x])
            selected_ams = st.multiselect("Account Manager", ams)

    # Generic extra column filter
    with cols[3]:
        selectable_cols = [c for c in low_cardinality_cols if c not in {status_col, am_col}]
        chosen_extra_col = st.selectbox("Extra filter column", ["None"] + selectable_cols)

    if chosen_extra_col != "None":
        extra_values = sorted([x for x in safe_str_series(df[chosen_extra_col]).unique() if x])
        selected_extra_values = st.multiselect(f"{chosen_extra_col}", extra_values)
        if selected_extra_values:
            filtered_df = filtered_df[
                safe_str_series(filtered_df[chosen_extra_col]).isin(selected_extra_values)
            ]

    if status_col and selected_statuses:
        filtered_df = filtered_df[
            safe_str_series(filtered_df[status_col]).isin(selected_statuses)
        ]

    if am_col and selected_ams:
        filtered_df = filtered_df[
            safe_str_series(filtered_df[am_col]).isin(selected_ams)
        ]

    if search_term:
        mask = pd.Series(False, index=filtered_df.index)
        for col in filtered_df.columns:
            mask = mask | safe_str_series(filtered_df[col]).str.contains(search_term, case=False, na=False)
        filtered_df = filtered_df[mask]

    return filtered_df


def get_customer_related_rows(sheets: dict[str, pd.DataFrame], customer_code: str) -> dict[str, pd.DataFrame]:
    related = {}

    for sheet_name, df in sheets.items():
        code_col = find_first_matching_column(df, POSSIBLE_CODE_COLUMNS)
        if code_col:
            matches = df[safe_str_series(df[code_col]) == str(customer_code).strip()]
            if not matches.empty:
                related[sheet_name] = matches

    return related


# -------------------------------------------------
# LOAD DATA
# -------------------------------------------------
sheets = load_workbook(FILE_PATH)

sheet_names = list(sheets.keys())
ordered_tabs = ["Customer Status"] if "Customer Status" in sheets else []
ordered_tabs += [name for name in sheet_names if name != "Customer Status"]

# -------------------------------------------------
# APP HEADER
# -------------------------------------------------
st.title("📊 Customer Contract Dashboard")
st.caption("Customer contract and MRC tracking")

# -------------------------------------------------
# TABS
# -------------------------------------------------
tabs = st.tabs(ordered_tabs)

for i, tab_name in enumerate(ordered_tabs):
    with tabs[i]:
        df = sheets[tab_name]

        # -----------------------------------------
        # MAIN TAB: CUSTOMER STATUS
        # -----------------------------------------
        if tab_name == "Customer Status":
            st.subheader("Customer Status")

            code_col = find_first_matching_column(df, POSSIBLE_CODE_COLUMNS)
            name_col = find_first_matching_column(df, POSSIBLE_NAME_COLUMNS)
            status_col = find_first_matching_column(df, POSSIBLE_STATUS_COLUMNS)
            am_col = find_first_matching_column(df, POSSIBLE_AM_COLUMNS)

            filtered_df = filter_dataframe_ui(df)

            # KPI row
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Visible Records", len(filtered_df))
            k2.metric("Total Records", len(df))
            k3.metric("Columns", len(df.columns))
            k4.metric("Unique Customers", filtered_df[code_col].nunique() if code_col else len(filtered_df))

            st.markdown("### Customer List")

            selected_customer_code = None

            # Streamlit row selection
            event = st.dataframe(
                filtered_df,
                use_container_width=True,
                hide_index=True,
                on_select="rerun",
                selection_mode="single-row"
            )

            selected_rows = []
            try:
                selected_rows = event.selection.rows
            except Exception:
                selected_rows = []

            if selected_rows and code_col:
                selected_row_index = selected_rows[0]
                selected_customer_code = str(filtered_df.iloc[selected_row_index][code_col]).strip()

            # Backup selector in case row selection isn't used
            if code_col:
                st.markdown("### Open Customer Details")
                customer_options = filtered_df[code_col].dropna().astype(str).unique().tolist()
                customer_options = sorted(customer_options)

                manual_customer_code = st.selectbox(
                    "Select customer code",
                    options=[""] + customer_options,
                    index=0
                )

                if manual_customer_code:
                    selected_customer_code = manual_customer_code

            if selected_customer_code:
                st.divider()
                st.markdown(f"## Customer Details: `{selected_customer_code}`")

                # Main customer record from Customer Status
                customer_main = filtered_df[
                    safe_str_series(filtered_df[code_col]) == str(selected_customer_code).strip()
                ] if code_col else pd.DataFrame()

                if customer_main.empty and code_col:
                    customer_main = df[
                        safe_str_series(df[code_col]) == str(selected_customer_code).strip()
                    ]

                if not customer_main.empty:
                    record = customer_main.iloc[0]

                    detail_col1, detail_col2 = st.columns(2)

                    with detail_col1:
                        if code_col:
                            st.write(f"**Customer Code:** {record.get(code_col, '')}")
                        if name_col:
                            st.write(f"**Customer Name:** {record.get(name_col, '')}")
                        if status_col:
                            st.write(f"**Status:** {record.get(status_col, '')}")

                    with detail_col2:
                        if am_col:
                            st.write(f"**Account Manager:** {record.get(am_col, '')}")

                    with st.expander("Full Customer Status Record", expanded=True):
                        st.dataframe(customer_main, use_container_width=True, hide_index=True)

                # Show related rows from all sheets
                st.markdown("### Related Information Across All Sheets")
                related = get_customer_related_rows(sheets, selected_customer_code)

                shown_any = False
                for related_sheet, related_df in related.items():
                    st.markdown(f"#### {related_sheet}")
                    st.dataframe(related_df, use_container_width=True, hide_index=True)
                    shown_any = True

                if not shown_any:
                    st.info("No related records found in other sheets for this customer code.")

            # Download filtered CSV
            st.download_button(
                label="Download filtered Customer Status CSV",
                data=filtered_df.to_csv(index=False).encode("utf-8"),
                file_name="customer_status_filtered.csv",
                mime="text/csv",
                key="download_customer_status_filtered"
            )

        # -----------------------------------------
        # OTHER SHEETS
        # -----------------------------------------
        else:
            st.subheader(tab_name)

            filtered_df = df.copy()

            # Light filtering on all other tabs
            search_term = st.text_input(
                f"Search in {tab_name}",
                placeholder="Type to filter this sheet...",
                key=f"search_{tab_name}"
            )

            if search_term:
                mask = pd.Series(False, index=filtered_df.index)
                for col in filtered_df.columns:
                    mask = mask | safe_str_series(filtered_df[col]).str.contains(search_term, case=False, na=False)
                filtered_df = filtered_df[mask]

            st.dataframe(
                filtered_df,
                use_container_width=True,
                hide_index=True
            )

            code_col = find_first_matching_column(filtered_df, POSSIBLE_CODE_COLUMNS)
            if code_col:
                customer_codes = sorted(filtered_df[code_col].dropna().astype(str).unique().tolist())
                selected_code = st.selectbox(
                    f"View a customer from {tab_name}",
                    options=[""] + customer_codes,
                    index=0,
                    key=f"customer_select_{tab_name}"
                )

                if selected_code:
                    details_df = filtered_df[
                        safe_str_series(filtered_df[code_col]) == str(selected_code).strip()
                    ]
                    st.markdown(f"### Records for Customer Code: `{selected_code}`")
                    st.dataframe(details_df, use_container_width=True, hide_index=True)

            st.download_button(
                label=f"Download {tab_name} CSV",
                data=filtered_df.to_csv(index=False).encode("utf-8"),
                file_name=f"{tab_name}.csv",
                mime="text/csv",
                key=f"download_{tab_name}"
            )
