import streamlit as st
import pandas as pd

# -----------------------------
# PAGE CONFIG
# -----------------------------
st.set_page_config(
    page_title="Customer Contract Dashboard",
    page_icon="📊",
    layout="wide"
)

# -----------------------------
# PASSWORD PROTECTION
# -----------------------------
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

# -----------------------------
# FILE PATH
# -----------------------------
FILE_PATH = "Customer Contract and MRC Tracking (1).xlsx"

# -----------------------------
# LOAD EXCEL WORKBOOK
# -----------------------------
@st.cache_data
def load_workbook(path):
    xls = pd.ExcelFile(path)
    sheets = {}
    for sheet in xls.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet)
        df = df.dropna(how="all")
        df.columns = [str(col).strip() for col in df.columns]
        sheets[sheet] = df
    return sheets


sheets = load_workbook(FILE_PATH)

# -----------------------------
# TAB ORDER
# -----------------------------
sheet_names = list(sheets.keys())

ordered_tabs = []
if "Customer Status" in sheets:
    ordered_tabs.append("Customer Status")

for sheet in sheet_names:
    if sheet != "Customer Status":
        ordered_tabs.append(sheet)

# -----------------------------
# HEADER
# -----------------------------
st.title("📊 Customer Contract Dashboard")
st.caption("Customer contract and MRC tracking dashboard")

# -----------------------------
# TABS
# -----------------------------
tabs = st.tabs(ordered_tabs)

for i, tab_name in enumerate(ordered_tabs):
    with tabs[i]:
        df = sheets[tab_name]

        if tab_name == "Customer Status":
            st.subheader("Customer Status")

            # Optional quick stats
            total_rows = len(df)
            total_columns = len(df.columns)

            c1, c2 = st.columns(2)
            c1.metric("Total Records", total_rows)
            c2.metric("Total Columns", total_columns)

            st.dataframe(df, use_container_width=True)

        else:
            st.subheader(tab_name)
            st.dataframe(df, use_container_width=True)

        # Download button for each sheet
        csv = df.to_csv(index=False).encode("utf-8")
        st.download_button(
            label=f"Download {tab_name} as CSV",
            data=csv,
            file_name=f"{tab_name}.csv",
            mime="text/csv",
            key=f"download_{tab_name}"
        )
