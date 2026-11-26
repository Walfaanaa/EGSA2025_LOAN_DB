from pathlib import Path
app_code = r'''import io
from datetime import date, timedelta
from zoneinfo import ZoneInfo

import mysql.connector
import pandas as pd
import streamlit as st


# ------------------------------------------------------------
# Page configuration
# ------------------------------------------------------------
st.set_page_config(
    page_title="EGSA2025 | Short-Term Loan Dashboard",
    page_icon="💰",
    layout="wide",
)

st.title("EGSA2025 Short-Term Loan Portfolio")
st.caption("Live reporting from TiDB Cloud")


# ------------------------------------------------------------
# TiDB connection
# Streamlit Cloud -> App settings -> Secrets must contain:
# DB_HOST, DB_PORT, DB_USER, DB_PASSWORD, DB_SSL_CA
# ------------------------------------------------------------
def get_connection():
    return mysql.connector.connect(
        host=st.secrets["DB_HOST"],
        port=int(st.secrets.get("DB_PORT", 4000)),
        user=st.secrets["DB_USER"],
        password=st.secrets["DB_PASSWORD"],
        database="lottery_db",
        ssl_ca=st.secrets["DB_SSL_CA"],
        ssl_verify_cert=True,
        ssl_verify_identity=True,
        connection_timeout=20,
    )


# ------------------------------------------------------------
# Load records from TiDB
# ------------------------------------------------------------
@st.cache_data(ttl=60, show_spinner=False)
def load_loan_data():
    connection = get_connection()
    cursor = None

    query = """
        SELECT
            loan_id,
            business_date,
            id,
            loan_type,
            disbursed_amount,
            interest_rate,
            interest_amount,
            collection_amount,
            from_account,
            to_account,
            due_date,
            phone_no,
            status
        FROM lottery_db.EGSA2025_short_term_loan
        ORDER BY loan_id DESC
    """

    try:
        cursor = connection.cursor()
        cursor.execute(query)
        rows = cursor.fetchall()
        columns = [column[0] for column in cursor.description]
        return pd.DataFrame(rows, columns=columns)
    finally:
        if cursor is not None:
            cursor.close()
        connection.close()


# ------------------------------------------------------------
# Data preparation
# The sample dates are in month/day/year format, e.g. 7/7/2026.
# ------------------------------------------------------------
def prepare_data(raw_df):
    df = raw_df.copy()

    numeric_columns = [
        "loan_id",
        "id",
        "disbursed_amount",
        "interest_rate",
        "interest_amount",
        "collection_amount",
        "from_account",
        "to_account",
        "phone_no",
    ]

    for column in numeric_columns:
        df[column] = pd.to_numeric(df[column], errors="coerce")

    df["business_date_parsed"] = pd.to_datetime(
        df["business_date"].astype("string"),
        format="mixed",
        dayfirst=False,
        errors="coerce",
    )

    df["due_date_parsed"] = pd.to_datetime(
        df["due_date"].astype("string"),
        format="mixed",
        dayfirst=False,
        errors="coerce",
    )

    df["loan_type"] = df["loan_type"].fillna("Unspecified").astype(str)
    df["status"] = df["status"].fillna("Unspecified").astype(str).str.strip()
    df["status_key"] = df["status"].str.casefold()

    # Collection_amount in the supplied sample equals principal + interest.
    # It is therefore treated as expected repayment, NOT actual cash received.
    df["expected_repayment"] = df["collection_amount"]

    return df


def format_money(value):
    return f"{value:,.2f}"


def make_excel_report(dataframe):
    output = io.BytesIO()

    export_columns = [
        "loan_id",
        "business_date",
        "id",
        "loan_type",
        "disbursed_amount",
        "interest_rate",
        "interest_amount",
        "collection_amount",
        "from_account",
        "to_account",
        "due_date",
        "phone_no",
        "status",
    ]

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        dataframe[export_columns].to_excel(
            writer,
            index=False,
            sheet_name="Loan Details",
        )

    output.seek(0)
    return output.getvalue()


# ------------------------------------------------------------
# Main dashboard
# ------------------------------------------------------------
try:
    with st.spinner("Loading loan records from TiDB Cloud..."):
        raw_data = load_loan_data()

    if raw_data.empty:
        st.info("The TiDB table contains no loan records.")
        st.stop()

    df = prepare_data(raw_data)

    # Local Ethiopia date for due-date monitoring
    today = date.today()
    try:
        today = pd.Timestamp.now(tz=ZoneInfo("Africa/Addis_Ababa")).date()
    except Exception:
        pass

    # --------------------------------------------------------
    # Sidebar filters
    # --------------------------------------------------------
    st.sidebar.header("Report Filters")

    loan_type_options = sorted(df["loan_type"].dropna().unique().tolist())
    selected_loan_types = st.sidebar.multiselect(
        "Loan type",
        options=loan_type_options,
        default=loan_type_options,
    )

    status_options = sorted(df["status"].dropna().unique().tolist())
    selected_statuses = st.sidebar.multiselect(
        "Loan status",
        options=status_options,
        default=status_options,
    )

    member_search = st.sidebar.text_input(
        "Search member ID",
        placeholder="Enter member ID",
    )

    business_dates = df["business_date_parsed"].dropna()

    filtered = df[
        df["loan_type"].isin(selected_loan_types)
        & df["status"].isin(selected_statuses)
    ].copy()

    if member_search.strip():
        filtered = filtered[
            filtered["id"].astype("Int64").astype(str).str.contains(
                member_search.strip(),
                case=False,
                na=False,
            )
        ]

    if not business_dates.empty:
        min_business_date = business_dates.min().date()
        max_business_date = business_dates.max().date()

        selected_date_range = st.sidebar.date_input(
            "Business date range",
            value=(min_business_date, max_business_date),
            min_value=min_business_date,
            max_value=max_business_date,
        )

        if isinstance(selected_date_range, (tuple, list)) and len(selected_date_range) == 2:
            start_date, end_date = selected_date_range
            filtered = filtered[
                filtered["business_date_parsed"].dt.date.between(
                    start_date, end_date
                )
            ]

    # --------------------------------------------------------
    # Derived due-date classifications
    # --------------------------------------------------------
    completed_mask = filtered["status_key"].eq("completed")
    filtered["days_to_due"] = (
        filtered["due_date_parsed"].dt.date - today
    ).apply(lambda value: value.days if pd.notna(value) else None)

    filtered["due_monitoring"] = "Date unavailable"
    filtered.loc[completed_mask, "due_monitoring"] = "Completed"

    active_mask = ~completed_mask & filtered["due_date_parsed"].notna()
    filtered.loc[
        active_mask & (filtered["days_to_due"] < 0),
        "due_monitoring",
    ] = "Overdue"
    filtered.loc[
        active_mask & filtered["days_to_due"].between(0, 30),
        "due_monitoring",
    ] = "Due within 30 days"
    filtered.loc[
        active_mask & (filtered["days_to_due"] > 30),
        "due_monitoring",
    ] = "Not yet due"

    # --------------------------------------------------------
    # KPI cards
    # --------------------------------------------------------
    st.subheader("Portfolio Summary")

    total_loans = int(filtered["loan_id"].nunique())
    total_members = int(filtered["id"].nunique())
    total_disbursed = filtered["disbursed_amount"].sum()
    total_interest = filtered["interest_amount"].sum()
    expected_repayment = filtered["expected_repayment"].sum()

    completed_count = int(completed_mask.sum())
    in_progress_count = int(filtered["status_key"].eq("in progress").sum())
    overdue_count = int(filtered["due_monitoring"].eq("Overdue").sum())

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total loans", f"{total_loans:,}")
    k2.metric("Unique members", f"{total_members:,}")
    k3.metric("Principal disbursed", format_money(total_disbursed))
    k4.metric("Expected repayment", format_money(expected_repayment))

    k5, k6, k7, k8 = st.columns(4)
    k5.metric("Expected interest", format_money(total_interest))
    k6.metric("Completed loans", f"{completed_count:,}")
    k7.metric("In-progress loans", f"{in_progress_count:,}")
    k8.metric("Overdue loans", f"{overdue_count:,}")

    st.caption(
        "Expected repayment is taken from collection_amount. "
        "Actual collections and outstanding balances cannot be calculated "
        "reliably until actual payment transactions are available."
    )

    st.divider()

    # --------------------------------------------------------
    # Loan status analysis
    # --------------------------------------------------------
    left, right = st.columns(2)

    with left:
        st.subheader("Loans by Status")
        status_summary = (
            filtered.groupby("status", dropna=False)
            .agg(Loan_Count=("loan_id", "nunique"))
            .sort_values("Loan_Count", ascending=False)
        )
        if not status_summary.empty:
            st.bar_chart(status_summary)

    with right:
        st.subheader("Portfolio by Loan Type")
        type_summary = (
            filtered.groupby("loan_type", dropna=False)
            .agg(
                Loan_Count=("loan_id", "nunique"),
                Principal=("disbursed_amount", "sum"),
                Expected_Repayment=("expected_repayment", "sum"),
            )
            .sort_values("Principal", ascending=False)
        )
        if not type_summary.empty:
            st.bar_chart(type_summary[["Principal"]])

    st.subheader("Loan Type Summary")
    st.dataframe(
        type_summary.reset_index(),
        use_container_width=True,
        hide_index=True,
    )

    # --------------------------------------------------------
    # Monthly disbursement analysis
    # --------------------------------------------------------
    st.subheader("Monthly Disbursement Trend")

    monthly_data = filtered.dropna(subset=["business_date_parsed"]).copy()

    if not monthly_data.empty:
        monthly_data["Month"] = (
            monthly_data["business_date_parsed"]
            .dt.to_period("M")
            .astype(str)
        )

        monthly_summary = (
            monthly_data.groupby("Month")
            .agg(
                Principal_Disbursed=("disbursed_amount", "sum"),
                Expected_Interest=("interest_amount", "sum"),
                Expected_Repayment=("expected_repayment", "sum"),
                Loan_Count=("loan_id", "nunique"),
            )
            .sort_index()
        )

        st.line_chart(
            monthly_summary[
                ["Principal_Disbursed", "Expected_Repayment"]
            ]
        )

        st.dataframe(
            monthly_summary.reset_index(),
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info("No valid business dates are available for monthly analysis.")

    # --------------------------------------------------------
    # Due-date monitoring
    # --------------------------------------------------------
    st.subheader("Due-Date Monitoring")

    due_summary = (
        filtered.groupby("due_monitoring", dropna=False)
        .agg(Loan_Count=("loan_id", "nunique"))
        .sort_values("Loan_Count", ascending=False)
    )

    st.dataframe(
        due_summary.reset_index(),
        use_container_width=True,
        hide_index=True,
    )

    due_records = filtered[
        filtered["due_monitoring"].isin(
            ["Overdue", "Due within 30 days"]
        )
    ].copy()

    if not due_records.empty:
        st.markdown("**Overdue and upcoming loans**")
        due_display = due_records[
            [
                "loan_id",
                "id",
                "loan_type",
                "disbursed_amount",
                "expected_repayment",
                "due_date",
                "status",
                "due_monitoring",
                "days_to_due",
            ]
        ].copy()

        st.dataframe(
            due_display,
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info("No overdue or upcoming loans found for the selected filters.")

    # --------------------------------------------------------
    # Member-level analysis
    # --------------------------------------------------------
    st.subheader("Member-Level Loan Summary")

    member_summary = (
        filtered.groupby("id", dropna=False)
        .agg(
            Loan_Count=("loan_id", "nunique"),
            Principal_Disbursed=("disbursed_amount", "sum"),
            Expected_Interest=("interest_amount", "sum"),
            Expected_Repayment=("expected_repayment", "sum"),
        )
        .reset_index()
        .sort_values("Principal_Disbursed", ascending=False)
    )

    st.dataframe(
        member_summary,
        use_container_width=True,
        hide_index=True,
    )

    # --------------------------------------------------------
    # Detailed loan records and exports
    # --------------------------------------------------------
    st.subheader("Detailed Loan Records")

    display_columns = [
        "loan_id",
        "business_date",
        "id",
        "loan_type",
        "disbursed_amount",
        "interest_rate",
        "interest_amount",
        "collection_amount",
        "from_account",
        "to_account",
        "due_date",
        "phone_no",
        "status",
    ]

    st.dataframe(
        filtered[display_columns],
        use_container_width=True,
        hide_index=True,
    )

    csv_data = filtered[display_columns].to_csv(index=False).encode("utf-8")
    excel_data = make_excel_report(filtered)

    download_col1, download_col2 = st.columns(2)

    with download_col1:
        st.download_button(
            "Download CSV Report",
            data=csv_data,
            file_name="EGSA2025_Short_Term_Loan_Report.csv",
            mime="text/csv",
            use_container_width=True,
        )

    with download_col2:
        st.download_button(
            "Download Excel Report",
            data=excel_data,
            file_name="EGSA2025_Short_Term_Loan_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.divider()

    # --------------------------------------------------------
    # Refresh and data timestamp
    # --------------------------------------------------------
    refresh_col, info_col = st.columns([1, 3])

    with refresh_col:
        if st.button("Refresh Dashboard", use_container_width=True):
            load_loan_data.clear()
            st.rerun()

    with info_col:
        st.caption(
            f"Records displayed: {len(filtered):,} | "
            f"Data cache duration: 60 seconds"
        )

except KeyError as error:
    st.error(
        f"Missing Streamlit secret: {error}. "
        "Check DB_HOST, DB_USER, DB_PASSWORD and DB_SSL_CA."
    )

except mysql.connector.Error as error:
    st.error(f"TiDB database connection/query error: {error}")

except Exception as error:
    st.error(f"Dashboard error: {error}")
'''

requirements = """streamlit
pandas
mysql-connector-python
openpyxl
"""

out_dir = Path("/mnt/data/egsa_loan_dashboard_update")
out_dir.mkdir(parents=True, exist_ok=True)
app_path = out_dir / "app.py"
req_path = out_dir / "requirements.txt"
app_path.write_text(app_code, encoding="utf-8")
req_path.write_text(requirements, encoding="utf-8")

print(f"Created: {app_path}")
print(f"Created: {req_path}")
