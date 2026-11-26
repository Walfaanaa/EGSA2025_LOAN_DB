import io
from datetime import date
from zoneinfo import ZoneInfo

import mysql.connector
import pandas as pd
import streamlit as st


# ============================================================
# PAGE CONFIGURATION
# ============================================================
st.set_page_config(
    page_title="EGSA2025 | Short-Term Loan Dashboard",
    page_icon="💰",
    layout="wide",
)

st.title("💰 EGSA2025 Short-Term Loan Portfolio")
st.caption("Live loan portfolio reporting | TiDB Cloud")


# ============================================================
# DATABASE CONNECTION
# Required Streamlit Secrets:
# DB_HOST, DB_PORT, DB_USER, DB_PASSWORD, DB_SSL_CA
# ============================================================
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


# ============================================================
# LOAD DATA FROM TIDB
# ============================================================
@st.cache_data(ttl=60, show_spinner=False)
def load_loan_data():

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

    connection = get_connection()
    cursor = None

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


# ============================================================
# DATA PREPARATION
# ============================================================
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
        df[column] = pd.to_numeric(
            df[column],
            errors="coerce",
        )

    # Convert dates
    df["business_date_parsed"] = pd.to_datetime(
        df["business_date"],
        format="mixed",
        dayfirst=False,
        errors="coerce",
    )

    df["due_date_parsed"] = pd.to_datetime(
        df["due_date"],
        format="mixed",
        dayfirst=False,
        errors="coerce",
    )

    # Text cleaning
    df["loan_type"] = (
        df["loan_type"]
        .fillna("Unspecified")
        .astype(str)
        .str.strip()
    )

    df["status"] = (
        df["status"]
        .fillna("Unspecified")
        .astype(str)
        .str.strip()
    )

    df["status_key"] = df["status"].str.casefold()

    # Based on the supplied sample, collection_amount represents
    # expected repayment, not actual cash collected.
    df["expected_repayment"] = df["collection_amount"]

    return df


# ============================================================
# FORMATTING
# ============================================================
def format_money(value):

    if pd.isna(value):
        value = 0

    return f"{value:,.2f}"


# ============================================================
# EXCEL EXPORT
# ============================================================
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

    with pd.ExcelWriter(
        output,
        engine="openpyxl",
    ) as writer:

        dataframe[export_columns].to_excel(
            writer,
            index=False,
            sheet_name="Loan Details",
        )

    output.seek(0)

    return output.getvalue()


# ============================================================
# ETHIOPIA LOCAL DATE
# ============================================================
def get_local_date():

    try:
        return pd.Timestamp.now(
            tz=ZoneInfo("Africa/Addis_Ababa")
        ).date()

    except Exception:
        return date.today()


# ============================================================
# MAIN DASHBOARD
# ============================================================
try:

    with st.spinner("Connecting to TiDB Cloud and loading loan records..."):
        raw_data = load_loan_data()

    if raw_data.empty:

        st.warning("No loan records were found in the database.")
        st.stop()

    df = prepare_data(raw_data)

    today = get_local_date()

    # ========================================================
    # SIDEBAR FILTERS
    # ========================================================
    st.sidebar.title("🔎 Report Filters")

    st.sidebar.caption(
        f"Reporting date: {today.strftime('%d-%b-%Y')}"
    )

    loan_type_options = sorted(
        df["loan_type"].dropna().unique().tolist()
    )

    selected_loan_types = st.sidebar.multiselect(
        "Loan Type",
        options=loan_type_options,
        default=loan_type_options,
    )

    status_options = sorted(
        df["status"].dropna().unique().tolist()
    )

    selected_statuses = st.sidebar.multiselect(
        "Loan Status",
        options=status_options,
        default=status_options,
    )

    member_search = st.sidebar.text_input(
        "Search Member ID",
        placeholder="Enter member ID",
    )

    # Filter by loan type and status
    filtered = df[
        df["loan_type"].isin(selected_loan_types)
        & df["status"].isin(selected_statuses)
    ].copy()

    # Member ID filter
    if member_search.strip():

        filtered = filtered[
            filtered["id"]
            .astype("Int64")
            .astype(str)
            .str.contains(
                member_search.strip(),
                case=False,
                na=False,
            )
        ]

    # Business date filter
    business_dates = df["business_date_parsed"].dropna()

    if not business_dates.empty:

        min_date = business_dates.min().date()
        max_date = business_dates.max().date()

        selected_date_range = st.sidebar.date_input(
            "Business Date Range",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date,
        )

        if (
            isinstance(selected_date_range, (tuple, list))
            and len(selected_date_range) == 2
        ):

            start_date, end_date = selected_date_range

            filtered = filtered[
                filtered["business_date_parsed"].dt.date.between(
                    start_date,
                    end_date,
                )
            ]

    if filtered.empty:

        st.warning(
            "No records match the selected filters. "
            "Please change your filter selections."
        )

        st.stop()

    # ========================================================
    # DUE DATE MONITORING
    # ========================================================
    completed_mask = filtered["status_key"].eq("completed")

    filtered["days_to_due"] = (
        filtered["due_date_parsed"].dt.date - today
    ).apply(
        lambda value: value.days
        if pd.notna(value)
        else None
    )

    filtered["due_monitoring"] = "Date unavailable"

    filtered.loc[
        completed_mask,
        "due_monitoring",
    ] = "Completed"

    active_mask = (
        ~completed_mask
        & filtered["due_date_parsed"].notna()
    )

    filtered.loc[
        active_mask & (filtered["days_to_due"] < 0),
        "due_monitoring",
    ] = "Overdue"

    filtered.loc[
        active_mask
        & filtered["days_to_due"].between(0, 30),
        "due_monitoring",
    ] = "Due within 30 days"

    filtered.loc[
        active_mask & (filtered["days_to_due"] > 30),
        "due_monitoring",
    ] = "Not yet due"

    # ========================================================
    # PORTFOLIO SUMMARY
    # ========================================================
    st.header("📊 Portfolio Summary")

    total_loans = int(
        filtered["loan_id"].nunique()
    )

    total_members = int(
        filtered["id"].nunique()
    )

    total_disbursed = filtered[
        "disbursed_amount"
    ].sum()

    total_interest = filtered[
        "interest_amount"
    ].sum()

    expected_repayment = filtered[
        "expected_repayment"
    ].sum()

    completed_count = int(
        completed_mask.sum()
    )

    in_progress_count = int(
        filtered["status_key"].eq("in progress").sum()
    )

    overdue_count = int(
        filtered["due_monitoring"].eq("Overdue").sum()
    )

    # First KPI row
    k1, k2, k3, k4 = st.columns(4)

    k1.metric(
        "Total Loans",
        f"{total_loans:,}",
    )

    k2.metric(
        "Unique Members",
        f"{total_members:,}",
    )

    k3.metric(
        "Principal Disbursed",
        format_money(total_disbursed),
    )

    k4.metric(
        "Expected Repayment",
        format_money(expected_repayment),
    )

    # Second KPI row
    k5, k6, k7, k8 = st.columns(4)

    k5.metric(
        "Expected Interest",
        format_money(total_interest),
    )

    k6.metric(
        "Completed Loans",
        f"{completed_count:,}",
    )

    k7.metric(
        "In-Progress Loans",
        f"{in_progress_count:,}",
    )

    k8.metric(
        "Overdue Loans",
        f"{overdue_count:,}",
    )

    st.info(
        "Reporting note: Expected repayment is based on the "
        "collection_amount column. Actual collections and outstanding "
        "balances require verified payment transaction records."
    )

    st.divider()

    # ========================================================
    # LOAN STATUS AND LOAN TYPE ANALYSIS
    # ========================================================
    st.header("📈 Loan Portfolio Analysis")

    left, right = st.columns(2)

    with left:

        st.subheader("Loans by Status")

        status_summary = (
            filtered.groupby("status")
            .agg(
                Loan_Count=("loan_id", "nunique")
            )
            .sort_values(
                "Loan_Count",
                ascending=False,
            )
        )

        st.bar_chart(status_summary)

    with right:

        st.subheader("Principal by Loan Type")

        type_chart = (
            filtered.groupby("loan_type")
            .agg(
                Principal=("disbursed_amount", "sum")
            )
            .sort_values(
                "Principal",
                ascending=False,
            )
        )

        st.bar_chart(type_chart)

    # Loan type summary
    st.subheader("Loan Type Summary")

    type_summary = (
        filtered.groupby("loan_type")
        .agg(
            Loan_Count=("loan_id", "nunique"),
            Principal_Disbursed=("disbursed_amount", "sum"),
            Expected_Interest=("interest_amount", "sum"),
            Expected_Repayment=("expected_repayment", "sum"),
        )
        .reset_index()
        .sort_values(
            "Principal_Disbursed",
            ascending=False,
        )
    )

    st.dataframe(
        type_summary,
        use_container_width=True,
        hide_index=True,
    )

    # ========================================================
    # MONTHLY DISBURSEMENT ANALYSIS
    # ========================================================
    st.header("📅 Monthly Disbursement Trend")

    monthly_data = filtered.dropna(
        subset=["business_date_parsed"]
    ).copy()

    if not monthly_data.empty:

        monthly_data["Month"] = (
            monthly_data["business_date_parsed"]
            .dt.to_period("M")
            .astype(str)
        )

        monthly_summary = (
            monthly_data.groupby("Month")
            .agg(
                Principal_Disbursed=(
                    "disbursed_amount",
                    "sum",
                ),
                Expected_Interest=(
                    "interest_amount",
                    "sum",
                ),
                Expected_Repayment=(
                    "expected_repayment",
                    "sum",
                ),
                Loan_Count=(
                    "loan_id",
                    "nunique",
                ),
            )
            .sort_index()
        )

        st.line_chart(
            monthly_summary[
                [
                    "Principal_Disbursed",
                    "Expected_Repayment",
                ]
            ]
        )

        st.subheader("Monthly Summary Details")

        st.dataframe(
            monthly_summary.reset_index(),
            use_container_width=True,
            hide_index=True,
        )

    else:

        st.info(
            "No valid business dates are available for monthly analysis."
        )

    # ========================================================
    # DUE DATE MONITORING
    # ========================================================
    st.header("⏰ Due-Date Monitoring")

    due_summary = (
        filtered.groupby("due_monitoring")
        .agg(
            Loan_Count=("loan_id", "nunique")
        )
        .sort_values(
            "Loan_Count",
            ascending=False,
        )
    )

    st.dataframe(
        due_summary.reset_index(),
        use_container_width=True,
        hide_index=True,
    )

    due_records = filtered[
        filtered["due_monitoring"].isin(
            [
                "Overdue",
                "Due within 30 days",
            ]
        )
    ].copy()

    if not due_records.empty:

        st.subheader("Overdue and Upcoming Loans")

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

        st.success(
            "No overdue or upcoming loans were found "
            "for the selected filters."
        )

    # ========================================================
    # MEMBER-LEVEL ANALYSIS
    # ========================================================
    st.header("👥 Member-Level Loan Summary")

    member_summary = (
        filtered.groupby("id", dropna=False)
        .agg(
            Loan_Count=("loan_id", "nunique"),
            Principal_Disbursed=("disbursed_amount", "sum"),
            Expected_Interest=("interest_amount", "sum"),
            Expected_Repayment=("expected_repayment", "sum"),
        )
        .reset_index()
        .sort_values(
            "Principal_Disbursed",
            ascending=False,
        )
    )

    st.dataframe(
        member_summary,
        use_container_width=True,
        hide_index=True,
    )

    # ========================================================
    # DETAILED LOAN RECORDS
    # ========================================================
    st.header("📋 Detailed Loan Records")

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

    # ========================================================
    # REPORT DOWNLOADS
    # ========================================================
    st.header("📥 Download Reports")

    csv_data = filtered[
        display_columns
    ].to_csv(
        index=False
    ).encode("utf-8-sig")

    excel_data = make_excel_report(filtered)

    download_col1, download_col2 = st.columns(2)

    with download_col1:

        st.download_button(
            label="⬇️ Download CSV Report",
            data=csv_data,
            file_name="EGSA2025_Short_Term_Loan_Report.csv",
            mime="text/csv",
            use_container_width=True,
        )

    with download_col2:

        st.download_button(
            label="⬇️ Download Excel Report",
            data=excel_data,
            file_name="EGSA2025_Short_Term_Loan_Report.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True,
        )

    # ========================================================
    # REFRESH DASHBOARD
    # ========================================================
    st.divider()

    refresh_col, info_col = st.columns([1, 3])

    with refresh_col:

        if st.button(
            "🔄 Refresh Dashboard",
            use_container_width=True,
        ):

            load_loan_data.clear()
            st.rerun()

    with info_col:

        st.caption(
            f"Records displayed: {len(filtered):,} | "
            "Automatic cache refresh: 60 seconds"
        )


# ============================================================
# ERROR HANDLING
# ============================================================
except KeyError as error:

    st.error(
        f"Missing Streamlit secret: {error}. "
        "Please verify your database secrets."
    )

except mysql.connector.Error as error:

    st.error(
        f"TiDB database connection/query error: {error}"
    )

except Exception as error:

    st.error(
        f"Dashboard error: {error}"
    )
