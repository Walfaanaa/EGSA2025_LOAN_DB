import streamlit as st
import pandas as pd
import mysql.connector

st.set_page_config(
    page_title="EGSA2025 Short-Term Loan Dashboard",
    page_icon="💰",
    layout="wide"
)

st.title("EGSA2025 Short-Term Loan Dashboard")

# --------------------------------
# TiDB Cloud connection
# --------------------------------

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
        connection_timeout=20
    )


# --------------------------------
# Load loan records
# --------------------------------

@st.cache_data(ttl=60)
def load_loan_data():
    conn = get_connection()
    cursor = None

    try:
        cursor = conn.cursor()

        cursor.execute("""
            SELECT *
            FROM lottery_db.EGSA2025_short_term_loan
        """)

        rows = cursor.fetchall()
        columns = [col[0] for col in cursor.description]

        return pd.DataFrame(rows, columns=columns)

    finally:
        if cursor:
            cursor.close()
        conn.close()


# --------------------------------
# Dashboard
# --------------------------------

try:
    with st.spinner("Connecting to TiDB Cloud..."):

        df = load_loan_data()

    st.success("Successfully connected to TiDB Cloud!")

    col1, col2 = st.columns(2)

    with col1:
        st.metric("Total Loan Records", f"{len(df):,}")

    with col2:
        st.metric("Total Columns", len(df.columns))

    st.subheader("Short-Term Loan Records")

    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True
    )

    # Download report
    csv = df.to_csv(index=False).encode("utf-8")

    st.download_button(
        "Download Loan Report (CSV)",
        data=csv,
        file_name="EGSA2025_short_term_loan_report.csv",
        mime="text/csv"
    )

except Exception as e:
    st.error(f"Database connection failed: {e}")

if st.button("Refresh Data"):
    load_loan_data.clear()
    st.rerun()
