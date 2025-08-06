import os
import sys
from datetime import datetime
import pandas as pd

from scripthelper import (
    Config, Logger, ConnectionManager, GeneralFuncs,
    BigQueryManager, ApiFuncs, EmailManager
)

# === Initialization ===
rpt_id = 1
config = Config(rpt_id=rpt_id)
logger = Logger(config)
connection_manager = ConnectionManager(config)
general_funcs = GeneralFuncs(config)
bigquery_manager = BigQueryManager(config)
api_funcs = ApiFuncs(config)
email_manager = EmailManager(config)

# === Constants ===
OUTPUT_DIR = r"C:\MyScriptFiles"  
os.makedirs(OUTPUT_DIR, exist_ok=True)

TEST_EMAIL = ["mattoh@cotality.com"]   #for testing purpose or fallback

SQL_FILE_TAXP = r"C:\WFM_Scripting\Automation\PayeeAdminEmailTAXP1.sql"
SQL_FILE_TAXS = r"C:\WFM_Scripting\Automation\PayeeAdminEmailTAXS4.sql"
RECIPIENT_SQL_FILE = r"C:\WFM_Scripting\Automation\GetReportDL.sql"
BUSINESS_DAY_SQL_FILE = r"C:\WFM_Scripting\Automation\GBQ_is_today_Business_day.sql"

# === Helper Functions ===
def format_html_table(df, title):
    if df.empty:
        return f"<h3>{title}</h3><p>No data available.</p>"
    style = """
    <style>
        table { border-collapse: collapse; width: 100%; font-family: Arial; }
        th, td { border: 1px solid black; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
    </style>
    """
    table = df.to_html(index=False, border=0)
    return f"{style}<h3>{title}</h3>{table}"

def fetch_recipient_emails():
    try:
        sql, _ = general_funcs.process_sql_input(RECIPIENT_SQL_FILE)
        sql = sql.replace("INSERTREPID", str(rpt_id))
        df = bigquery_manager.run_gbq_sql(sql, return_dataframe=True)
        if df.empty or "Email_Addr" not in df.columns:
            raise ValueError("No recipient emails found.")
        return df["Email_Addr"].tolist()
    except Exception as e:
        logger.warning(f"Falling back to TEST_EMAIL due to error: {e}")
        return TEST_EMAIL

def is_today_business_day():
    try:
        df = bigquery_manager.run_gbq_sql(BUSINESS_DAY_SQL_FILE, return_dataframe=True)
        if not df.empty and 'bus_day' in df.columns:
            return bool(df.loc[0, 'bus_day'])
        else:
            logger.warning("Business day check query returned no results or missing 'bus_day' column.")
            return False
    except Exception as e:
        logger.error(f"Failed to run business day check: {e}")
        return False

# === Main Logic ===
def main():
    conn_taxp = conn_taxs = None
    try:
        logger.info("Rpt_1 started.")

        if not is_today_business_day():
            logger.info("Today is NOT a business day. Exiting.")
            return

        conn_taxp = connection_manager.connect_to_oracle(db_connection="taxp")
        df_taxp = api_funcs.fetch_oracle_data(sql_input=SQL_FILE_TAXP, connection=conn_taxp, return_dataframe=True)
        logger.info(f"TAXP rows: {len(df_taxp)}")

        conn_taxs = connection_manager.connect_to_oracle(db_connection="taxs")
        df_taxs = api_funcs.fetch_oracle_data(sql_input=SQL_FILE_TAXS, connection=conn_taxs, return_dataframe=True)
        logger.info(f"TAXS rows: {len(df_taxs)}")

        if df_taxp.empty and df_taxs.empty:
            logger.info("No data to send.")
            return

        html = (
            "<p>Afternoon,</p><br>"
            + format_html_table(df_taxp, "PROD")
            + "<br><br>"
            + format_html_table(df_taxs, "STAGE")
            + "<br><p>Best,</p>"
        )

        to_emails = fetch_recipient_emails()
        logger.info(f"Recipients: {to_emails}")

        # Send email via EmailManager (HTML body, recipients)
        email_manager.send_email(
            subject=f"Payees Pending Approval - Daily Report {datetime.now():%m-%d-%Y}",
            body=html,
            is_html=True,
            recipient_emails=to_emails
        )

        bigquery_manager.update_log_in_bigquery()

        logger.info("Rpt_1 completed successfully.")

    except Exception as e:
        logger.error(f"Rpt_1 failed: {e}")
        sys.exit(1)
    finally:
        if conn_taxp:
            conn_taxp.close()
            logger.info("Closed TAXP connection.")
        if conn_taxs:
            conn_taxs.close()
            logger.info("Closed TAXS connection.")

if __name__ == "__main__":
    main()
