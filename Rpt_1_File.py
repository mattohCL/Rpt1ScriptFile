import os
import sys
from datetime import datetime
import pandas as pd
import oracledb
import win32com.client as win32

# Add helper module path
sys.path.append(r"C:\WFM_Scripting\Automation")

from scripthelper import (
    Config, Logger, ConnectionManager, GeneralFuncs,
    BigQueryManager, EmailManager, ApiFuncs
)
ghghghghghghghgh
# Initialize Oracle client
oracledb.init_oracle_client(lib_dir=r"C:\oracle\instantclient\instantclient_23_5")

# === Initialization ===
rpt_id = 1
config = Config(rpt_id=rpt_id)
logger = Logger(config)
connection_manager = ConnectionManager(config)
general_funcs = GeneralFuncs(config)
bigquery_manager = BigQueryManager(config)
email_manager = EmailManager(config)
api_funcs = ApiFuncs(config)

# === Constants ===
OUTPUT_DIR = r"C:\Scripting\Rpt_1"
os.makedirs(OUTPUT_DIR, exist_ok=True)

SHARED_MAILBOX = "TAX-RA-CoreLogicCXWFM@cotality.com"
TEST_EMAIL = ["mattoh@cotality.com"]

SQL_FILE_TAXP = r"C:\Scripting\Automation\PayeeAdminEmailTAXP1.sql"
SQL_FILE_TAXS = r"C:\Scripting\Automation\PayeeAdminEmailTAXS4.sql"
RECIPIENT_SQL_FILE = r"C:\Scripting\Automation\GetReportDL.sql"
BUSINESS_DAY_SQL_FILE = r"C:\WFM_Scripting\Automation\GBQ_is_today_Business_day.sql"

TAXP_DSN = "tax-taxp1-p1-oc-uc1-reg.data.solutions.corelogic.com:1525/TAX_LHA_RO"
TAXS_DSN = "tax-taxuat-u5-oc-uc1-uat.data.solutions.corelogic.com:1525/TAX_UAT"

# === Helper Functions ===

def connect_oracle(user_env, pass_env, dsn):
    """Get Oracle connection via environment variables."""
    try:
        user = os.getenv(user_env)
        pwd = os.getenv(pass_env)
        return connection_manager.connect_to_oracle(username=user, password=pwd, dsn=dsn)
    except Exception as e:
        logger.error(f"Oracle connection failed ({dsn}): {e}")
        email_manager.send_teams_notification(f"Oracle connection failed: {dsn}\n{e}")
        raise
        
def run_oracle_sql(sql_query, connection):
    """Execute a SQL query on the given Oracle connection and return a pandas DataFrame."""
    cursor = None
    try:
        cursor = connection.cursor()
        cursor.execute(sql_query)
        rows = cursor.fetchall()
        columns = [desc[0] for desc in cursor.description]
        df = pd.DataFrame(rows, columns=columns)
        return df
    finally:
        if cursor:
            cursor.close()

def fetch_data(sql_file, connection):
    """Execute Oracle SQL and return results as a DataFrame."""
    try:
        sql, _ = general_funcs.process_sql_input(sql_file)
        sql = sql.strip().rstrip(";")
        return run_oracle_sql(sql, connection)
    except Exception as e:
        logger.error(f"SQL execution failed ({sql_file}): {e}")
        raise

def format_html_table(df, title):
    """Return HTML formatted table (or fallback message)."""
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
    """Retrieve email list from BigQuery."""
    try:
        sql, _ = general_funcs.process_sql_input(RECIPIENT_SQL_FILE)
        sql = sql.replace("INSERTREPID", str(rpt_id))
        df = bigquery_manager.run_gbq_sql(sql, return_dataframe=True)
        if df.empty or "Email_Addr" not in df.columns:
            raise ValueError("No recipient emails found.")
        return df["Email_Addr"].tolist()
    except Exception as e:
        logger.warning(f"Falling back to TEST_EMAIL due to error: {e}")
        email_manager.send_teams_notification(f"Error fetching recipient emails: {e}")
        return TEST_EMAIL

def send_email(to_emails, html_body, attachments):
    """Send HTML email with optional attachments."""
    try:
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.SentOnBehalfOfName = SHARED_MAILBOX
        mail.Subject = f"Payees Pending Approval - Daily Report {datetime.now():%Y_%m_%d}"
        mail.HTMLBody = html_body
        mail.To = ";".join(to_emails)
        for file in attachments:
            mail.Attachments.Add(Source=file)
        mail.Send()
        logger.info("Email sent successfully.")
    except Exception as e:
        logger.error(f"Failed to send email: {e}")
        raise

def is_today_business_day():
    """Run the business day check query on BigQuery and return True/False."""
    try:
        with open(BUSINESS_DAY_SQL_FILE, "r") as f:
            sql = f.read()
        df = bigquery_manager.run_gbq_sql(sql, return_dataframe=True)
        if not df.empty and 'bus_day' in df.columns:
            return bool(df.loc[0, 'bus_day'])
        else:
            logger.warning("Business day check query returned no results or missing 'bus_day' column.")
            return False
    except Exception as e:
        logger.error(f"Failed to run business day check: {e}")
        email_manager.send_teams_notification(f"Business day check failed: {e}")
        return False

def main():
    conn_taxp = conn_taxs = None
    try:
        logger.info("Rpt_1 started.")

        # Connect BigQuery early for business day check
        connection_manager.connect_to_gbq()

        # Check if today is business day
        if not is_today_business_day():
            logger.info("Today is NOT a business day. Exiting without sending report.")
            return

        # TAXP
        conn_taxp = connect_oracle("cxwfm_taxp1_username", "cxwfm_taxp1_password", TAXP_DSN)
        df_taxp = fetch_data(SQL_FILE_TAXP, conn_taxp)
        logger.info(f"TAXP rows: {len(df_taxp)}")

        # TAXS
        conn_taxs = connect_oracle("cxwfm_taxs4_username", "cxwfm_taxs4_password", TAXS_DSN)
        df_taxs = fetch_data(SQL_FILE_TAXS, conn_taxs)
        logger.info(f"TAXS rows: {len(df_taxs)}")

        # Save Excel files
        attachments = []
        if not df_taxp.empty:
            file_taxp = os.path.join(OUTPUT_DIR, "PROD.xlsx")
            df_taxp.to_excel(file_taxp, index=False)
            attachments.append(file_taxp)

        if not df_taxs.empty:
            file_taxs = os.path.join(OUTPUT_DIR, "STAGE.xlsx")
            df_taxs.to_excel(file_taxs, index=False)
            attachments.append(file_taxs)

        if df_taxp.empty and df_taxs.empty:
            logger.info("No data to email.")
            return

        # Email body with greeting and sign-off
        html = (
            "<p>Afternoon,</p><br>"
            + format_html_table(df_taxp, "PROD")
            + "<br><br>"
            + format_html_table(df_taxs, "STAGE")
            + "<br><p>Best,</p>"
        )

        to_emails = fetch_recipient_emails()
        logger.info(f"Email recipients: {to_emails}")
        send_email(TEST_EMAIL, html, attachments)

        logger.info("Rpt_1 completed successfully.")
        bigquery_manager.update_log_in_bigquery()

    except Exception as e:
        logger.error(f"Rpt_1 failed: {e}")
        email_manager.send_teams_notification(f"Rpt_1 failed: {e}")
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
