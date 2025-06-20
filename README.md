# Rpt_1 Automation Script

## Purpose
This Python script replaces `Rpt1ScriptFile.vbs` by automating the following:

- Checks if today is a business day using BigQuery
- Queries Oracle databases (TAXP & TAXS)
- Exports results to Excel if data exists
- Sends a formatted HTML email via Outlook shared mailbox
- Logs execution to BigQuery


## Requirements
- Python 3.10+
- Oracle Instant Client
- Access to GCP / BigQuery
- Outlook Desktop App (for email)
- `scripthelper` module (used internally for shared logic)


## How to Run
```bash
python Rpt_1_File.py
