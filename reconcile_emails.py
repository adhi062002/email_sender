import re
import argparse
from pathlib import Path
from collections import defaultdict
from email.utils import parseaddr

import pandas as pd


# ============================================================
# CONFIGURATION
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

INPUT_DIR = BASE_DIR / "input"

OUTPUT_FILE = BASE_DIR / "email_reconciliation.xlsx"

EMAIL_COLUMN = "Email"

SUCCESS_MARKERS = [
    "[SENT]",
    "[SMTP ACCEPTED]",
]

EMAIL_PATTERN = re.compile(
    r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}"
)


# ============================================================
# EMAIL EXTRACTION
# ============================================================

def normalize_email(email):

    return email.strip().casefold()


def extract_emails(text):

    if pd.isna(text):
        return []

    matches = EMAIL_PATTERN.findall(str(text))

    result = []

    for email in matches:

        _, address = parseaddr(email)

        if address:
            result.append(normalize_email(address))

    return result


# ============================================================
# READ EXCEL FILES
# ============================================================

def load_excel_recipients():

    recipients = defaultdict(list)

    invalid_rows = []

    excel_files = sorted(INPUT_DIR.glob("*.xlsx"))

    if not excel_files:

        raise FileNotFoundError(
            f"No Excel files found in {INPUT_DIR}"
        )

    print("\n================================")
    print("READING EXCEL FILES")
    print("================================")

    for file in excel_files:

        print(f"Reading: {file.name}")

        df = pd.read_excel(file, dtype=object)

        if EMAIL_COLUMN not in df.columns:

            raise ValueError(
                f"Missing '{EMAIL_COLUMN}' column in {file.name}"
            )

        for index, row in df.iterrows():

            raw_email = row[EMAIL_COLUMN]

            emails = extract_emails(raw_email)

            if not emails:

                invalid_rows.append({

                    "Source File": file.name,
                    "Source Row": index + 2,
                    "Original Email": raw_email,
                    "Status": "INVALID EMAIL"

                })

                continue

            # Preserve the complete original Excel row.
            full_row = row.to_dict()

            for email in emails:

                recipients[email].append({

                    "Email": email,

                    "Source File": file.name,

                    "Source Row": index + 2,

                    "Name": row.get("Name", ""),

                    "Original Email": raw_email,

                    "Full Row": full_row

                })

    total_records = sum(
        len(records)
        for records in recipients.values()
    )

    print("\nEXCEL SUMMARY")

    print(f"Excel files processed: {len(excel_files)}")

    print(f"Total valid records: {total_records}")

    print(f"Unique email addresses: {len(recipients)}")

    return recipients, invalid_rows


# ============================================================
# READ TERMINAL LOG
# ============================================================

def load_sent_recipients(log_file):

    sent_recipients = defaultdict(list)

    print("\n================================")
    print("READING TERMINAL LOG")
    print("================================")

    with open(
        log_file,
        "r",
        encoding="utf-8",
        errors="replace"
    ) as file:

        for line_number, line in enumerate(file, start=1):

            marker_found = None

            for marker in SUCCESS_MARKERS:

                if marker.casefold() in line.casefold():

                    marker_found = marker

                    break

            if not marker_found:
                continue

            # Find the recipient portion after the success marker.
            marker_position = line.casefold().find(
                marker_found.casefold()
            )

            recipient_part = line[
                marker_position + len(marker_found):
            ]

            # Remove sender information after the pipe.
            recipient_part = recipient_part.split("|", 1)[0]

            emails = extract_emails(recipient_part)

            if not emails:
                continue

            # First email after the success marker is the recipient.
            email = emails[0]

            sent_recipients[email].append({

                "Log Line": line_number,

                "Original Log": line.strip()

            })

    print(
        "Unique successfully logged recipients: "
        f"{len(sent_recipients)}"
    )

    return sent_recipients


# ============================================================
# IDENTIFY DUPLICATES
# ============================================================

def identify_duplicates(excel_recipients):

    duplicates = []

    print("\n================================")
    print("IDENTIFYING DUPLICATE ADDRESSES")
    print("================================")

    for email, records in excel_recipients.items():

        if len(records) <= 1:
            continue

        # First occurrence is considered the original.
        original = records[0]

        # Compare every subsequent occurrence against
        # the first occurrence.
        for duplicate in records[1:]:

            comparison = {

                "Duplicate Email": email,

                "Original File": original["Source File"],

                "Original Excel Row": original["Source Row"],

                "Duplicate File": duplicate["Source File"],

                "Duplicate Excel Row": duplicate["Source Row"],

                "Total Email Occurrences": len(records)

            }

            # Add all original Excel columns.
            for column, value in original["Full Row"].items():

                comparison[f"Original - {column}"] = value

            # Add all duplicate Excel columns.
            for column, value in duplicate["Full Row"].items():

                comparison[f"Duplicate - {column}"] = value

            duplicates.append(comparison)

    duplicate_addresses = sum(
        1
        for records in excel_recipients.values()
        if len(records) > 1
    )

    print(f"Duplicate email addresses: {duplicate_addresses}")

    print(f"Duplicate comparisons generated: {len(duplicates)}")

    return duplicates


# ============================================================
# CREATE COMPLETE EXCEL RECORDS
# ============================================================

def create_all_excel_records(excel_recipients):

    all_records = []

    for email, records in excel_recipients.items():

        for record in records:

            output_record = {

                "Email": email,

                "Source File": record["Source File"],

                "Source Row": record["Source Row"],

            }

            # Preserve every original Excel column.
            output_record.update(record["Full Row"])

            all_records.append(output_record)

    return all_records


# ============================================================
# RECONCILIATION
# ============================================================

def reconcile(excel_recipients, sent_recipients):

    successful = []

    not_sent = []

    mismatches = []

    for email, records in excel_recipients.items():

        source_files = ", ".join(
            sorted(set(
                record["Source File"]
                for record in records
            ))
        )

        source_rows = ", ".join(
            f"{record['Source File']}:{record['Source Row']}"
            for record in records
        )

        name = records[0]["Name"]

        if email in sent_recipients:

            successful.append({

                "Email": email,

                "Name": name,

                "Source Files": source_files,

                "Source Rows": source_rows,

                "Occurrences in Excel": len(records),

                "Occurrences in Log": len(sent_recipients[email]),

                "Status": "SUCCESSFULLY LOGGED"

            })

        else:

            not_sent.append({

                "Email": email,

                "Name": name,

                "Source Files": source_files,

                "Source Rows": source_rows,

                "Occurrences in Excel": len(records),

                "Status": "NO SUCCESS RECORD IN LOG"

            })

    # Log addresses that do not exist in Excel.
    for email, records in sent_recipients.items():

        if email not in excel_recipients:

            mismatches.append({

                "Email": email,

                "Log Occurrences": len(records),

                "Log Lines": ", ".join(
                    str(record["Log Line"])
                    for record in records
                ),

                "Status": "NOT FOUND IN EXCEL"

            })

    return successful, not_sent, mismatches


# ============================================================
# GENERATE REPORT
# ============================================================

def generate_report(
    excel_recipients,
    sent_recipients,
    successful,
    not_sent,
    mismatches,
    invalid_rows,
    duplicates,
    all_records
):

    total_excel_records = sum(
        len(records)
        for records in excel_recipients.values()
    )

    duplicate_addresses = sum(
        1
        for records in excel_recipients.values()
        if len(records) > 1
    )

    duplicate_occurrences = sum(
        max(0, len(records) - 1)
        for records in excel_recipients.values()
    )

    summary = pd.DataFrame([

        {
            "Metric": "Total Excel records",
            "Count": total_excel_records
        },

        {
            "Metric": "Unique Excel recipients",
            "Count": len(excel_recipients)
        },

        {
            "Metric": "Unique successfully logged recipients",
            "Count": len(
                set(excel_recipients) & set(sent_recipients)
            )
        },

        {
            "Metric": "Recipients without success record",
            "Count": len(not_sent)
        },

        {
            "Metric": "Log recipients not found in Excel",
            "Count": len(mismatches)
        },

        {
            "Metric": "Invalid Excel rows",
            "Count": len(invalid_rows)
        },

        {
            "Metric": "Duplicate email addresses",
            "Count": duplicate_addresses
        },

        {
            "Metric": "Additional duplicate occurrences",
            "Count": duplicate_occurrences
        },

        {
            "Metric": "Total records retained",
            "Count": len(all_records)
        },

        {
            "Metric": "Repeated successful log entries",
            "Count": sum(
                max(0, len(records) - 1)
                for records in sent_recipients.values()
            )
        }

    ])

    # ========================================================
    # WRITE EXCEL REPORT
    # ========================================================

    print("\nGENERATING EXCEL REPORT...")

    with pd.ExcelWriter(
        OUTPUT_FILE,
        engine="openpyxl"
    ) as writer:

        summary.to_excel(
            writer,
            sheet_name="Summary",
            index=False
        )

        pd.DataFrame(all_records).to_excel(
            writer,
            sheet_name="All Excel Records",
            index=False
        )

        pd.DataFrame(duplicates).to_excel(
            writer,
            sheet_name="Duplicate Comparison",
            index=False
        )

        pd.DataFrame(successful).to_excel(
            writer,
            sheet_name="Successfully Sent",
            index=False
        )

        pd.DataFrame(not_sent).to_excel(
            writer,
            sheet_name="Not Sent",
            index=False
        )

        pd.DataFrame(mismatches).to_excel(
            writer,
            sheet_name="Log Mismatches",
            index=False
        )

        pd.DataFrame(invalid_rows).to_excel(
            writer,
            sheet_name="Invalid Excel Rows",
            index=False
        )

    # ========================================================
    # TERMINAL SUMMARY
    # ========================================================

    print("\n========================================")

    print("       EMAIL RECONCILIATION COMPLETE")

    print("========================================")

    print(f"Total Excel records: {total_excel_records}")

    print(f"Unique Excel recipients: {len(excel_recipients)}")

    print(f"Successfully logged: {len(successful)}")

    print(f"No success record: {len(not_sent)}")

    print(f"Log mismatches: {len(mismatches)}")

    print(f"Invalid Excel rows: {len(invalid_rows)}")

    print("\nDUPLICATE ANALYSIS")

    print(f"Duplicate addresses: {duplicate_addresses}")

    print(f"Additional duplicate occurrences: {duplicate_occurrences}")

    print(f"Side-by-side comparisons: {len(duplicates)}")

    print("\nAll original Excel records have been retained.")

    print("\nReport generated:")

    print(OUTPUT_FILE)

    print("========================================")


# ============================================================
# MAIN
# ============================================================

def main():

    parser = argparse.ArgumentParser(

        description=(
            "Cross-check Excel recipients against terminal "
            "send logs and identify duplicate addresses."
        )

    )

    parser.add_argument(

        "--log",

        required=True,

        help="Path to the terminal log text file."

    )

    args = parser.parse_args()

    log_file = Path(args.log)

    if not log_file.is_file():

        raise FileNotFoundError(log_file)

    if not INPUT_DIR.is_dir():

        raise FileNotFoundError(
            f"Input directory not found: {INPUT_DIR}"
        )

    # Step 1: Read Excel files.
    excel_recipients, invalid_rows = load_excel_recipients()

    # Step 2: Read terminal log.
    sent_recipients = load_sent_recipients(log_file)

    # Step 3: Reconcile Excel and log.
    successful, not_sent, mismatches = reconcile(
        excel_recipients,
        sent_recipients
    )

    # Step 4: Identify duplicate addresses.
    duplicates = identify_duplicates(excel_recipients)

    # Step 5: Preserve all original records.
    all_records = create_all_excel_records(excel_recipients)

    # Step 6: Generate report.
    generate_report(
        excel_recipients,
        sent_recipients,
        successful,
        not_sent,
        mismatches,
        invalid_rows,
        duplicates,
        all_records
    )


if __name__ == "__main__":

    main()
