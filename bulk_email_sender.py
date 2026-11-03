#!/usr/bin/env python3
"""
Campaign-aware AGM email sender.

Features
--------
- Automatically discovers .xlsx recipient files in ./input
- Validates workbook columns and parses multiple addresses per cell
- Deduplicates recipients across all input workbooks
- Sends one individually addressed email per recipient
- Uses persistent campaign-specific job state and daily application caps
- Records uncertain submissions as UNKNOWN rather than silently retrying
- Supports explicit retry of failed or unknown jobs
- Produces Excel failure and summary reports
- Uses an exclusive process lock and atomic JSON state writes
- Defaults to dry-run; sending requires --send and campaign confirmation

Requirements:
    python -m pip install pandas openpyxl

Place the AGM PDF beside this script, or update ATTACHMENT_PATH below.
Configure EMAIL_1/PASSWORD_1 through EMAIL_4/PASSWORD_4 as environment variables.
For Gmail, use an app password where eligible or an approved OAuth solution.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import smtplib
import sys
import time
from contextlib import contextmanager
from datetime import date, datetime
from email.message import EmailMessage
from email.utils import parseaddr
from pathlib import Path
from typing import Any

import pandas as pd


# ---------------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "input"
LOG_DIR = BASE_DIR / "logs"
ATTACHMENT_PATH = BASE_DIR / "ISAMPE AGM Notice 2026.pdf"

SUBJECT = "ISAMPE Election Notice 2026"
CAMPAIGN_ID = "isampe_elec_2026"

# This is an application-level cap only. It does not override provider limits.
SENDER_DAILY_CAP = 450
DELAY_SECONDS = 7
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587
SMTP_TIMEOUT_SECONDS = 30

STATE_FILE = LOG_DIR / "email_progress.json"
FAILURE_FILE = LOG_DIR / "failed_emails.xlsx"
SUMMARY_FILE = LOG_DIR / "email_summary.xlsx"
LOCK_FILE = LOG_DIR / "email_sender.lock"

DEDUPLICATE_ACROSS_FILES = True
REQUIRED_COLUMNS = {"Name", "Email"}
SENDER_ENVIRONMENT_SLOTS = 4


# ---------------------------------------------------------------------------
# GENERAL UTILITIES
# ---------------------------------------------------------------------------

def now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def ensure_directories() -> None:
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    LOG_DIR.mkdir(parents=True, exist_ok=True)


def atomic_write_json(path: Path, data: dict[str, Any]) -> None:
    """Write JSON atomically so a crash cannot leave a half-written state file."""
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_name(path.name + ".tmp")
    with temp_path.open("w", encoding="utf-8") as handle:
        json.dump(data, handle, indent=2, ensure_ascii=False)
        handle.flush()
        os.fsync(handle.fileno())
    os.replace(temp_path, path)


@contextmanager
def single_instance_lock(lock_path: Path):
    """Prevent two copies of this program from sending or modifying state at once."""
    lock_path.parent.mkdir(parents=True, exist_ok=True)
    handle = lock_path.open("a+", encoding="utf-8")
    locked = False

    try:
        if os.name == "nt":
            import msvcrt
            try:
                handle.seek(0)
                # Lock one byte; ensure the file contains a byte to lock.
                if lock_path.stat().st_size == 0:
                    handle.write(" ")
                    handle.flush()
                handle.seek(0)
                msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
                locked = True
            except OSError as exc:
                raise RuntimeError(
                    "Another sender process may already be running. "
                    "If not, inspect the lock file and try again."
                ) from exc
        else:
            import fcntl
            try:
                fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                locked = True
            except OSError as exc:
                raise RuntimeError(
                    "Another sender process may already be running. "
                    "If not, inspect the lock file and try again."
                ) from exc

        yield
    finally:
        if locked:
            try:
                if os.name == "nt":
                    import msvcrt
                    handle.seek(0)
                    msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
                else:
                    import fcntl
                    fcntl.flock(handle.fileno(), fcntl.LOCK_UN)
            except OSError:
                pass
        handle.close()


def safe_string(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value).strip()


# ---------------------------------------------------------------------------
# EMAIL BODY AND ADDRESS PARSING
# ---------------------------------------------------------------------------

def build_body(name: Any, salutation: Any) -> str:
    greetings = {
        "MR": "Mr.",
        "DR": "Dr.",
        "MS": "Ms.",
        "COL": "Col.",
        "GP.CAPT.": "Gp. Capt.",
        "MISS": "Ms.",
        "MRS": "Mrs.",
    }

    name_text = safe_string(name)
    salutation_text = safe_string(salutation).upper()
    title = greetings.get(salutation_text, "")

    if title and name_text:
        greeting = f"Dear {title} {name_text},"
    elif name_text:
        greeting = f"Dear {name_text},"
    else:
        greeting = "Dear Member,"

    return f"""{greeting}

Dear Members,

ELCETION NOTICE

As per the Rules and Regulations of ISAMPE, election is to be held this
year to fill up the vacancies of Treasurer, Joint Secretary-2 and
Director-1 in the Executive Council of the Society for the official years
2026-28. According to the Election Procedure, vide Sec 5.3 of the Rules
and Regulations; the Executive Council shall nominate candidates for
the election to each office. The Executive council in its meeting (online)
held on 23rd October 2026, has made the following nominations:

Treasurer
 Mr. M J Augustin
Scientist -F , Adv. Composites Division,
CSIR-National Aerospace Laboratories
(CSIR-NAL), Kodihalli, Bangalore-560017

Joint Secretary-2 
Dr. R V Lakshmi
Scientist- F, Surface Engg Division,
CSIR-National Aerospace Laboratories,
Kodihalli, Bangalore-560017

Director-1 
Dr. D Saji
Professor of Practice, Dept of Aerospace Engg
IIT(M), Chennai-600036

Members have the right to nominate any other member after asserting
his/her willingness to hold the post, if elected. Such nomination should
be made as per the election procedure defined in Section 5.3 of the
Rules and Regulations and should reach the Secretary latest by 18th
November 2026.
"""


def parse_emails(raw_value: Any) -> list[str]:
    """Extract basic email-address candidates separated by comma/semicolon/space."""
    raw_text = safe_string(raw_value)
    if not raw_text:
        return []

    candidates = re.split(r"[,;\s]+", raw_text)
    results: list[str] = []
    seen: set[str] = set()

    for candidate in candidates:
        candidate = candidate.strip().strip("<>()[]{}\"'")
        if not candidate:
            continue

        _, address = parseaddr(candidate)
        if not address:
            address = candidate

        # Deliberately basic syntax validation; it cannot prove mailbox existence.
        if not re.fullmatch(r"[^@\s<>]+@[^@\s<>]+\.[^@\s<>]+", address):
            continue

        normalized = address.casefold()
        if normalized not in seen:
            seen.add(normalized)
            results.append(address)

    return results


# ---------------------------------------------------------------------------
# SENDER ACCOUNTS
# ---------------------------------------------------------------------------

def load_senders() -> list[dict[str, Any]]:
    senders: list[dict[str, Any]] = []

    for slot in range(1, SENDER_ENVIRONMENT_SLOTS + 1):
        email_address = os.getenv(f"EMAIL_{slot}", "").strip()
        password = os.getenv(f"PASSWORD_{slot}", "").strip()

        if not email_address and not password:
            continue
        if not email_address or not password:
            raise RuntimeError(
                f"EMAIL_{slot} and PASSWORD_{slot} must both be configured."
            )

        if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", email_address):
            raise RuntimeError(f"EMAIL_{slot} does not look like a valid address.")

        senders.append({
            "id": slot,
            "email": email_address,
            "password": password,
        })

    # Avoid counting the same sender account twice.
    unique: dict[str, dict[str, Any]] = {}
    for sender in senders:
        key = sender["email"].casefold()
        if key in unique:
            raise RuntimeError(
                f"Sender account {sender['email']} is configured more than once."
            )
        unique[key] = sender

    if not unique:
        raise RuntimeError(
            "No sender accounts configured. Set EMAIL_1/PASSWORD_1, etc. "
            "in your environment before running with --send."
        )

    return list(unique.values())


def connect_sender(sender: dict[str, Any]) -> smtplib.SMTP:
    server = smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=SMTP_TIMEOUT_SECONDS)
    try:
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(sender["email"], sender["password"])
        setattr(server, "_campaign_sender_email", sender["email"].casefold())
        return server
    except Exception:
        try:
            server.quit()
        except Exception:
            server.close()
        raise


def close_connection(server: smtplib.SMTP | None) -> None:
    if server is None:
        return
    try:
        server.quit()
    except Exception:
        try:
            server.close()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# STATE
# ---------------------------------------------------------------------------

def fresh_state() -> dict[str, Any]:
    return {
        "schema_version": 2,
        "daily_counts": {},  # date -> normalized sender email -> count
        "campaigns": {},     # campaign id -> jobs and failure history
    }


def load_state() -> dict[str, Any]:
    if not STATE_FILE.exists():
        return fresh_state()

    try:
        with STATE_FILE.open("r", encoding="utf-8") as handle:
            state = json.load(handle)
    except (OSError, json.JSONDecodeError) as exc:
        backup = STATE_FILE.with_name(
            f"{STATE_FILE.stem}.corrupt-{datetime.now().strftime('%Y%m%d-%H%M%S')}.json"
        )
        try:
            os.replace(STATE_FILE, backup)
        except OSError:
            pass
        raise RuntimeError(
            f"Progress file is unreadable. It was preserved as {backup.name} "
            "if possible. Do not send until the state has been reviewed."
        ) from exc

    if not isinstance(state, dict):
        raise RuntimeError("Progress file has an invalid structure.")

    # Migrate the original script's state format safely. Completed recipients are
    # carried forward as SMTP_ACCEPTED; old failures become UNKNOWN so they are
    # never blindly resent without review.
    if "campaigns" not in state and (
        "completed" in state or "failures" in state or "daily_counts" in state
    ):
        legacy = state
        migrated = fresh_state()
        today = date.today().isoformat()

        if legacy.get("date") == today and isinstance(legacy.get("daily_counts"), dict):
            migrated["daily_counts"][today] = {
                str(email).casefold(): max(0, int(count))
                for email, count in legacy["daily_counts"].items()
                if str(email).strip()
            }

        campaign = {
            "created_at": now_iso(),
            "updated_at": now_iso(),
            "jobs": {},
            "failure_history": [],
            "migration_note": "Migrated from the previous email sender state format.",
        }

        for item in legacy.get("completed", []):
            if not isinstance(item, dict):
                continue
            recipient = safe_string(item.get("recipient"))
            if not recipient or "@" not in recipient:
                continue
            normalized = recipient.casefold()
            key = f"{CAMPAIGN_ID}|{normalized}"
            campaign["jobs"][key] = {
                "job_key": key,
                "campaign_id": CAMPAIGN_ID,
                "recipient": recipient,
                "source_file": item.get("source_file", ""),
                "source_row": item.get("source_row", ""),
                "sender": item.get("sender", ""),
                "status": "SMTP_ACCEPTED",
                "accepted_at": item.get("sent_at", ""),
                "updated_at": now_iso(),
                "migration_note": "Imported from legacy completed record.",
            }

        for item in legacy.get("failures", []):
            if not isinstance(item, dict):
                continue
            recipient = safe_string(item.get("recipient"))
            old_error = safe_string(item.get("error"))
            key = f"{CAMPAIGN_ID}|{recipient.casefold()}" if recipient else ""
            record = {
                "job_key": key,
                "campaign_id": CAMPAIGN_ID,
                "recipient": recipient,
                "source_file": item.get("source_file", ""),
                "source_row": item.get("source_row", ""),
                "sender": item.get("sender", ""),
                "status": "UNKNOWN" if recipient and "@" in recipient else "INVALID_INPUT",
                "error": (
                    "Migrated legacy failure; review whether the message was accepted "
                    "before retrying. Original error: " + old_error
                ),
                "recorded_at": item.get("failure_time", now_iso()),
                "migration_note": "Imported from legacy failure record.",
            }
            campaign["failure_history"].append(record.copy())
            if key and key not in campaign["jobs"]:
                campaign["jobs"][key] = record

        migrated["campaigns"][CAMPAIGN_ID] = campaign
        state = migrated
        print(
            "Notice: migrated the previous progress-file format. "
            "Legacy failures are marked UNKNOWN and will not be resent automatically."
        )

    state.setdefault("schema_version", 2)
    state.setdefault("daily_counts", {})
    state.setdefault("campaigns", {})

    if not isinstance(state["daily_counts"], dict):
        raise RuntimeError("Progress file daily_counts field is invalid.")
    if not isinstance(state["campaigns"], dict):
        raise RuntimeError("Progress file campaigns field is invalid.")

    return state


def save_state(state: dict[str, Any]) -> None:
    atomic_write_json(STATE_FILE, state)


def get_campaign(state: dict[str, Any], campaign_id: str) -> dict[str, Any]:
    campaigns = state["campaigns"]
    campaign = campaigns.setdefault(
        campaign_id,
        {
            "created_at": now_iso(),
            "updated_at": now_iso(),
            "jobs": {},
            "failure_history": [],
        },
    )
    campaign.setdefault("jobs", {})
    campaign.setdefault("failure_history", [])
    return campaign


def daily_count(state: dict[str, Any], sender_email: str, day: str | None = None) -> int:
    day = day or date.today().isoformat()
    day_counts = state["daily_counts"].setdefault(day, {})
    try:
        return int(day_counts.get(sender_email.casefold(), 0))
    except (TypeError, ValueError):
        return 0


def increment_daily_count(state: dict[str, Any], sender_email: str) -> None:
    day = date.today().isoformat()
    day_counts = state["daily_counts"].setdefault(day, {})
    key = sender_email.casefold()
    day_counts[key] = daily_count(state, key, day) + 1


# ---------------------------------------------------------------------------
# INPUT DISCOVERY AND VALIDATION
# ---------------------------------------------------------------------------

def discover_input_files() -> list[Path]:
    files = sorted(
        path for path in INPUT_DIR.glob("*.xlsx")
        if path.is_file() and not path.name.startswith("~$")
    )
    # Only the input directory is scanned; reports and attachments belong elsewhere.
    return files


def load_jobs(input_files: list[Path]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    jobs: list[dict[str, Any]] = []
    invalid_rows: list[dict[str, Any]] = []

    for path in input_files:
        try:
            frame = pd.read_excel(path, dtype=object)
        except Exception as exc:
            raise RuntimeError(f"Could not read workbook {path.name}: {exc}") from exc

        missing = REQUIRED_COLUMNS - set(frame.columns)
        if missing:
            raise ValueError(
                f"{path.name} is missing required column(s): {', '.join(sorted(missing))}. "
                f"Required columns are: {', '.join(sorted(REQUIRED_COLUMNS))}."
            )

        for index, row in frame.iterrows():
            original_row = {
                str(column): safe_string(value)
                for column, value in row.to_dict().items()
            }
            raw_email = row.get("Email", "")
            recipients = parse_emails(raw_email)

            if not recipients:
                invalid_rows.append({
                    "source_file": path.name,
                    "source_row": int(index) + 2,
                    "recipient": safe_string(raw_email),
                    "name": safe_string(row.get("Name", "")),
                    "salutation": safe_string(row.get("Salutation", "")),
                    "error": "No valid email address found in Email column.",
                    "original_row": original_row,
                    "status": "INVALID_INPUT",
                    "recorded_at": now_iso(),
                })
                continue

            for recipient in recipients:
                jobs.append({
                    "campaign_id": CAMPAIGN_ID,
                    "recipient": recipient,
                    "recipient_key": recipient.casefold(),
                    "name": safe_string(row.get("Name", "")),
                    "salutation": safe_string(row.get("Salutation", "")),
                    "source_file": path.name,
                    "source_row": int(index) + 2,
                    "original_row": original_row,
                })

    if DEDUPLICATE_ACROSS_FILES:
        # First occurrence wins, in sorted filename/row order.
        unique_jobs: list[dict[str, Any]] = []
        seen: set[str] = set()
        for job in jobs:
            key = job["recipient_key"]
            if key in seen:
                continue
            seen.add(key)
            unique_jobs.append(job)
        jobs = unique_jobs

    return jobs, invalid_rows


def job_key_for(job: dict[str, Any], campaign_id: str) -> str:
    # A campaign has at most one job per normalized recipient.
    return f"{campaign_id}|{job['recipient_key']}"


# ---------------------------------------------------------------------------
# EMAIL SUBMISSION AND FAILURE CLASSIFICATION
# ---------------------------------------------------------------------------

class SubmissionOutcomeUnknown(RuntimeError):
    """Raised when the client cannot determine whether SMTP accepted the message."""


class ProviderLimitError(RuntimeError):
    """Raised for provider responses indicating rate limiting or sending restrictions."""


def send_one_email(
    server: smtplib.SMTP,
    sender: dict[str, Any],
    job: dict[str, Any],
    attachment_path: Path,
) -> None:
    message = EmailMessage()
    message["From"] = sender["email"]
    message["To"] = job["recipient"]
    message["Subject"] = SUBJECT
    message.set_content(build_body(job.get("name"), job.get("salutation")))

    with attachment_path.open("rb") as handle:
        message.add_attachment(
            handle.read(),
            maintype="application",
            subtype="pdf",
            filename=attachment_path.name,
        )

    try:
        refused = server.send_message(
            message,
            from_addr=sender["email"],
            to_addrs=[job["recipient"]],
        )
    except smtplib.SMTPResponseException as exc:
        code = int(exc.smtp_code)
        response = str(exc.smtp_error).lower()
        limit_words = (
            "quota", "rate", "limit", "too many", "daily",
            "temporarily blocked", "sending limit", "try again later",
        )
        if code in (421, 450, 451, 452, 454) or any(
            word in response for word in limit_words
        ):
            raise ProviderLimitError(
                f"Provider returned SMTP {code}: {exc.smtp_error!r}"
            ) from exc
        # A server response explicitly rejected the operation; record as failure.
        raise

    except (smtplib.SMTPServerDisconnected, TimeoutError, OSError) as exc:
        # The connection may have failed after the server accepted DATA.
        raise SubmissionOutcomeUnknown(
            f"Connection outcome is uncertain: {type(exc).__name__}: {exc}"
        ) from exc

    except smtplib.SMTPException as exc:
        # Conservative handling: do not automatically resend if acceptance is unclear.
        raise SubmissionOutcomeUnknown(
            f"SMTP outcome is uncertain: {type(exc).__name__}: {exc}"
        ) from exc

    if refused:
        raise smtplib.SMTPRecipientsRefused(refused)


def record_failure(
    state: dict[str, Any],
    campaign: dict[str, Any],
    job: dict[str, Any],
    status: str,
    error: str,
    sender_email: str = "",
) -> None:
    record = {
        "job_key": job_key_for(job, CAMPAIGN_ID),
        "campaign_id": CAMPAIGN_ID,
        "recipient": job["recipient"],
        "source_file": job["source_file"],
        "source_row": job["source_row"],
        "sender": sender_email,
        "status": status,
        "error": error,
        "recorded_at": now_iso(),
        "original_row": job.get("original_row", {}),
    }
    campaign["jobs"][record["job_key"]] = record
    campaign["failure_history"].append(record.copy())
    campaign["updated_at"] = now_iso()
    save_state(state)


def export_failure_report(
    campaign: dict[str, Any],
    invalid_rows: list[dict[str, Any]],
) -> None:
    rows: list[dict[str, Any]] = []

    for item in campaign.get("failure_history", []):
        original = item.get("original_row", {})
        rows.append({
            **original,
            "Campaign ID": item.get("campaign_id", CAMPAIGN_ID),
            "Source File": item.get("source_file", ""),
            "Source Row": item.get("source_row", ""),
            "Recipient": item.get("recipient", ""),
            "Sender": item.get("sender", ""),
            "Status": item.get("status", ""),
            "Error": item.get("error", ""),
            "Recorded At": item.get("recorded_at", ""),
        })

    for item in invalid_rows:
        rows.append({
            **item.get("original_row", {}),
            "Campaign ID": CAMPAIGN_ID,
            "Source File": item.get("source_file", ""),
            "Source Row": item.get("source_row", ""),
            "Recipient": item.get("recipient", ""),
            "Sender": "",
            "Status": item.get("status", "INVALID_INPUT"),
            "Error": item.get("error", ""),
            "Recorded At": item.get("recorded_at", ""),
        })

    columns = [
        "Campaign ID", "Source File", "Source Row", "Recipient", "Sender",
        "Status", "Error", "Recorded At",
    ]
    frame = pd.DataFrame(rows)
    if not frame.empty:
        ordered = columns + [col for col in frame.columns if col not in columns]
        frame = frame.reindex(columns=ordered)
    else:
        frame = pd.DataFrame(columns=columns)

    temp_path = FAILURE_FILE.with_name(FAILURE_FILE.stem + ".tmp.xlsx")
    frame.to_excel(temp_path, index=False)
    os.replace(temp_path, FAILURE_FILE)


def export_summary(
    state: dict[str, Any],
    campaign: dict[str, Any],
    senders: list[dict[str, Any]],
    input_count: int,
    unique_count: int,
    invalid_count: int,
) -> None:
    jobs = campaign.get("jobs", {})
    status_counts: dict[str, int] = {}
    for item in jobs.values():
        status = str(item.get("status", "UNKNOWN"))
        status_counts[status] = status_counts.get(status, 0) + 1

    summary_rows = [
        {"Metric": "Campaign ID", "Value": CAMPAIGN_ID},
        {"Metric": "Current date", "Value": date.today().isoformat()},
        {"Metric": "Input workbook count", "Value": input_count},
        {"Metric": "Unique valid recipients in current inputs", "Value": unique_count},
        {"Metric": "Invalid input rows", "Value": invalid_count},
        {"Metric": "Jobs in campaign state", "Value": len(jobs)},
    ]
    for status, count in sorted(status_counts.items()):
        summary_rows.append({"Metric": f"Jobs with status {status}", "Value": count})

    sender_rows = []
    for sender in senders:
        count = daily_count(state, sender["email"])
        sender_rows.append({
            "Sender": sender["email"],
            "Successful SMTP acceptances recorded today": count,
            "Configured application cap": SENDER_DAILY_CAP,
            "Remaining application capacity": max(0, SENDER_DAILY_CAP - count),
        })

    with pd.ExcelWriter(SUMMARY_FILE, engine="openpyxl") as writer:
        pd.DataFrame(summary_rows).to_excel(writer, sheet_name="Summary", index=False)
        pd.DataFrame(sender_rows).to_excel(writer, sheet_name="Senders", index=False)

        job_rows = []
        for key, item in jobs.items():
            job_rows.append({
                "Job Key": key,
                "Recipient": item.get("recipient", ""),
                "Status": item.get("status", ""),
                "Sender": item.get("sender", ""),
                "Updated At": item.get("updated_at", item.get("recorded_at", "")),
                "Error": item.get("error", ""),
            })
        pd.DataFrame(job_rows).to_excel(writer, sheet_name="Campaign Jobs", index=False)


# ---------------------------------------------------------------------------
# CAMPAIGN RUNNER
# ---------------------------------------------------------------------------

def choose_sender(
    senders: list[dict[str, Any]],
    state: dict[str, Any],
    start_index: int,
) -> tuple[dict[str, Any] | None, int]:
    for offset in range(len(senders)):
        index = (start_index + offset) % len(senders)
        sender = senders[index]
        if daily_count(state, sender["email"]) < SENDER_DAILY_CAP:
            return sender, index
    return None, start_index


def run_campaign(args: argparse.Namespace) -> int:
    ensure_directories()

    if not ATTACHMENT_PATH.is_file():
        raise FileNotFoundError(
            f"Attachment not found: {ATTACHMENT_PATH}\n"
            "Place the AGM PDF beside this script or update ATTACHMENT_PATH."
        )

    input_files = discover_input_files()
    if not input_files:
        print(f"No .xlsx files found in {INPUT_DIR}")
        print("Place recipient workbooks in the input folder and run again.")
        return 1

    jobs, invalid_rows = load_jobs(input_files)
    if not jobs and not invalid_rows:
        print("No recipient rows found in the input workbooks.")
        return 1

    print(f"Campaign: {CAMPAIGN_ID}")
    print(f"Input folder: {INPUT_DIR}")
    print(f"Input workbooks: {len(input_files)}")
    for path in input_files:
        print(f"  - {path.name}")
    print(f"Unique valid recipients: {len(jobs)}")
    print(f"Rows without a valid email address: {len(invalid_rows)}")
    print(f"Attachment: {ATTACHMENT_PATH.name}")
    print(f"Application cap per sender/day: {SENDER_DAILY_CAP}")
    print(f"Mode: {'SEND' if args.send else 'DRY RUN'}")

    if not args.send:
        print("\nDry run only. No email has been sent.")
        print("Inspect the workbook list and recipient counts.")
        print("To send, run with --send.")
        return 0

    senders = load_senders()
    state = load_state()
    campaign = get_campaign(state, CAMPAIGN_ID)

    # A process may have stopped after persisting SENDING but before recording
    # the result. Such a job must not be automatically resubmitted.
    for key, item in list(campaign["jobs"].items()):
        if item.get("status") == "SENDING":
            item["status"] = "UNKNOWN"
            item["error"] = (
                "Process stopped while this job was marked SENDING. "
                "SMTP acceptance cannot be determined automatically; review before retrying."
            )
            item["updated_at"] = now_iso()
            campaign["failure_history"].append(item.copy())
    save_state(state)

    # Validate/reconcile existing jobs and decide which statuses are eligible.
    sendable: list[dict[str, Any]] = []
    skipped = 0
    for job in jobs:
        key = job_key_for(job, CAMPAIGN_ID)
        prior = campaign["jobs"].get(key)
        if prior:
            status = prior.get("status")
            if status == "SMTP_ACCEPTED":
                skipped += 1
                continue
            if status == "UNKNOWN" and not args.retry_unknown:
                skipped += 1
                continue
            if status in {"FAILED_RETRYABLE", "FAILED_PERMANENT", "INVALID_INPUT"}:
                if not args.retry_failed:
                    skipped += 1
                    continue
            elif status not in {"UNKNOWN", "FAILED_RETRYABLE", "FAILED_PERMANENT"}:
                skipped += 1
                continue
        sendable.append(job)

    print(f"Previously accepted/skipped or held for review: {skipped}")
    print(f"Jobs eligible in this run: {len(sendable)}")

    if not sendable:
        print("No eligible jobs to send.")
        export_failure_report(campaign, invalid_rows)
        export_summary(state, campaign, senders, len(input_files), len(jobs), len(invalid_rows))
        return 0

    if not args.yes:
        confirmation = input(
            f"\nType the campaign ID '{CAMPAIGN_ID}' to confirm sending: "
        ).strip()
        if confirmation != CAMPAIGN_ID:
            print("Confirmation did not match. Nothing was sent.")
            return 1

    server: smtplib.SMTP | None = None
    active_sender_email: str | None = None
    sender_start_index = 0
    sent_this_run = 0
    failed_this_run = 0
    unknown_this_run = 0
    stop_reason = ""

    try:
        for job in sendable:
            key = job_key_for(job, CAMPAIGN_ID)
            selected_sender, selected_index = choose_sender(
                senders, state, sender_start_index
            )

            if selected_sender is None:
                stop_reason = (
                    "All configured senders reached the application-level daily cap."
                )
                print(f"\n[STOP] {stop_reason}")
                break

            sender_start_index = (selected_index + 1) % len(senders)
            sender_email = selected_sender["email"]
            sender_key = sender_email.casefold()

            # Persist a pre-submission marker. If the process crashes, startup
            # changes SENDING to UNKNOWN and will not blindly retry it.
            campaign["jobs"][key] = {
                "job_key": key,
                "campaign_id": CAMPAIGN_ID,
                "recipient": job["recipient"],
                "source_file": job["source_file"],
                "source_row": job["source_row"],
                "sender": sender_email,
                "status": "SENDING",
                "started_at": now_iso(),
                "updated_at": now_iso(),
                "original_row": job.get("original_row", {}),
            }
            campaign["updated_at"] = now_iso()
            save_state(state)

            try:
                if server is None or active_sender_email != sender_key:
                    close_connection(server)
                    server = connect_sender(selected_sender)
                    active_sender_email = sender_key

                send_one_email(server, selected_sender, job, ATTACHMENT_PATH)

                accepted = {
                    "job_key": key,
                    "campaign_id": CAMPAIGN_ID,
                    "recipient": job["recipient"],
                    "source_file": job["source_file"],
                    "source_row": job["source_row"],
                    "sender": sender_email,
                    "status": "SMTP_ACCEPTED",
                    "accepted_at": now_iso(),
                    "updated_at": now_iso(),
                    "original_row": job.get("original_row", {}),
                }
                campaign["jobs"][key] = accepted
                increment_daily_count(state, sender_email)
                campaign["updated_at"] = now_iso()
                save_state(state)

                sent_this_run += 1
                print(
                    f"[SMTP ACCEPTED] {job['recipient']} | "
                    f"{sender_email} | "
                    f"{daily_count(state, sender_email)}/{SENDER_DAILY_CAP}"
                )
                time.sleep(max(0, DELAY_SECONDS))

            except ProviderLimitError as exc:
                # Do not rotate to another account to work around a provider limit.
                message = str(exc)
                campaign["jobs"][key] = {
                    "job_key": key,
                    "campaign_id": CAMPAIGN_ID,
                    "recipient": job["recipient"],
                    "source_file": job["source_file"],
                    "source_row": job["source_row"],
                    "sender": sender_email,
                    "status": "UNKNOWN",
                    "error": message + " Sending stopped; inspect the provider response.",
                    "updated_at": now_iso(),
                    "original_row": job.get("original_row", {}),
                }
                campaign["failure_history"].append(campaign["jobs"][key].copy())
                campaign["updated_at"] = now_iso()
                save_state(state)
                unknown_this_run += 1
                stop_reason = (
                    "Provider rate/quota restriction received. Stopping the whole campaign; "
                    "not switching accounts to bypass the restriction."
                )
                print(f"\n[STOP] {stop_reason}\n{message}")
                break

            except SubmissionOutcomeUnknown as exc:
                message = str(exc)
                record_failure(
                    state, campaign, job, "UNKNOWN", message, sender_email
                )
                unknown_this_run += 1
                print(f"[UNKNOWN] {job['recipient']}: {message}")
                close_connection(server)
                server = None
                active_sender_email = None
                # Continue only to other distinct recipients; this job will not
                # be auto-retried on the next run.
                continue

            except smtplib.SMTPAuthenticationError as exc:
                message = (
                    f"SMTP authentication failed for {sender_email}: "
                    f"{exc.smtp_code} {exc.smtp_error!r}"
                )
                record_failure(
                    state, campaign, job, "FAILED_PERMANENT", message, sender_email
                )
                failed_this_run += 1
                print(f"[AUTH ERROR] {message}")
                close_connection(server)
                server = None
                active_sender_email = None
                # Stop rather than rapidly cycling through accounts after auth failure.
                stop_reason = "Authentication failed; correct credentials and review state."
                break

            except smtplib.SMTPResponseException as exc:
                code = int(exc.smtp_code)
                message = f"SMTP {code}: {exc.smtp_error!r}"
                status = "FAILED_PERMANENT" if 500 <= code < 600 else "FAILED_RETRYABLE"
                record_failure(
                    state, campaign, job, status, message, sender_email
                )
                failed_this_run += 1
                print(f"[{status}] {job['recipient']}: {message}")
                close_connection(server)
                server = None
                active_sender_email = None

            except smtplib.SMTPRecipientsRefused as exc:
                message = f"Recipient refused: {exc}"
                record_failure(
                    state, campaign, job, "FAILED_PERMANENT", message, sender_email
                )
                failed_this_run += 1
                print(f"[FAILED] {job['recipient']}: {message}")
                close_connection(server)
                server = None
                active_sender_email = None

            except (OSError, smtplib.SMTPException) as exc:
                # Conservative default for transport errors: acceptance may be unclear.
                message = f"{type(exc).__name__}: {exc}"
                record_failure(
                    state, campaign, job, "UNKNOWN", message, sender_email
                )
                unknown_this_run += 1
                print(f"[UNKNOWN] {job['recipient']}: {message}")
                close_connection(server)
                server = None
                active_sender_email = None

            except Exception as exc:
                # Exceptions before SMTP submission can still be retried, but the
                # code cannot reliably infer the exact stage for every exception.
                message = f"{type(exc).__name__}: {exc}"
                record_failure(
                    state, campaign, job, "FAILED_RETRYABLE", message, sender_email
                )
                failed_this_run += 1
                print(f"[FAILED] {job['recipient']}: {message}")
                close_connection(server)
                server = None
                active_sender_email = None

    except KeyboardInterrupt:
        print("\nInterrupted by user. Saving campaign state...")
        save_state(state)

    finally:
        close_connection(server)
        save_state(state)
        try:
            export_failure_report(campaign, invalid_rows)
            export_summary(
                state, campaign, senders, len(input_files), len(jobs), len(invalid_rows)
            )
        except Exception as exc:
            print(f"[REPORT ERROR] Could not write report: {exc}", file=sys.stderr)

    print("\n================ CAMPAIGN SUMMARY ================")
    print(f"Campaign ID: {CAMPAIGN_ID}")
    print(f"SMTP acceptances recorded this run: {sent_this_run}")
    print(f"Known failures this run: {failed_this_run}")
    print(f"Unknown outcomes this run: {unknown_this_run}")
    print(f"Previously skipped/held jobs: {skipped}")
    if stop_reason:
        print(f"Stop reason: {stop_reason}")
    print(f"Failure report: {FAILURE_FILE}")
    print(f"Summary report: {SUMMARY_FILE}")
    print(f"Persistent state: {STATE_FILE}")
    print("Note: SMTP acceptance does not guarantee delivery to the recipient's inbox.")
    return 0


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Send individually addressed AGM notices from Excel files."
    )
    parser.add_argument(
        "--send",
        action="store_true",
        help="Actually submit emails. Without this flag, the program performs a dry run.",
    )
    parser.add_argument(
        "--yes",
        action="store_true",
        help="Skip interactive campaign confirmation (use only in controlled automation).",
    )
    parser.add_argument(
        "--retry-failed",
        action="store_true",
        help="Retry jobs previously marked FAILED_RETRYABLE or FAILED_PERMANENT.",
    )
    parser.add_argument(
        "--retry-unknown",
        action="store_true",
        help=(
            "Allow resubmission of jobs marked UNKNOWN. Only use after manually checking "
            "whether those messages were accepted; duplicates may result."
        ),
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    ensure_directories()

    try:
        with single_instance_lock(LOCK_FILE):
            return run_campaign(args)
    except Exception as exc:
        print(f"[FATAL] {type(exc).__name__}: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
