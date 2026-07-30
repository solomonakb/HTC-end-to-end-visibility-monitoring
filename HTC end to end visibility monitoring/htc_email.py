"""
htc_email.py  –  HTC End-to-End Visibility Monitoring subscription mailer
==========================================================================
Sends Alert B (XXX Serial Number) and Alert C (Missing Mandatory Component)
digest emails to subscribers, via the Ethiopian Airlines Exchange / OWA
server, using the same exchangelib + NTLM transport pattern already used
by Email_server.py in the WTT portal.

Sender mailbox : htcvisibilitymonitoring@ethiopianairlines.com
Credential resolution order:
    1. DB app_config table  →  key = 'HTC_OWA_PASSWORD'   (set via ⚙ admin panel)
    2. Environment variable →  HTC_OWA_PASSWORD
    3. Hard-coded legacy default (logs a warning — replace immediately)

No SMTP or win32com dependencies.
"""

import logging
import os
import time
from datetime import datetime
from html import escape

from exchangelib import (
    Account,
    Configuration,
    Credentials,
    HTMLBody,
    Mailbox,
    Message,
    DELEGATE,
    NTLM,
)
from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter

import database

logger = logging.getLogger(__name__)

# ── Disable SSL verification once at import time (internal Exchange) ──────────
BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter

# ── Exchange constants ────────────────────────────────────────────────────────
_OWA_SERVER = "etlbha.ethiopianairlines.com"
_SENDER     = "htcvisibilitymonitoring@ethiopianairlines.com"

_CONFIG_KEY_PASSWORD = "HTC_OWA_PASSWORD"

# NTLM auth wants DOMAIN\sAMAccountName, not the email/UPN form — OWA webmail
# tolerates the email form but raw NTLM handshakes (as exchangelib does them)
# frequently reject it.
_NTLM_USERNAME = r"et\htcvisibilitymonitor"

# ── Retry configuration ───────────────────────────────────────────────────────
_MAX_RETRIES   = 2
_RETRY_DELAY_S = 5

ALERT_LABELS = {
    "ALERT_B": "Alert B — HTC XXX Serial Number Monitoring",
    "ALERT_C": "Alert C — Missing Mandatory Component (MMC)",
}


# ─────────────────────────────────────────────────────────────────────────────
# Credential resolution / account builder
# ─────────────────────────────────────────────────────────────────────────────

def _get_owa_password(db_path: str) -> str:
    """
    Resolve the OWA password for htcvisibilitymonitoring@ethiopianairlines.com.

    Priority (highest → lowest):
      1. DB app_config table  (key = 'HTC_OWA_PASSWORD')  ← set via ⚙ admin panel
      2. HTC_OWA_PASSWORD environment variable
      3. Hard-coded legacy default                          ← triggers a warning
    """
    try:
        val = database.get_config_value(db_path, _CONFIG_KEY_PASSWORD)
        if val:
            logger.info("HTC OWA password source: database (app_config table).")
            return val
    except Exception as exc:
        logger.warning(
            "Could not read %s from DB — falling back to env/default: %s",
            _CONFIG_KEY_PASSWORD, exc
        )

    env_pw = os.environ.get(_CONFIG_KEY_PASSWORD, "")
    if env_pw:
        logger.info("HTC OWA password source: %s environment variable.", _CONFIG_KEY_PASSWORD)
        return env_pw

    logger.warning(
        "HTC OWA password source: LEGACY HARD-CODED DEFAULT. "
        "Set a real password via the ⚙ admin settings panel immediately."
    )
    return "ChangeMe@2026"


def _build_exchange_account(db_path: str) -> Account:
    password = _get_owa_password(db_path)
    creds    = Credentials(username=_NTLM_USERNAME, password=password)
    config   = Configuration(server=_OWA_SERVER, credentials=creds, auth_type=NTLM)
    try:
        account = Account(
            primary_smtp_address=_SENDER,
            config=config,
            autodiscover=False,
            access_type=DELEGATE,
        )
        return account
    except Exception as exc:
        raise RuntimeError(
            f"Exchange authentication failed for {_SENDER} on {_OWA_SERVER}: {exc}"
        ) from exc


def test_owa_connection(db_path: str):
    """Attempt to authenticate only (no email sent). Returns (ok: bool, message: str)."""
    try:
        _build_exchange_account(db_path)
        return True, f"Successfully authenticated as {_SENDER}."
    except Exception as exc:
        return False, str(exc)


# ─────────────────────────────────────────────────────────────────────────────
# HTML body builders
# ─────────────────────────────────────────────────────────────────────────────

def _stat_box(label: str, value, colour: str) -> str:
    return (
        f'<td width="25%" align="center" '
        f'style="padding:14px 8px;background-color:{colour};'
        f'border-radius:6px;color:#ffffff;">'
        f'<div style="font-family:\'Segoe UI\',Arial,sans-serif;'
        f'font-size:26px;font-weight:700;line-height:1;">'
        f'<span style="color:#ffffff !important;">{value}</span>'
        f'</div>'
        f'<div style="font-family:\'Segoe UI\',Arial,sans-serif;'
        f'font-size:11px;margin-top:4px;font-weight:600;">'
        f'<span style="color:#ffffff !important;">{label}</span>'
        f'</div>'
        f'</td>'
    )


def _table_header(cols):
    ths = "".join(
        f'<th style="padding:8px 10px;background:#1a56db;color:#ffffff;'
        f'font-family:Arial,sans-serif;font-size:11px;text-align:left;'
        f'white-space:nowrap;">{escape(c)}</th>'
        for c in cols
    )
    return f"<tr>{ths}</tr>"


def _td(val, bold_red=False):
    style = (
        "padding:6px 10px;font-family:Arial,sans-serif;font-size:12px;"
        "border-bottom:1px solid #e5e7eb;white-space:nowrap;"
    )
    if bold_red:
        style += "color:#b91c1c;font-weight:700;"
    else:
        style += "color:#1f2937;"
    return f'<td style="{style}">{escape(str(val) if val is not None else "")}</td>'


def _alert_b_rows(rows):
    trs = []
    stripe = ["#ffffff", "#f4f7ff"]
    for i, r in enumerate(rows):
        bg = stripe[i % 2]
        trs.append(
            f'<tr style="background:{bg};">'
            + _td(r.get("event_dt"))
            + _td(r.get("aircraft"))
            + _td(r.get("config_slot_code"))
            + _td(r.get("config_slot"))
            + _td(r.get("part_no"))
            + _td(r.get("part_desc"))
            + _td(r.get("serial_number"), bold_red=True)
            + _td(r.get("barcode"))
            + _td(r.get("performed_by_user"))
            + '</tr>'
        )
    return "\n".join(trs)


def _alert_c_rows(rows):
    trs = []
    stripe = ["#ffffff", "#f4f7ff"]
    for i, r in enumerate(rows):
        bg = stripe[i % 2]
        severity = r.get("mmc_severity", "")
        days = r.get("days_since_removal", "")
        trs.append(
            f'<tr style="background:{bg};">'
            + _td(r.get("event_dt"))
            + _td(r.get("aircraft"))
            + _td(r.get("config_slot_code"))
            + _td(r.get("part_no"))
            + _td(r.get("part_group_name"))
            + _td(r.get("barcode"))
            + _td(r.get("performed_by_user"))
            + _td(days, bold_red=(severity == "CRITICAL"))
            + '</tr>'
        )
    return "\n".join(trs)


_ALERT_B_COLUMNS = [
    "Event Date", "Aircraft", "Config Slot Code", "Config Slot",
    "Part No", "Part Desc", "Serial Number", "Barcode", "Performed By",
]
_ALERT_C_COLUMNS = [
    "Removal Date", "Aircraft", "Config Slot Code", "Part No",
    "Part Group", "Barcode", "Performed By", "Days Overdue",
]


def build_digest_html(fleet_type: str, alert_type: str, rows: list) -> str:
    """Build the subscription digest HTML: summary stat row + full data table."""
    date_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    alert_label = ALERT_LABELS.get(alert_type, alert_type)
    total = len(rows)

    if alert_type == "ALERT_B":
        critical_label = "XXX S/N Alerts"
        critical_count = total
        table_cols = _ALERT_B_COLUMNS
        table_rows_html = _alert_b_rows(rows)
        banner_colour = "#FF4500"
    else:  # ALERT_C
        critical_count = sum(1 for r in rows if r.get("mmc_severity") == "CRITICAL")
        warning_count = total - critical_count
        critical_label = "Critical (Overdue)"
        table_cols = _ALERT_C_COLUMNS
        table_rows_html = _alert_c_rows(rows)
        banner_colour = "#FF0000"

    if alert_type == "ALERT_B":
        stats_html = (
            '<table width="100%" cellpadding="0" cellspacing="0"><tr>'
            + _stat_box("Total Events", total, "#1a56db")
            + _stat_box(critical_label, critical_count, "#b91c1c")
            + _stat_box("Fleet", escape(fleet_type), "#374151")
            + _stat_box("Generated", date_str.split(" ")[0], "#374151")
            + '</tr></table>'
        )
    else:
        stats_html = (
            '<table width="100%" cellpadding="0" cellspacing="0"><tr>'
            + _stat_box("Total Open MMC", total, "#1a56db")
            + _stat_box("Critical (>7d)", critical_count, "#b91c1c")
            + _stat_box("Warning (3-6d)", warning_count, "#d97706")
            + _stat_box("Fleet", escape(fleet_type), "#374151")
            + '</tr></table>'
        )

    if table_rows_html:
        table_html = (
            f'<table width="100%" cellpadding="0" cellspacing="0" '
            f'style="border-collapse:collapse;margin-top:10px;">'
            f'<thead>{_table_header(table_cols)}</thead>'
            f'<tbody>{table_rows_html}</tbody></table>'
        )
    else:
        table_html = (
            '<p style="font-family:Arial,sans-serif;font-size:13px;color:#6b7280;'
            'padding:16px 0;">No open items for this fleet at this time. ✅</p>'
        )

    return f"""<!DOCTYPE html>
<html>
<head><meta charset="utf-8"></head>
<body style="margin:0;padding:0;background-color:#f4f4f4;">
<center>
<table border="0" cellpadding="0" cellspacing="0" width="100%"
       style="background-color:#f4f4f4;">
  <tr>
    <td align="center" style="padding:20px;">
      <table border="0" cellpadding="0" cellspacing="0" width="100%"
             style="max-width:760px;background-color:#ffffff;border:1px solid #dddddd;">
        <tr>
          <td align="center" bgcolor="{banner_colour}" style="padding:20px;">
            <h2 style="margin:0;font-family:'Segoe UI',Arial,sans-serif;
                        color:#ffffff;font-size:20px;">{escape(alert_label)}</h2>
            <p style="margin:5px 0 0;font-family:'Segoe UI',Arial,sans-serif;
                       color:#ffffff;opacity:0.9;font-size:12px;">
              HTC End-to-End Visibility Monitoring &nbsp;|&nbsp; Fleet: {escape(fleet_type)}
              &nbsp;|&nbsp; {date_str}
            </p>
          </td>
        </tr>
        <tr>
          <td style="padding:20px 24px 8px;">
            {stats_html}
          </td>
        </tr>
        <tr>
          <td style="padding:8px 24px 20px;overflow-x:auto;">
            {table_html}
          </td>
        </tr>
        <tr>
          <td style="padding:0 24px 24px;font-family:Arial,sans-serif;
                     font-size:12px;color:#4b5563;line-height:1.6;">
            <p style="margin:0;">
              This is your subscribed {escape(alert_label)} digest for the
              <strong>{escape(fleet_type)}</strong> fleet. Manage or cancel this
              subscription anytime from the Email Subscriptions tab of the
              HTC Visibility Monitoring portal.
            </p>
          </td>
        </tr>
        <tr>
          <td align="center"
              style="padding:12px 28px;background:#f9f9f9;
                     border-top:1px solid #eeeeee;
                     font-family:Arial,sans-serif;font-size:11px;color:#888888;">
            Ethiopian Airlines Group | Aircraft Engineering &amp; Planning
            &nbsp;&middot;&nbsp; Automated subscription report. Do not reply.
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</center>
</body>
</html>"""


# ─────────────────────────────────────────────────────────────────────────────
# Public sender
# ─────────────────────────────────────────────────────────────────────────────

def send_subscription_digest(db_path: str, to_email: str, fleet_type: str,
                              alert_type: str, rows: list):
    """Send a single subscription digest email. Returns (ok: bool, message: str)."""
    to_email = (to_email or "").strip().lower()
    if not to_email or "@" not in to_email:
        return False, "Invalid recipient email address."

    alert_label = ALERT_LABELS.get(alert_type, alert_type)
    subject = f"HTC Monitoring — {alert_label} — {fleet_type} — {datetime.now().strftime('%Y-%m-%d')}"
    html_body = build_digest_html(fleet_type, alert_type, rows)

    last_exc = None
    for attempt in range(1, _MAX_RETRIES + 1):
        try:
            account = _build_exchange_account(db_path)
            msg = Message(
                account=account,
                folder=account.sent,
                subject=subject,
                body=HTMLBody(html_body),
                to_recipients=[Mailbox(email_address=to_email)],
            )
            msg.send()
            logger.info(
                "HTC subscription digest sent to %s (fleet=%s, alert=%s, rows=%d, attempt %d/%d).",
                to_email, fleet_type, alert_type, len(rows), attempt, _MAX_RETRIES,
            )
            return True, "Sent."
        except Exception as exc:
            last_exc = exc
            logger.warning(
                "HTC digest send attempt %d/%d failed (to=%s): %s",
                attempt, _MAX_RETRIES, to_email, exc,
            )
            if attempt < _MAX_RETRIES:
                time.sleep(_RETRY_DELAY_S)

    logger.error(
        "HTC digest send FAILED after %d attempt(s) (to=%s): %s",
        _MAX_RETRIES, to_email, last_exc, exc_info=True,
    )
    return False, str(last_exc)


# ─────────────────────────────────────────────────────────────────────────────
# Quick CLI test
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import sys
    logging.basicConfig(level=logging.INFO)

    if len(sys.argv) < 2:
        print("Usage: python htc_email.py test-connection <db_path>")
        print("       python htc_email.py send <db_path> <to@example.com> <fleet> <ALERT_B|ALERT_C>")
        sys.exit(1)

    cmd = sys.argv[1]
    if cmd == "test-connection":
        ok, msg = test_owa_connection(sys.argv[2])
        print(("OK: " if ok else "FAIL: ") + msg)
    elif cmd == "send":
        _, _, db_path, to_addr, fleet, alert_type = sys.argv
        ok, msg = send_subscription_digest(db_path, to_addr, fleet, alert_type, [])
        print(("OK: " if ok else "FAIL: ") + msg)