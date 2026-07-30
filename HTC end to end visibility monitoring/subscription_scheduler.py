"""
subscription_scheduler.py  –  HTC Email Subscription runner
=============================================================
Background thread checking, once per minute, whether any saved
email_subscriptions row is due to fire (based on frequency / day_of_week /
run_time), and if so, builds and sends the Alert B / Alert C digest via
htc_email.py.

A subscription is considered "due" when:
  - now's HH:MM matches run_time (within the polling granularity), AND
  - for weekly subscriptions, today's day-of-week matches day_of_week, AND
  - for every_3_days subscriptions, at least 3 full days have elapsed since
    last_sent_at (or it has never been sent), AND
  - last_sent_at is not already within today's (or this week's) send window,
    to avoid double-sends if the poll loop overlaps a matching minute twice.
"""

import logging
import threading
import time
from datetime import datetime

import database
import htc_email

logger = logging.getLogger(__name__)

_DAY_MAP = {0: "mon", 1: "tue", 2: "wed", 3: "thu", 4: "fri", 5: "sat", 6: "sun"}

_POLL_SECONDS = 60
_EVERY_3_DAYS_INTERVAL = 3


def _already_sent_today(sub: dict, now: datetime) -> bool:
    last_sent = sub.get("last_sent_at")
    if not last_sent:
        return False
    try:
        last_dt = datetime.fromisoformat(last_sent)
    except ValueError:
        return False
    return last_dt.date() == now.date()


def _interval_elapsed(sub: dict, now: datetime, days: int) -> bool:
    """True if `days` or more full calendar days have passed since
    last_sent_at, or if the subscription has never been sent."""
    last_sent = sub.get("last_sent_at")
    if not last_sent:
        return True
    try:
        last_dt = datetime.fromisoformat(last_sent)
    except ValueError:
        return True
    return (now.date() - last_dt.date()).days >= days


def _is_due(sub: dict, now: datetime) -> bool:
    run_time = sub.get("run_time", "")
    try:
        run_h, run_m = (int(x) for x in run_time.split(":"))
    except (ValueError, AttributeError):
        return False

    if now.hour != run_h or now.minute != run_m:
        return False

    frequency = sub.get("frequency")

    if frequency == "weekly":
        if _DAY_MAP.get(now.weekday()) != sub.get("day_of_week"):
            return False
    elif frequency == "every_3_days":
        if not _interval_elapsed(sub, now, _EVERY_3_DAYS_INTERVAL):
            return False

    if _already_sent_today(sub, now):
        return False

    return True


def _rows_for_alert_type(db_path: str, fleet: str, alert_type: str):
    if alert_type == "ALERT_B":
        return database.get_alert_b_events(db_path, fleet=fleet)
    else:
        return database.get_mmc_alerts(db_path, fleet=fleet)


def _run_subscription(db_path: str, sub: dict):
    alert_types = sub.get("alert_types") or [
        a for a in (sub.get("alert_type") or "").split(",") if a
    ]
    fleet = sub.get("fleet_type")
    any_failed = False

    for alert_type in alert_types:
        try:
            rows = _rows_for_alert_type(db_path, fleet, alert_type)
            ok, msg = htc_email.send_subscription_digest(
                db_path,
                to_email=sub.get("email"),
                fleet_type=fleet,
                alert_type=alert_type,
                rows=rows,
            )
            if ok:
                logger.info(
                    "Subscription #%s digest sent to %s (fleet=%s, alert=%s, rows=%d).",
                    sub["id"], sub.get("email"), fleet, alert_type, len(rows),
                )
            else:
                any_failed = True
                logger.error(
                    "Subscription #%s digest FAILED to %s (alert=%s): %s",
                    sub["id"], sub.get("email"), alert_type, msg,
                )
        except Exception as exc:
            any_failed = True
            logger.error(
                "Error running subscription #%s (alert=%s): %s",
                sub.get("id"), alert_type, exc, exc_info=True,
            )

    # Mark as sent even if one of several alert types failed, so a persistent
    # failure (e.g. bad fleet data) doesn't retry every minute for the rest
    # of the day; failures are still fully logged above for follow-up.
    database.mark_subscription_sent(db_path, sub["id"])
    if any_failed:
        logger.warning("Subscription #%s completed with at least one failed alert-type send.", sub["id"])


def _poll_loop(db_path: str):
    logger.info("HTC subscription scheduler started (poll every %ds).", _POLL_SECONDS)
    while True:
        try:
            now = datetime.now()
            subs = database.get_subscriptions(db_path, active_only=True)
            for sub in subs:
                if _is_due(sub, now):
                    _run_subscription(db_path, sub)
        except Exception as exc:
            logger.error("Subscription scheduler poll error: %s", exc, exc_info=True)
        time.sleep(_POLL_SECONDS)


def start_subscription_scheduler(db_path: str):
    t = threading.Thread(target=_poll_loop, args=(db_path,), daemon=True,
                          name="subscription-scheduler")
    t.start()
    return t