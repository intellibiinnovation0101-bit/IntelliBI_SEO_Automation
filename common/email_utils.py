"""
Email delivery (common/email_utils.py).

Sends the generated SEO Walk-In report by Gmail SMTP, reusing the SAME
credentials approach as the other IntelliBI reports (credentials/email_config.py
-> GMAIL_SENDER / GMAIL_APP_PASS). A KPI-card HTML summary is included and the
.xlsx workbook is attached; the Drive link is shown when available.

Sending is OPTIONAL — controlled by config (email.enabled) and the CLI flags
--email / --no-email. Nothing is sent unless explicitly enabled.
"""
from __future__ import annotations
import os
import sys
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

import paths
import config_loader as cfg
import logging_utils

log = logging_utils.get_logger("seo_email")

# credentials/email_config.py lives beside the other secrets.
if str(paths.CREDENTIALS_DIR) not in sys.path:
    sys.path.insert(0, str(paths.CREDENTIALS_DIR))

NAVY = "#1B355E"; GOLD = "#B7791F"; GREY = "#5b6b86"


def _creds(sender_wanted: str):
    """Return (sender, app_password) from credentials/email_config.py, matching
    the configured sender to the right account (primary or _DIGITAL)."""
    try:
        import email_config as ec
    except Exception as e:
        raise RuntimeError(
            "credentials/email_config.py not found or invalid. Copy it from the "
            "Operations project (same GMAIL_SENDER / GMAIL_APP_PASS).") from e
    sw = (sender_wanted or "").strip().lower()
    pairs = [(getattr(ec, "GMAIL_SENDER", ""), getattr(ec, "GMAIL_APP_PASS", ""))]
    if hasattr(ec, "GMAIL_SENDER_DIGITAL"):
        pairs.append((ec.GMAIL_SENDER_DIGITAL, getattr(ec, "GMAIL_APP_PASS_DIGITAL", "")))
    for s, p in pairs:
        if s and s.strip().lower() == sw:
            return s, p
    # default to the primary account
    return pairs[0]


def _pct(a, b):
    return (a - b) / b if b else None


# Row background colour from Growth/Decline % (same thresholds as the report).
def _growth_fill(cur, prev):
    if prev == 0:
        return "#D5F5E3" if cur > 0 else "#FCF3CF"
    g = (cur - prev) / prev
    if g >= 0.10:
        return "#D5F5E3"       # green
    if g >= 0.0:
        return "#FCF3CF"       # yellow
    if g >= -0.10:
        return "#FAE5D3"       # orange
    return "#F5B7B1"           # red


# ── Weekly / Monthly Walk-In Lead Performance vs target (e-mail section) ──────
# Same presentation as the "Performance vs Goals" bars of the Sales lead report
# (pyConsolidatedLeadPerformanceReport.build_email_body): name + value on one
# line, a rounded bar with a target marker, a note line underneath. Figures come
# from walkin_targets.trend_for_period — the SAME rows as the workbook's
# "Weekly / Monthly Lead Trend" tab.
_ABOVE_HEX, _BELOW_HEX = "2E7D32", "C0392B"
_ABOVE_CUR_HEX, _BELOW_CUR_HEX = "66BB6A", "E57373"


def _sec(title):
    return ("<div style='font-size:11px;font-weight:700;letter-spacing:.06em;"
            "text-transform:uppercase;color:#5b6b86;margin:18px 0 8px;padding-bottom:5px;"
            f"border-bottom:1px solid #e2e8f0'>{title}</div>")


def _bar(name, value_label, pct, hexc, goal_pct, note, value_hex=None):
    pct = max(0, min(100, pct))
    goal = (f"<div style='position:absolute;top:-2px;bottom:-2px;left:{goal_pct:.0f}%;"
            "width:2px;background:#334155'></div>") if goal_pct is not None else ""
    return (
        "<div style='margin:11px 0'>"
        "<table role='presentation' width='100%' style='border-collapse:collapse'><tr>"
        f"<td style='font-size:13px;color:#1a2a48'>{name}</td>"
        f"<td style='font-size:13px;font-weight:700;color:#{value_hex or hexc};text-align:right;"
        f"white-space:nowrap'>{value_label}</td>"
        "</tr></table>"
        "<div style='height:12px;border-radius:999px;background:#eef1f6;position:relative;"
        "overflow:hidden;margin-top:5px'>"
        f"<div style='height:100%;border-radius:999px;background:#{hexc};width:{pct:.0f}%'></div>"
        f"{goal}</div>"
        f"<div style='font-size:10.5px;color:#5b6b86;margin-top:3px'>{note}</div></div>")


def lead_performance_section(trend) -> str:
    """'Weekly Walk-In Lead Performance – Last 5 Completed Weeks + Current Week'
    (or the Monthly equivalent): one bar per period, actual leads against the
    period's target (marker), green above / red at-or-below, the current period
    flagged as in progress."""
    if not trend or not trend.get("email"):
        return ""
    import walkin_targets as WT
    rows = trend["email"]
    is_week = trend["kind"] == "week"
    unit, Unit = ("week", "Week") if is_week else ("month", "Month")
    n_done = len(rows) - 1
    title = (f"{'Weekly' if is_week else 'Monthly'} Walk-In Lead Performance &ndash; "
             f"Last {n_done} Completed {Unit}s + Current {Unit}")
    scale = max([x["actual"] for x in rows] + [x["target_exact"] for x in rows] + [1]) * 1.12
    done = [x for x in rows if not x["current"]]
    above = sum(1 for x in done if x["above"])
    html = _sec(title)
    html += (f"<p style='margin:0 0 4px;color:#5b6b86;font-size:12px;line-height:1.5'>"
             f"<b style='color:#1a2a48'>{above} of {len(done)}</b> completed {unit}s above target "
             f"&nbsp;&middot;&nbsp; {WT.target_formula_text(trend['kind'], trend['monthly_target'])} "
             f"&nbsp;&middot;&nbsp; <span style='color:#{_ABOVE_HEX}'>&#9632;</span> above target "
             f"<span style='color:#{_BELOW_HEX}'>&#9632;</span> at / below target "
             f"&nbsp;&middot;&nbsp; | = target</p>")
    for x in reversed(rows):                      # newest first, current at the top
        hexc = ((_ABOVE_CUR_HEX if x["above"] else _BELOW_CUR_HEX) if x["in_progress"]
                else (_ABOVE_HEX if x["above"] else _BELOW_HEX))
        vhex = _ABOVE_HEX if x["above"] else _BELOW_HEX
        ach = f"{x['achievement']:.0f}%" if x["achievement"] is not None else "&mdash;"
        name = WT.period_label(x)
        if x["current"]:
            chip = ("In progress" if x["in_progress"] else f"Current {unit}")
            name = (f"<b>{name}</b> <span style='display:inline-block;padding:1px 7px;"
                    f"border-radius:999px;background:#1B355E;color:#ffffff;font-size:10px;"
                    f"font-weight:700;margin-left:4px'>{chip}</span>")
        note = (f"{'Weekly' if is_week else 'Monthly'} target {x['target']:g} &middot; "
                f"<b style='color:#{vhex}'>{x['status']}</b>")
        if x["in_progress"]:
            note += " &middot; " + WT.progress_text(x).replace("In progress · ", "")
        html += _bar(name, f"{x['actual']} / {x['target']:g} &middot; {ach}",
                     x["actual"] / scale * 100.0, hexc, x["target_exact"] / scale * 100.0,
                     note, value_hex=vhex)
    return html


def build_body(period, cur, prev, gs_label, drive_link, gen_stamp, trend=None):
    tot_c, tot_p = len(cur), len(prev)
    diff = tot_c - tot_p
    g = _pct(tot_c, tot_p)
    growth_txt = f"{g*100:+.1f}%" if g is not None else "New"
    rowbg = _growth_fill(tot_c, tot_p)                       # colour code from Growth/Decline %
    grow_color = "#2E7D32" if (g or 0) >= 0 else "#C0392B"
    trend_tab_note = (f", {'Weekly' if trend['kind'] == 'week' else 'Monthly'} Lead Trend"
                      if trend else "")
    perf_html = lead_performance_section(trend)
    perf_html = (perf_html + "\n    ") if perf_html else ""

    def card(label, value, value_color=NAVY):
        # card style like the attachment, tinted by the Growth/Decline % colour code
        return (f"<td style='padding:6px;width:25%'>"
                f"<div style='background:{rowbg};border:1px solid #d3dce8;"
                f"border-radius:8px;padding:14px 8px;text-align:center'>"
                f"<div style='font-size:24px;font-weight:700;color:{value_color}'>{value}</div>"
                f"<div style='font-size:12px;color:{GREY};margin-top:2px'>{label}</div></div></td>")

    link_html = ""
    if drive_link:
        link_html = (f"<p style='margin:16px 0 6px'><a href='{drive_link}' "
                     f"style='background:{NAVY};color:#fff;text-decoration:none;"
                     f"padding:9px 16px;border-radius:6px;font-size:13px'>"
                     f"Open the {period.label} report in Drive</a></p>")

    return f"""\
<div style='font-family:Arial,Helvetica,sans-serif;max-width:640px;margin:auto;color:#1F2A44'>
  <div style='background:{NAVY};color:#fff;padding:14px 18px;border-radius:8px 8px 0 0'>
    <div style='font-size:18px;font-weight:700'>IntelliBI — SEO Walk-In Analysis</div>
    <div style='font-size:12px;opacity:.85'>{period.label} Report</div>
  </div>
  <div style='border:1px solid #e2e8f0;border-top:none;border-radius:0 0 8px 8px;padding:16px 18px'>
    <p style='font-size:12px;color:{GREY};margin:0 0 10px'>
      Current: {period.cur_start.strftime('%d-%b-%Y')} &rarr; {period.cur_end.strftime('%d-%b-%Y')}
      &nbsp;|&nbsp; Previous (same {period.elapsed_days} day(s)):
      {period.prev_start.strftime('%d-%b-%Y')} &rarr; {period.prev_end.strftime('%d-%b-%Y')}
      &nbsp;|&nbsp; Generated: {gen_stamp}</p>
    <div style='font-size:14px;font-weight:700;color:{NAVY};margin:2px 0 6px'>Total Walk-Ins</div>
    <table role='presentation' width='100%' style='border-collapse:separate'><tr>
      {card("Current", tot_c)}
      {card("Previous", tot_p)}
      {card("Difference", f"{diff:+d}")}
      {card("Growth / Decline %", growth_txt, grow_color)}
    </tr></table>
    {perf_html}{link_html}
    <p style='font-size:11px;color:{GREY};margin-top:14px'>
      Full detail — Summary, Lead Source Trend, Technology Trend, Lead Type Trend{trend_tab_note} —
      is in the attached workbook. Automated report from IntelliBI SEO Automation.</p>
  </div>
</div>"""


def _mailbox_password(mailbox: str):
    """App password of EXACTLY this Gmail account from credentials/email_config.py
    (GMAIL_SENDER or GMAIL_SENDER_DIGITAL), else None — never another account's."""
    try:
        import email_config as ec
    except Exception:                                        # noqa: BLE001
        return None
    mb = (mailbox or "").strip().lower()
    for s_attr, p_attr in (("GMAIL_SENDER", "GMAIL_APP_PASS"),
                           ("GMAIL_SENDER_DIGITAL", "GMAIL_APP_PASS_DIGITAL")):
        s = str(getattr(ec, s_attr, "") or "").strip()
        if s and s.lower() == mb:
            return getattr(ec, p_attr, "") or None
    return None


def _gmail_star():
    """common/gmail_star.py (shared with Sales / Operations), or None."""
    try:
        import gmail_star
        return gmail_star
    except Exception as e:                                   # noqa: BLE001
        log.warning("  [Email] ★ starring unavailable — common/gmail_star.py: %s", e)
        return None


def _star_log(m):
    m = str(m).strip()
    (log.warning if "not starred" in m else log.info)(m)


def _star_plan(star_mailbox, recipients):
    """(gmail_star module, mailbox, app password) when the e-mail should be starred
    in `star_mailbox` (the Info mailbox), else None. Only that mailbox is ever
    starred, and only when it is one of the recipients — other recipients' copies
    are never touched."""
    mb = (star_mailbox or "").strip()
    if not mb:
        return None
    if mb.lower() not in {str(r).strip().lower() for r in recipients}:
        log.info("  [Email] ★ not starred — %s is not a recipient of this e-mail.", mb)
        return None
    pw = _mailbox_password(mb)
    if not pw:
        log.warning("  [Email] ★ not starred — no app password for %s in "
                    "credentials/email_config.py (the e-mail is still sent).", mb)
        return None
    gs = _gmail_star()
    return (gs, mb, pw) if gs is not None else None


def send(subject, html_body, recipients, sender, attachment_path=None, dry_run=False,
         star_mailbox=None):
    """Send the email. When dry_run is True the MIME message is built and returned
    but NOT sent (used for validation).
    star_mailbox: after a SUCCESSFUL send, mark the e-mail Starred (★) in THAT Gmail
    mailbox only (the Info mailbox) via common/gmail_star.py — best-effort, never
    changes the subject, recipients, body or the send result."""
    sender, app_pass = _creds(sender)
    if not recipients:
        log.warning("No recipients configured — email not sent.")
        return False
    msg = MIMEMultipart("mixed")
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = ", ".join(recipients)
    star = _star_plan(star_mailbox, recipients) if (star_mailbox and not dry_run) else None
    if star is not None:                     # a known Message-ID lets the copy be found & starred
        msg["Message-ID"] = star[0].new_message_id(sender)
    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText("This report is best viewed as HTML.", "plain", "utf-8"))
    alt.attach(MIMEText(html_body, "html", "utf-8"))
    msg.attach(alt)
    if attachment_path and os.path.exists(attachment_path):
        with open(attachment_path, "rb") as fh:
            part = MIMEApplication(fh.read(),
                                   _subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        part.add_header("Content-Disposition", "attachment",
                        filename=os.path.basename(attachment_path))
        msg.attach(part)
    if dry_run:
        return msg
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as srv:
            srv.login(sender, app_pass)
            srv.sendmail(sender, recipients, msg.as_string())
        log.info("  emailed report to %s", ", ".join(recipients))
    except Exception as e:
        log.warning("  email send failed: %s", e)
        return False
    if star is not None:                     # best-effort; never changes the result
        try:
            gs, mb, pw = star
            gs.star_sent_message(mb, pw, msg["Message-ID"], subject, log=_star_log)
        except Exception as e:               # noqa: BLE001
            log.warning("  [Email] ★ not starred — %s (the e-mail itself was sent).", e)
    return True


def send_report(period, cur, prev, gs_label, drive_link, gen_stamp,
                attachment_path, dry_run=False, trend=None, star_mailbox=None):
    """High-level: build the body from config recipients/sender and send."""
    recipients = cfg.get("email.recipients", []) or []
    sender = cfg.get("email.sender", "info@intellibiinnovationstechnologies.in")
    attach = attachment_path if cfg.get("email.attach_workbook", True) else None
    subject = (f"IntelliBI SEO Walk-In — {period.label} "
               f"({period.cur_start.strftime('%d-%b')}–{period.cur_end.strftime('%d-%b-%Y')})")
    body = build_body(period, cur, prev, gs_label, drive_link, gen_stamp, trend=trend)
    return send(subject, body, recipients, sender, attach, dry_run=dry_run,
                star_mailbox=star_mailbox)
