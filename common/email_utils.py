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


def build_body(period, cur, prev, gs_label, drive_link, gen_stamp):
    tot_c, tot_p = len(cur), len(prev)
    diff = tot_c - tot_p
    g = _pct(tot_c, tot_p)
    growth_txt = f"{g*100:+.1f}%" if g is not None else "New"
    rowbg = _growth_fill(tot_c, tot_p)                       # colour code from Growth/Decline %
    grow_color = "#2E7D32" if (g or 0) >= 0 else "#C0392B"

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
    {link_html}
    <p style='font-size:11px;color:{GREY};margin-top:14px'>
      Full detail — Summary, Lead Source Trend, Technology Trend, Lead Type Trend —
      is in the attached workbook. Automated report from IntelliBI SEO Automation.</p>
  </div>
</div>"""


def send(subject, html_body, recipients, sender, attachment_path=None, dry_run=False):
    """Send the email. When dry_run is True the MIME message is built and returned
    but NOT sent (used for validation)."""
    sender, app_pass = _creds(sender)
    if not recipients:
        log.warning("No recipients configured — email not sent.")
        return False
    msg = MIMEMultipart("mixed")
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = ", ".join(recipients)
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
        return True
    except Exception as e:
        log.warning("  email send failed: %s", e)
        return False


def send_report(period, cur, prev, gs_label, drive_link, gen_stamp,
                attachment_path, dry_run=False):
    """High-level: build the body from config recipients/sender and send."""
    recipients = cfg.get("email.recipients", []) or []
    sender = cfg.get("email.sender", "info@intellibiinnovationstechnologies.in")
    attach = attachment_path if cfg.get("email.attach_workbook", True) else None
    subject = (f"IntelliBI SEO Walk-In — {period.label} "
               f"({period.cur_start.strftime('%d-%b')}–{period.cur_end.strftime('%d-%b-%Y')})")
    body = build_body(period, cur, prev, gs_label, drive_link, gen_stamp)
    return send(subject, body, recipients, sender, attach, dry_run=dry_run)
