"""
Verification of Gmail starring for the SEO Walk-In report e-mails
(common/gmail_star.py + common/email_utils.send + the script's STAR_* settings).

Offline: SMTP and Gmail IMAP are replaced by fakes — nothing is sent, no mailbox
is opened, no real credentials are read (a fake email_config is injected).
Run from the project root:
    python seo_validation\\verify_gmail_star.py
"""
import datetime as dt
import os
import random
import sys
import tempfile
import types

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, "common"))
sys.path.insert(0, os.path.join(ROOT, "seo_reports"))
import paths                                                    # noqa: E402
paths.LOGS_DIR = paths.Path(tempfile.mkdtemp(prefix="seo_logs_"))   # never write the real logs/
paths.OUTPUT_DIR = paths.Path(tempfile.mkdtemp(prefix="seo_out_"))

INFO = "info@intellibiinnovationstechnologies.in"
OTHER = "seo.team@example.com"                                  # synthetic recipient

# fake credentials/email_config.py (test values only)
fake_ec = types.ModuleType("email_config")
fake_ec.GMAIL_SENDER, fake_ec.GMAIL_APP_PASS = INFO, "test-pass-info"
fake_ec.GMAIL_SENDER_DIGITAL, fake_ec.GMAIL_APP_PASS_DIGITAL = "digital@example.com", "test-pass-digital"
sys.modules["email_config"] = fake_ec

import email_utils as EU                                        # noqa: E402
import gmail_star as GS                                         # noqa: E402
import report_periods as RP                                     # noqa: E402

FAIL = []


def check(label, got, want):
    ok = got == want
    print(f"[{'pass' if ok else 'FAIL'}] {label}: {got!r}" + ("" if ok else f"  (expected {want!r})"))
    if not ok:
        FAIL.append(label)


# ── fakes ────────────────────────────────────────────────────────────────────
SENT, STARRED = [], []
SMTP_FAIL = [False]


class FakeSMTP:
    def __init__(self, *a, **k): pass
    def __enter__(self): return self
    def __exit__(self, *a): return False
    def login(self, u, p): self.user = u
    def sendmail(self, frm, to, raw):
        if SMTP_FAIL[0]:
            raise OSError("synthetic SMTP failure")
        SENT.append({"from": frm, "to": list(to), "raw": raw})


EU.smtplib.SMTP_SSL = FakeSMTP
_real_star = GS.star_sent_message
STAR_MODE = ["ok"]


def fake_star(mailbox, pw, msgid, subject=None, log=print, delays=None):
    STARRED.append({"mailbox": mailbox, "pw": pw, "msgid": msgid, "subject": subject})
    if STAR_MODE[0] == "raise":
        raise RuntimeError("synthetic IMAP crash")
    return STAR_MODE[0] == "ok"


GS.star_sent_message = fake_star


def reset():
    SENT.clear(); STARRED.clear(); SMTP_FAIL[0] = False; STAR_MODE[0] = "ok"


def hdr(raw, name):
    import email
    import email.policy
    v = email.message_from_string(raw, policy=email.policy.default).get(name)
    return str(v) if v is not None else None


SUBJ = "IntelliBI SEO Walk-In — Weekly (05-Oct–08-Oct-2026)"
BODY = "<p>synthetic</p>"

print("== 1. Info is a recipient -> starred in the Info mailbox only ==")
reset()
r = EU.send(SUBJ, BODY, [OTHER, INFO], INFO, star_mailbox=INFO)
mid = hdr(SENT[0]["raw"], "Message-ID") if SENT else None
check("sent (result True), same recipients, same subject (no ★ added)",
      (r, SENT[0]["to"], hdr(SENT[0]["raw"], "Subject") == SUBJ),
      (True, [OTHER, INFO], True))
check("To header unchanged", hdr(SENT[0]["raw"], "To"), f"{OTHER}, {INFO}")
check("unique Message-ID stamped on the sender's domain",
      bool(mid) and mid.endswith("@intellibiinnovationstechnologies.in>"), True)
check("starred exactly once, in the Info mailbox with its own app password, by that Message-ID",
      [(s["mailbox"], s["pw"], s["msgid"], s["subject"]) for s in STARRED], [(INFO, "test-pass-info", mid, SUBJ)])

print("\n== 2. Starring off / not applicable ==")
reset()
r = EU.send(SUBJ, BODY, [OTHER, INFO], INFO)                    # star_mailbox=None
check("star_mailbox=None: sent, no Message-ID added, nothing starred",
      (r, hdr(SENT[0]["raw"], "Message-ID"), len(STARRED)), (True, None, 0))
reset()
r = EU.send(SUBJ, BODY, [OTHER], INFO, star_mailbox=INFO)
check("Info not a recipient: sent, never starred (other recipients untouched)",
      (r, len(SENT), hdr(SENT[0]["raw"], "Message-ID"), len(STARRED)), (True, 1, None, 0))
reset()
r = EU.send(SUBJ, BODY, [OTHER, INFO, "nobody@example.com"], INFO, star_mailbox="nobody@example.com")
check("a mailbox without its own app password is never starred with another account's",
      (r, len(STARRED)), (True, 0))
reset()
m = EU.send(SUBJ, BODY, [OTHER, INFO], INFO, dry_run=True, star_mailbox=INFO)
check("dry_run: message built, not sent, not starred", (m.get("Subject"), len(SENT), len(STARRED)), (SUBJ, 0, 0))

print("\n== 3. Failures never change the send result ==")
reset(); SMTP_FAIL[0] = True
check("SMTP failure: result False, nothing starred", (EU.send(SUBJ, BODY, [INFO], INFO, star_mailbox=INFO), len(STARRED)),
      (False, 0))
reset(); STAR_MODE[0] = "fail"
check("star returns False: result still True", EU.send(SUBJ, BODY, [INFO], INFO, star_mailbox=INFO), True)
reset(); STAR_MODE[0] = "raise"
check("star raises: result still True", EU.send(SUBJ, BODY, [INFO], INFO, star_mailbox=INFO), True)
reset()
_saved = EU._gmail_star
EU._gmail_star = lambda: None
check("gmail_star.py unavailable: sent, not starred", (EU.send(SUBJ, BODY, [INFO], INFO, star_mailbox=INFO), len(STARRED)),
      (True, 0))
EU._gmail_star = _saved

print("\n== 4. The shared gmail_star.py really stars (fake Gmail IMAP) ==")


class FakeIMAP:
    calls = []
    def __init__(self, host, port, timeout=None): FakeIMAP.calls.append(("connect", host, port))
    def login(self, u, p): FakeIMAP.calls.append(("login", u))
    def list(self): return "OK", [b'(\\HasNoChildren \\All) "/" "[Gmail]/All Mail"']
    def select(self, box, readonly=False): FakeIMAP.calls.append(("select", box, readonly)); return "OK", [b"1"]
    def noop(self): return "OK", [b""]
    def uid(self, cmd, *a):
        FakeIMAP.calls.append(("uid", cmd) + a)
        return ("OK", [b"42"]) if cmd == "SEARCH" else ("OK", [b""])
    def logout(self): FakeIMAP.calls.append(("logout",))


_real_imap = GS.imaplib.IMAP4_SSL
GS.imaplib.IMAP4_SSL = FakeIMAP
GS._GIVE_UP.clear()
logs = []
ok = _real_star(INFO, "test-pass-info", "<abc@intellibiinnovationstechnologies.in>", SUBJ, log=logs.append, delays=())
check("real star_sent_message: logs in to Info, opens All Mail, sets \\Flagged on the found message",
      (ok, ("login", INFO) in FakeIMAP.calls, ("select", '"[Gmail]/All Mail"', False) in FakeIMAP.calls,
       ("uid", "STORE", b"42", "+FLAGS", "(\\Flagged)") in FakeIMAP.calls),
      (True, True, True, True))
GS.imaplib.IMAP4_SSL = _real_imap

print("\n== 5. Script: Weekly and Monthly e-mails are starred in Info ==")
import pySEOWalkInAnalysisReport as REP                         # noqa: E402
check("defaults: STAR_EMAIL_IN_GMAIL on, STAR_MAILBOX = Info", (REP.STAR_EMAIL_IN_GMAIL, REP.STAR_MAILBOX), (True, INFO))
import config_loader as CFG                                     # noqa: E402
check("Info is one of the configured report recipients (config.yaml)",
      INFO.lower() in [str(x).lower() for x in CFG.get("email.recipients", [])], True)

rnd = random.Random(3)
recs, d = [], dt.date(2025, 9, 1)
while d <= dt.date(2026, 10, 8):
    for _ in range(rnd.randint(0, 4)):
        recs.append({"date": d, "lead_source": "Google Search", "technology": "Power BI",
                     "lead_type": "Student", "mobile": "", "name": "x", "tab": "Walk-In New"})
    d += dt.timedelta(days=1)
fake_g = types.ModuleType("google_utils"); fake_g.get_services = lambda: (None, None)
sys.modules["google_utils"] = fake_g
REP.walkin_data.load_walkins = lambda *a, **k: recs
_cfg_get = CFG.get
CFG.get = lambda key, default=None: [OTHER, INFO] if key == "email.recipients" else (
    INFO if key == "email.sender" else _cfg_get(key, default))      # synthetic recipients for the run
reset()
ok = REP.run([RP.weekly(dt.date(2026, 10, 8)), RP.monthly(dt.date(2026, 10, 8))], upload=False, send_email=True)
subjects = [hdr(s["raw"], "Subject") for s in SENT]
check("run OK: Weekly + Monthly e-mails sent to the unchanged recipient list",
      (ok, len(SENT), [s["to"] for s in SENT]), (True, 2, [[OTHER, INFO], [OTHER, INFO]]))
check("each e-mail starred once, in Info only, by its own Message-ID",
      [(s["mailbox"], s["msgid"]) for s in STARRED], [(INFO, hdr(x["raw"], "Message-ID")) for x in SENT])
check("subjects unchanged (Weekly / Monthly, no ★)",
      [("— Weekly (" in s or "— Monthly (" in s) and "★" not in s and s.startswith("IntelliBI SEO Walk-In") for s in subjects],
      [True, True])
REP.STAR_EMAIL_IN_GMAIL = False
reset()
REP.run([RP.weekly(dt.date(2026, 10, 8))], upload=False, send_email=True)
check("STAR_EMAIL_IN_GMAIL = False: sent, nothing starred, no Message-ID added",
      (len(SENT), len(STARRED), hdr(SENT[0]["raw"], "Message-ID")), (1, 0, None))
REP.STAR_EMAIL_IN_GMAIL = True
CFG.get = _cfg_get

base_dir = os.environ.get("SEO_BASE_DIR")
if base_dir:                        # optional: compare with the pre-change module
    import importlib.util
    spec = importlib.util.spec_from_file_location("eu_old", os.path.join(base_dir, "common", "email_utils.py"))
    old = importlib.util.module_from_spec(spec); spec.loader.exec_module(old)
    old.smtplib.SMTP_SSL = FakeSMTP
    reset()
    old.send(SUBJ, BODY, [OTHER, INFO], INFO); EU.send(SUBJ, BODY, [OTHER, INFO], INFO)
    import email
    a, b = (email.message_from_string(s["raw"]) for s in SENT)
    check("starring off: same headers and body as before the change",
          ([(k, v) for k, v in a.items() if k != "Content-Type"], [p.get_payload() for p in a.walk() if not p.is_multipart()]),
          ([(k, v) for k, v in b.items() if k != "Content-Type"], [p.get_payload() for p in b.walk() if not p.is_multipart()]))

print("\nALL CHECKS PASSED" if not FAIL else f"\n{len(FAIL)} CHECK(S) FAILED: {FAIL}")
sys.exit(1 if FAIL else 0)
