#!/usr/bin/env python3
"""
Consulting client onboarding workflow (The Pathway Group).

  1. new     WeChat text (or name/email) -> AI extracts client info -> fills the
             contract PDF -> emails it with the signing instructions.
  2. watch   Checks your inbox for the signed contract coming back -> saves it ->
             sends the client a calendar invite for the consulting session.
  3. invite  Sends the calendar invite manually (if you'd rather not use watch).
  4. list    Shows every client and where they are in the pipeline.

Run `python onboard.py <command> --help` for options.
"""
import argparse
import email
import imaplib
import io
import json
import mimetypes
import os
import re
import smtplib
import sys
import uuid
import urllib.request
from datetime import datetime, timedelta
from email.message import EmailMessage
from email.utils import parseaddr
from pathlib import Path
from zoneinfo import ZoneInfo

import pdfplumber
from pypdf import PdfReader, PdfWriter
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas

HERE = Path(__file__).resolve().parent
CONFIG_PATH = HERE / "config.json"
DB_PATH = HERE / "clients.json"
OUT_DIR = HERE / "contracts"


# --------------------------------------------------------------------------- #
# Config + client records
# --------------------------------------------------------------------------- #
def load_config():
    if not CONFIG_PATH.exists():
        sys.exit("Missing config.json. Copy config.example.json to config.json and fill it in.")
    cfg = json.loads(CONFIG_PATH.read_text())
    cfg["email_password"] = os.environ.get(cfg.get("email_password_env", "ONBOARD_EMAIL_PASSWORD"), "")
    return cfg


def session_dt(client):
    return datetime.fromisoformat(client["session_start"]).astimezone(
        ZoneInfo(client.get("tz", "America/Los_Angeles")))


def load_db():
    return json.loads(DB_PATH.read_text()) if DB_PATH.exists() else {}


def save_db(db):
    DB_PATH.write_text(json.dumps(db, indent=2, ensure_ascii=False))


# --------------------------------------------------------------------------- #
# Step 1a: AI extraction from a pasted WeChat message
# --------------------------------------------------------------------------- #
EXTRACT_PROMPT = """Extract the client's contact details from this WeChat message.
Return ONLY a JSON object, no other text:
{"first_name": "...", "last_name": "...", "email": "..."}
Use the name in English/pinyin if given; for a Chinese name like 张伟, family name is last_name.
Use null for anything not present. Do not guess an email address.

Message:
"""


def ai_extract(text, cfg):
    key = os.environ.get("ANTHROPIC_API_KEY")
    if not key:
        return None
    body = json.dumps({
        "model": cfg.get("ai_model", "claude-haiku-5-5"),
        "max_tokens": 200,
        "messages": [{"role": "user", "content": EXTRACT_PROMPT + text}],
    }).encode()
    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages", data=body,
        headers={"x-api-key": key, "anthropic-version": "2023-06-01",
                 "content-type": "application/json"})
    with urllib.request.urlopen(req, timeout=30) as r:
        data = json.load(r)
    out = "".join(b.get("text", "") for b in data["content"])
    out = re.sub(r"```(json)?", "", out).strip()
    return json.loads(out)


def fallback_extract(text):
    """No API key: find the email, and the name for common phrasings like
    'my name is ...', 'I'm ...', '我叫...', '我的名字是...'."""
    m = re.search(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", text)
    first = last = None
    n = re.search(r"(?:我的名字是|我的名字叫|名字是|我叫|我是|姓名[:：]?|my name is|my name's|"
                  r"i am|i'm|this is|name[:：])\s*([A-Za-z][A-Za-z .'-]*?)"
                  r"(?=\s*(?:[,，.。;；!！?？\n]|\band\b|和|邮箱|email|$))", text, flags=re.I)
    if n:
        parts = n.group(1).split()
        if len(parts) >= 2:
            first, last = " ".join(parts[:-1]), parts[-1]
        elif parts:
            first = parts[0]
    return {"first_name": first, "last_name": last, "email": m.group(0) if m else None}

# --------------------------------------------------------------------------- #
# Step 1b: Fill the contract PDF
# --------------------------------------------------------------------------- #
def ordinal(n):
    return f"{n}{'th' if 11 <= n % 100 <= 13 else {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')}"


def fmt_contract_date(d):          # Oct 18th, 2026
    return f"{d:%b} {ordinal(d.day)}, {d.year}"


def fmt_session(dt):               # Wednesday Oct 21st 8pm PDT
    hour = dt.strftime("%I").lstrip("0")
    t = f"{hour}{dt:%p}".lower() if dt.minute == 0 else f"{hour}:{dt:%M}{dt:%p}".lower()
    return f"{dt:%A} {dt:%b} {ordinal(dt.day)} {t} {dt.tzname()}"


def _find_phrase(words, phrase):
    toks = phrase.split()
    for i in range(len(words) - len(toks) + 1):
        if [w["text"] for w in words[i:i + len(toks)]] == toks:
            return i, i + len(toks)
    return None


def fill_contract(template, out_path, fields):
    """
    fields: list of (placeholder phrase, replacement text).
    Each placeholder is whited out together with the rest of its line, and the
    replacement + original trailing words are redrawn in the same font/size/colour,
    shrinking slightly if a long name would run past the right margin.
    """
    reader = PdfReader(template)
    overlays = {}
    found = set()
    with pdfplumber.open(template) as pdf:
        for pno, page in enumerate(pdf.pages):
            words = page.extract_words(extra_attrs=["size", "non_stroking_color"])
            if not words:
                continue
            right_margin = max(w["x1"] for w in words)
            H = float(page.height)
            for phrase, repl in fields:
                hit = _find_phrase(words, phrase)
                if not hit:
                    continue
                found.add(phrase)
                a, b = hit
                first = words[a]
                line = [w for w in words if abs(w["top"] - first["top"]) < 1.5]
                trailing = [w for w in line if w["x0"] >= words[b - 1]["x1"] - 0.5]
                text = " ".join([repl] + [w["text"] for w in trailing])
                x0 = first["x0"]
                x1 = max([words[b - 1]["x1"]] + [w["x1"] for w in trailing])
                # baseline from the first character of the placeholder
                ch = next(c for c in page.chars
                          if abs(c["x0"] - x0) < 0.6 and abs(c["top"] - first["top"]) < 1.5)
                baseline = ch["matrix"][5]
                size = first["size"]
                avail = right_margin - x0
                while stringWidth(text, "Helvetica", size) > avail and size > 8:
                    size -= 0.25
                # Remove the template's yellow highlight behind this placeholder too
                top, bottom = first["top"] - 1, first["bottom"] + 1
                for r in page.rects:
                    fill = r.get("non_stroking_color")
                    if (fill and len(fill) == 3 and fill[2] < 0.5 and fill[0] > 0.8
                            and r["top"] < first["bottom"] and r["bottom"] > first["top"]
                            and r["x0"] < x1 and r["x1"] > x0):
                        top, bottom = min(top, r["top"]), max(bottom, r["bottom"])
                        x1 = max(x1, r["x1"])
                color = first.get("non_stroking_color") or (0, 0, 0)
                if len(color) == 1:
                    color = (color[0],) * 3
                overlays.setdefault(pno, []).append(
                    (x0, H - bottom - 0.5, x1 - x0 + 1, bottom - top + 1,
                     baseline, text, size, color))
    missing = [p for p, _ in fields if p not in found]
    if missing:
        sys.exit(f"Could not find these placeholders in the template: {missing}")

    writer = PdfWriter()
    for pno, page in enumerate(reader.pages):
        if pno in overlays:
            buf = io.BytesIO()
            w, h = float(page.mediabox.width), float(page.mediabox.height)
            c = canvas.Canvas(buf, pagesize=(w, h))
            for (x, y, bw, bh, base, text, size, color) in overlays[pno]:
                c.setFillColorRGB(1, 1, 1)
                c.rect(x - 0.5, y, bw, bh, stroke=0, fill=1)
                c.setFillColorRGB(*color[:3])
                c.setFont("Helvetica", size)   # metric-compatible with the template's Arial
                c.drawString(x, base, text)
            c.save()
            buf.seek(0)
            page.merge_page(PdfReader(buf).pages[0])
        writer.add_page(page)
    with open(out_path, "wb") as f:
        writer.write(f)


# --------------------------------------------------------------------------- #
# Email helpers
# --------------------------------------------------------------------------- #
def attach_file(msg, path):
    ctype, _ = mimetypes.guess_type(str(path))
    maintype, subtype = (ctype or "application/octet-stream").split("/", 1)
    msg.add_attachment(Path(path).read_bytes(), maintype=maintype, subtype=subtype,
                       filename=Path(path).name)


MAIL_APPLESCRIPT = r'''
on run argv
    set subj to item 1 of argv
    set bodyText to item 2 of argv
    set toAddr to item 3 of argv
    set bccAddr to item 4 of argv
    set senderAddr to item 5 of argv
    if (count of argv) > 5 then
        set fileList to items 6 thru -1 of argv
    else
        set fileList to {}
    end if
    tell application "Mail"
        set m to make new outgoing message with properties {subject:subj, content:bodyText & return & return, visible:true}
        tell m
            make new to recipient at end of to recipients with properties {address:toAddr}
            if bccAddr is not "" then make new bcc recipient at end of bcc recipients with properties {address:bccAddr}
            if senderAddr is not "" then set sender to senderAddr
        end tell
        repeat with f in fileList
            tell content of m
                make new attachment with properties {file name:((POSIX file (f as text)) as alias)} at after last paragraph
            end tell
            delay 0.5
        end repeat
        activate
    end tell
end run
'''


def open_in_mail_app(msg, cfg):
    """Build a ready-to-send draft in the Mac Mail app. You review it and click Send."""
    import subprocess
    folder = OUT_DIR / "outgoing" / datetime.now().strftime("%Y%m%d_%H%M%S")
    folder.mkdir(parents=True, exist_ok=True)
    files = []
    for part in msg.iter_attachments():
        f = folder / part.get_filename()
        f.write_bytes(part.get_payload(decode=True))
        if f.suffix.lower() in (".png", ".jpg", ".jpeg"):
            # Mac Mail shows images inside the email body; a PDF stays a normal attachment
            from PIL import Image
            pdf = f.with_suffix(".pdf")
            Image.open(f).convert("RGB").save(pdf, "PDF", resolution=150)
            f = pdf
        files.append(str(f))
    body = msg.get_body(preferencelist=("plain",)).get_content().rstrip()
    sender = f"{cfg['sender_name']} <{cfg['email_address']}>"
    args = ["osascript", "-e", MAIL_APPLESCRIPT, msg["Subject"], body, msg["To"],
            msg["Bcc"] or "", sender] + files
    result = subprocess.run(args, capture_output=True, text=True)
    if result.returncode == 0:
        print(f"Draft opened in Mail for {msg['To']}. Check it, then click Send.")
    else:
        print("Couldn't open the draft in Mail automatically:")
        print("  " + result.stderr.strip())
        print(f"The attachments are in {folder} (opening it now). "
              "Create the email by hand and drag them in.")
        subprocess.run(["open", str(folder)])


def send(msg, cfg, dry_run):
    if cfg.get("send_mode") == "mail_app" and not dry_run:
        open_in_mail_app(msg, cfg)
        return
    if dry_run:
        OUT_DIR.mkdir(exist_ok=True)
        p = OUT_DIR / f"preview_{uuid.uuid4().hex[:6]}.eml"
        p.write_bytes(bytes(msg))
        print(f"[dry run] Email not sent. Preview saved to {p}")
        return
    if not cfg["email_password"]:
        sys.exit(f"Set the {cfg.get('email_password_env')} environment variable to your email app password.")
    with smtplib.SMTP(cfg["smtp_host"], cfg.get("smtp_port", 587)) as s:
        s.starttls()
        s.login(cfg["email_address"], cfg["email_password"])
        s.send_message(msg)
    print(f"Sent: {msg['Subject']} -> {msg['To']}")


# --------------------------------------------------------------------------- #
# Command: new
# --------------------------------------------------------------------------- #
US_TIMEZONES = {
    "PST": "America/Los_Angeles", "PDT": "America/Los_Angeles", "PT": "America/Los_Angeles",
    "PACIFIC": "America/Los_Angeles",
    "MST": "America/Denver", "MDT": "America/Denver", "MT": "America/Denver", "MOUNTAIN": "America/Denver",
    "AZ": "America/Phoenix", "ARIZONA": "America/Phoenix",
    "CST": "America/Chicago", "CDT": "America/Chicago", "CT": "America/Chicago", "CENTRAL": "America/Chicago",
    "EST": "America/New_York", "EDT": "America/New_York", "ET": "America/New_York", "EASTERN": "America/New_York",
    "AKST": "America/Anchorage", "AKDT": "America/Anchorage", "HST": "Pacific/Honolulu",
}
DATE_FORMATS = ["%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%B %d %Y", "%b %d %Y", "%m/%d", "%B %d", "%b %d"]


def _clean(text):
    t = text.strip().replace(",", " ")
    t = re.sub(r"\b([ap])\.?\s*m\.?(?=\s|$)", r"\1m", t, flags=re.I)   # a.m. / p.m. -> am / pm
    t = re.sub(r"(\d)(st|nd|rd|th)\b", r"\1", t, flags=re.I)             # 19th -> 19
    t = re.sub(r"\bsept\b", "sep", t, flags=re.I)
    t = re.sub(r"\bat\b", " ", t, flags=re.I)
    return " ".join(t.split())


def parse_date(text, tz):
    """Accepts 10/19/2026, 10/19, 2026-10-19, October 19th 2026, Oct 19."""
    t = _clean(text)
    for f in DATE_FORMATS:
        try:
            d = datetime.strptime(t, f)
        except ValueError:
            continue
        if "%Y" not in f and "%y" not in f:      # no year typed: use the next upcoming date
            now = datetime.now(tz)
            d = d.replace(year=now.year)
            if d.date() < now.date():
                d = d.replace(year=now.year + 1)
        return d.replace(tzinfo=tz)
    raise ValueError


def parse_session(text, default_tz):
    """
    Accepts e.g. 'October 19th, 11:30 a.m. PST', '10/19 8pm EST', '10/19/2026 20:00 CT'.
    The timezone word is optional (your config timezone is used if left off).
    PST/PDT/PT all mean US Pacific time; the correct daylight/standard label is
    filled in automatically for that date.
    """
    t = _clean(text)
    tz = default_tz
    words = t.split()
    if words and words[-1].upper() in US_TIMEZONES:
        tz = ZoneInfo(US_TIMEZONES[words[-1].upper()])
        words = words[:-1]
    t = " ".join(words)
    m = re.search(r"(\d{1,2})(?::(\d{2}))?\s*(am|pm)?$", t, flags=re.I)
    if not m or (not m.group(2) and not m.group(3)):
        raise ValueError
    hour, minute, ampm = int(m.group(1)), int(m.group(2) or 0), (m.group(3) or "").lower()
    if ampm:
        if not 1 <= hour <= 12:
            raise ValueError
        hour = hour % 12 + (12 if ampm == "pm" else 0)
    if hour > 23 or minute > 59:
        raise ValueError
    d = parse_date(t[:m.start()], tz)
    return d.replace(hour=hour, minute=minute, tzinfo=tz)


def ask(prompt, parser, default=None):
    while True:
        raw = input(prompt).strip()
        if not raw and default is not None:
            return default
        try:
            return parser(raw)
        except ValueError:
            print("  Sorry, I couldn't read that. Please try again using the example format.")


def cmd_new(args):
    cfg = load_config()
    tz = ZoneInfo(cfg.get("timezone", "America/Los_Angeles"))

    info = {"first_name": args.first, "last_name": args.last, "email": args.email}
    wechat = args.wechat
    if not wechat and not all(info.values()):
        print("Paste the client's WeChat message, then press Enter on an empty line.")
        print("(Or just press Enter to type the details yourself.)")
        lines = []
        while True:
            line = input()
            if not line.strip():
                break
            lines.append(line)
        wechat = "\n".join(lines)
    if wechat:
        try:
            extracted = ai_extract(wechat, cfg) or fallback_extract(wechat)
        except Exception as e:
            print(f"  (AI step skipped: {e})")
            extracted = fallback_extract(wechat)
        for k, v in extracted.items():
            info[k] = info[k] or v
    for k, label in [("first_name", "Client first name"), ("last_name", "Client last name"),
                     ("email", "Client email")]:
        if args.first and args.last and args.email:
            break
        default = info.get(k)
        while True:
            hint = f" [press Enter for {default}]" if default else ""
            val = input(f"{label}{hint}: ").strip() or (default or "")
            if k == "email" and re.fullmatch(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", val):
                break
            if k != "email" and re.fullmatch(r"[A-Za-z][A-Za-z .'-]*", val):
                break
            print("  Please enter " + ("a valid email address." if k == "email" else
                  "the name in English letters (the contract can't show Chinese characters)."))
        info[k] = val

    today = datetime.now(tz)
    if args.date:
        contract_date = parse_date(args.date, tz)
    else:
        contract_date = ask(
            f"Contract date (e.g. Oct 18 2026 or 10/18/2026) [press Enter for today, {today:%m/%d/%Y}]: ",
            lambda x: parse_date(x, tz), default=today)
    if args.session:
        session = parse_session(args.session, tz)
    else:
        session = ask("Session date, time and timezone (e.g. Oct 21 8pm PST, 10/21/2026 11:30am EST): ",
                      lambda x: parse_session(x, tz))
    if session < today:
        print("  Note: that session time is in the past.")
    if args.rate:
        rate = args.rate
    else:
        undergrad = input("Is the client a current undergrad? [y/N]: ").strip().lower() == "y"
        rate = cfg.get("undergrad_hourly_rate", "120") if undergrad else cfg.get("default_hourly_rate", "150")
    full_name = f"{info['first_name']} {info['last_name']}"

    print("\n--- Please check ---")
    print(f"Client:        {full_name}")
    print(f"Email:         {info['email']}")
    print(f"Contract date: {fmt_contract_date(contract_date)}")
    print(f"Session:       {fmt_session(session)} ({args.duration} min)")
    if session.astimezone(tz).utcoffset() != session.utcoffset():
        print(f"               = {fmt_session(session.astimezone(tz))} your time")
    print(f"Rate:          ${rate}/hour")
    if not args.yes and input("Create and send contract? [y/N] ").strip().lower() != "y":
        sys.exit("Cancelled.")

    if not args.dry_run and cfg.get("send_mode") != "mail_app" and not cfg["email_password"]:
        sys.exit(f"\nYour email password isn't set yet. In this terminal, run:\n"
                 f'  export {cfg.get("email_password_env")}="your email password"\n'
                 f"then run this command again. (Use --dry-run to test without sending.)")

    OUT_DIR.mkdir(exist_ok=True)
    safe = re.sub(r"[^A-Za-z0-9]+", "_", full_name).strip("_") or "client"
    out_pdf = OUT_DIR / f"Career_Consulting_Agreement_{safe}_{contract_date:%Y%m%d}.pdf"
    fill_contract(HERE / cfg["template"], out_pdf, [
        ("Oct 18th, 2026,", fmt_contract_date(contract_date) + ","),
        ("Client First Name, Client Last Name", full_name),
        ("Wednesday Oct 14th 8pm PST.", fmt_session(session) + "."),
        ("$1XX", f"${rate}"),
        ("XX@XX", info["email"]),
        ("XXX", full_name),
    ])
    print(f"Contract created: {out_pdf}")

    msg = EmailMessage()
    msg["From"] = f"{cfg['sender_name']} <{cfg['email_address']}>"
    msg["To"] = info["email"]
    msg["Subject"] = cfg.get("contract_subject", "Career Consulting Agreement")
    msg.set_content(cfg["contract_email_body"].format(first_name=info["first_name"]))
    attach_file(msg, out_pdf)
    for img in cfg["instruction_images"]:
        attach_file(msg, HERE / img)
    send(msg, cfg, args.dry_run)

    db = load_db()
    db[info["email"].lower()] = {
        "first_name": info["first_name"], "last_name": info["last_name"],
        "email": info["email"], "contract": str(out_pdf),
        "session_start": session.isoformat(), "duration_min": args.duration,
        "tz": str(session.tzinfo), "status": "contract_sent", "sent_at": datetime.now(tz).isoformat(),
    }
    save_db(db)


# --------------------------------------------------------------------------- #
# Calendar invite (.ics) -- shows up in Gmail/Outlook with Accept/Decline
# --------------------------------------------------------------------------- #
def build_ics(client, cfg):
    start = datetime.fromisoformat(client["session_start"]).astimezone(ZoneInfo("UTC"))
    end = start + timedelta(minutes=client["duration_min"])
    f = lambda d: d.strftime("%Y%m%dT%H%M%SZ")
    link = cfg.get("meeting_link", "")
    desc = f"Career consulting session.\\nJoin: {link}" if link else "Career consulting session."
    return "\r\n".join([
        "BEGIN:VCALENDAR", "PRODID:-//Pathway Group//Onboard//EN", "VERSION:2.0", "METHOD:REQUEST",
        "BEGIN:VEVENT",
        f"UID:{uuid.uuid4()}@pathwaygroup", f"DTSTAMP:{f(datetime.now(ZoneInfo('UTC')))}",
        f"DTSTART:{f(start)}", f"DTEND:{f(end)}",
        f"SUMMARY:{cfg.get('meeting_title', 'Career Consulting Session')} - {client['first_name']} {client['last_name']}",
        f"DESCRIPTION:{desc}", f"LOCATION:{link or 'Microsoft Teams'}",
        f"ORGANIZER;CN={cfg['sender_name']}:mailto:{cfg['email_address']}",
        f"ATTENDEE;CN={client['first_name']} {client['last_name']};RSVP=TRUE;ROLE=REQ-PARTICIPANT:mailto:{client['email']}",
        "BEGIN:VALARM", "TRIGGER:-PT30M", "ACTION:DISPLAY", "DESCRIPTION:Reminder", "END:VALARM",
        "END:VEVENT", "END:VCALENDAR", ""])


def send_invite(client, cfg, dry_run):
    msg = EmailMessage()
    msg["From"] = f"{cfg['sender_name']} <{cfg['email_address']}>"
    msg["To"] = client["email"]
    msg["Bcc"] = cfg["email_address"]          # so the invite also lands on your calendar
    msg["Subject"] = f"Invitation: {cfg.get('meeting_title', 'Career Consulting Session')}"
    msg.set_content(cfg["invite_email_body"].format(
        first_name=client["first_name"],
        session=fmt_session(session_dt(client)),
        meeting_link=cfg.get("meeting_link", "")))
    msg.add_attachment(build_ics(client, cfg).encode(), maintype="text", subtype="calendar",
                       filename="invite.ics", params={"method": "REQUEST"})
    send(msg, cfg, dry_run)


# --------------------------------------------------------------------------- #
# Command: watch  (signed contract comes back -> send invite)
# --------------------------------------------------------------------------- #
def cmd_watch(args):
    cfg = load_config()
    db = load_db()
    pending = {e: c for e, c in db.items() if c["status"] == "contract_sent"}
    if not pending:
        print("No clients waiting on a signed contract.")
        return
    if not cfg["email_password"]:
        sys.exit(f"Set the {cfg.get('email_password_env')} environment variable.")
    M = imaplib.IMAP4_SSL(cfg["imap_host"])
    M.login(cfg["email_address"], cfg["email_password"])
    M.select("INBOX")
    signed_dir = OUT_DIR / "signed"
    signed_dir.mkdir(parents=True, exist_ok=True)
    for addr, client in pending.items():
        since = datetime.fromisoformat(client["sent_at"]).strftime("%d-%b-%Y")
        _, ids = M.search(None, f'(FROM "{addr}" SINCE {since})')
        for mid in ids[0].split():
            _, data = M.fetch(mid, "(RFC822)")
            m = email.message_from_bytes(data[0][1])
            pdfs = [p for p in m.walk()
                    if (p.get_filename() or "").lower().endswith(".pdf")]
            if not pdfs:
                continue
            saved = signed_dir / f"SIGNED_{Path(client['contract']).name}"
            saved.write_bytes(pdfs[0].get_payload(decode=True))
            print(f"Signed contract received from {addr} -> {saved}")
            send_invite(client, cfg, args.dry_run)
            client["status"] = "invite_sent"
            client["signed_contract"] = str(saved)
            save_db(db)
            break
    M.logout()


def cmd_invite(args):
    cfg = load_config()
    db = load_db()
    client = db.get(args.email.lower())
    if not client:
        sys.exit("No such client. Run `python onboard.py list`.")
    send_invite(client, cfg, args.dry_run)
    client["status"] = "invite_sent"
    save_db(db)


def cmd_list(args):
    for c in load_db().values():
        print(f"{c['first_name']} {c['last_name']:<15} {c['email']:<30} "
              f"{fmt_session(session_dt(c)):<32} {c['status']}")


# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    n = sub.add_parser("new", help="Fill and send a contract to a new client")
    n.add_argument("--wechat", help="Paste the client's WeChat message (AI pulls out name + email)")
    n.add_argument("--first"), n.add_argument("--last"), n.add_argument("--email")
    n.add_argument("--session", help='Session start, e.g. "Oct 21 8pm PST" (asked if omitted)')
    n.add_argument("--duration", type=int, default=60, help="Session length in minutes (default 60)")
    n.add_argument("--rate", help="Hourly rate, e.g. 150 (asked if omitted)")
    n.add_argument("--date", help="Contract date, e.g. 10/18/2026 (asked if omitted)")
    n.add_argument("--yes", action="store_true", help="Skip the confirmation prompt")
    n.add_argument("--dry-run", action="store_true", help="Create the PDF and email preview without sending")
    n.set_defaults(func=cmd_new)

    w = sub.add_parser("watch", help="Check inbox for signed contracts and send invites")
    w.add_argument("--dry-run", action="store_true")
    w.set_defaults(func=cmd_watch)

    i = sub.add_parser("invite", help="Send the meeting invite manually")
    i.add_argument("email")
    i.add_argument("--dry-run", action="store_true")
    i.set_defaults(func=cmd_invite)

    sub.add_parser("list", help="Show all clients and their status").set_defaults(func=cmd_list)

    args = ap.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()