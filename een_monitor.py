"""
EEN Business Offer daily digest
Source: https://een.ec.europa.eu/partnering-opportunities

Each run:
  1. Reads the newest listing pages (stops when posts are older than LOOKBACK_DAYS).
  2. Keeps only "Business Offer" profiles not seen before (by POD reference).
  3. Opens each new offer to get full details.
  4. Adds them as rows to een_offers.xlsx (full history, one row per offer).
  5. Emails one digest: Title, Location, Type of partnership, Market keywords,
     Description, Link - with a download link to the full Excel file.
  5. Saves the seen list (seen_een.json) only after the email is sent,
     so a crash never loses offers - they are simply picked up next run.

First run (no seen_een.json yet): only offers from the last 24 hours are emailed,
everything else found is marked as seen - so you don't get a huge first email.

Env variables (set as GitHub Secrets):
  GMAIL_USER          sender Gmail address
  GMAIL_APP_PASSWORD  16-character Gmail app password
  EEN_RECIPIENTS      comma-separated list of recipient emails
If GMAIL_USER is not set, the digest is saved to digest_preview.html instead.
"""

import html
import json
import os
import re
import smtplib
import sys
import time
from datetime import datetime, timezone, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

BASE = "https://een.ec.europa.eu"
LIST_URL = BASE + "/partnering-opportunities"
SEEN_FILE = os.environ.get("SEEN_FILE", "seen_een.json")
EXCEL_FILE = os.environ.get("EXCEL_FILE", "een_offers.xlsx")
EXCEL_COLUMNS = ["Date found", "Reference", "Title", "Location", "Posted", "Valid until",
                 "Target market", "Type of partnership", "Market keywords", "Description", "Link"]
EXCEL_WIDTHS = [12, 18, 50, 16, 13, 13, 30, 30, 45, 90, 50]
LOOKBACK_DAYS = float(os.environ.get("LOOKBACK_DAYS", "3"))
MAX_PAGES = int(os.environ.get("MAX_PAGES", "25"))
WANTED_TYPE = "business offer"
DELAY = 1.5  # seconds between requests, to be polite

session = requests.Session()
session.headers["User-Agent"] = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/128.0 Safari/537.36"
)


def clean(text):
    return re.sub(r"\s+", " ", text or "").strip()


def get(url, tries=3):
    for i in range(tries):
        try:
            r = session.get(url, timeout=60)
            r.raise_for_status()
            return r.text
        except Exception as e:
            print(f"  ! fetch failed ({i + 1}/{tries}) {url}: {e}")
            time.sleep(5 * (i + 1))
    raise RuntimeError(f"Could not load {url}")


def age_in_days(text):
    """'53 minutes 4 seconds ago' -> 0.04, '2 days 8 hours ago' -> 2.33"""
    t = text.lower()
    units = {"year": 365, "month": 30, "week": 7, "day": 1,
             "hour": 1 / 24, "minute": 1 / 1440, "second": 1 / 86400}
    total, found = 0.0, False
    for num, unit in re.findall(r"(\d+)\s*(year|month|week|day|hour|minute|second)", t):
        total += int(num) * units[unit]
        found = True
    return total if found else 0.0


def parse_listing(page_html):
    soup = BeautifulSoup(page_html, "html.parser")
    cards = []
    for block in soup.select(".ecl-content-block.ecl-card__content-block"):
        meta = [clean(li.get_text()) for li in block.select(".ecl-content-block__primary-meta-item")]
        link = block.select_one(".ecl-content-block__title a")
        if not link or len(meta) < 2:
            continue
        sec = [clean(x.get_text()) for x in block.select(".ecl-content-block__secondary-meta-label")]
        cards.append({
            "type": meta[0],
            "ref": meta[1],
            "title": clean(link.get_text()),
            "link": BASE + link["href"] if link["href"].startswith("/") else link["href"],
            "age_days": age_in_days(sec[0]) if sec else 0.0,
            "location": sec[1] if len(sec) > 1 else "",
        })
    return cards


def dd_values(dd):
    items = dd.select("li")
    if items:
        return [clean(li.get_text()) for li in items]
    for br in dd.find_all("br"):
        br.replace_with("\n")
    text = dd.get_text()
    lines = [re.sub(r"[ \t]+", " ", ln).strip() for ln in text.split("\n")]
    return ["\n".join(ln for ln in lines if ln)]


def parse_detail(page_html):
    soup = BeautifulSoup(page_html, "html.parser")
    fields = {}
    for dt in soup.select("dt"):
        dd = dt.find_next_sibling("dd")
        if dd is None:
            continue
        label = clean(dt.get_text())
        if label not in fields:  # "Type of partnership" appears twice - keep the first
            fields[label] = dd_values(dd)
    h1 = soup.select_one("h1")
    return {
        "title": clean(h1.get_text()) if h1 else "",
        "type": ", ".join(fields.get("Profile Type", [])),
        "location": ", ".join(fields.get("Company's Country", [])),
        "partnership": fields.get("Type of partnership", []),
        "market_keywords": fields.get("Market keywords", []),
        "description": (fields.get("Full Description") or fields.get("Short Summary") or [""])[0],
        "summary": (fields.get("Short Summary") or [""])[0],
        "validity": (fields.get("Term of Validity") or [""])[0],
        "targets": fields.get("Targeted Countries") or fields.get("Targeted countries") or [],
    }


def parse_date(text):
    for fmt in ("%d %B %Y", "%d %b %Y"):
        try:
            return datetime.strptime(clean(text), fmt).date()
        except ValueError:
            pass
    return None


def add_dates(o):
    """Posted = start of Term of Validity (fallback: date inside the POD reference, e.g. BOUA20261016001).
    Updated = when the offer appeared at the top of the EEN list, if that is later than the posting."""
    parts = re.split(r"\s+-\s+", o.get("validity") or "")
    posted = parse_date(parts[0]) if parts and parts[0] else None
    valid_until = parse_date(parts[1]) if len(parts) > 1 else None
    if not posted:
        m = re.search(r"(20\d{2})(\d{2})(\d{2})", o.get("ref", ""))
        if m:
            try:
                posted = datetime(int(m[1]), int(m[2]), int(m[3])).date()
            except ValueError:
                pass
    updated = (datetime.now(timezone(timedelta(hours=7))) - timedelta(days=o.get("age_days", 0))).date()
    o["posted"] = posted.strftime("%d %b %Y") if posted else ""
    o["valid_until"] = valid_until.strftime("%d %b %Y") if valid_until else ""
    o["updated"] = updated.strftime("%d %b %Y") if posted and (updated - posted).days >= 3 else ""
    return o


def load_seen():
    if not os.path.exists(SEEN_FILE):
        return None
    with open(SEEN_FILE, encoding="utf-8") as f:
        return json.load(f)


def save_seen(seen):
    tmp = SEEN_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(seen, f, indent=1, sort_keys=True)
    os.replace(tmp, SEEN_FILE)  # atomic - file is never half-written


def save_to_excel(offers):
    """Add new offers as rows. Skips references already in the file, so a re-run never doubles rows."""
    wb = None
    if os.path.exists(EXCEL_FILE):
        wb = load_workbook(EXCEL_FILE)
        ws = wb.active
        if [c.value for c in ws[1]][:len(EXCEL_COLUMNS)] != EXCEL_COLUMNS:
            # Old column layout - keep it as a backup and start a fresh file
            os.replace(EXCEL_FILE, EXCEL_FILE.replace(".xlsx", "_old.xlsx"))
            print("Excel: old column layout saved as backup, starting new file")
            wb = None
    if wb is None:
        wb = Workbook()
        ws = wb.active
        ws.title = "Business Offers"
        ws.append(EXCEL_COLUMNS)
        for i, w in enumerate(EXCEL_WIDTHS, 1):
            ws.column_dimensions[ws.cell(1, i).column_letter].width = w
            c = ws.cell(1, i)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="0B5394")
        ws.freeze_panes = "A2"

    existing = {ws.cell(r, 2).value for r in range(2, ws.max_row + 1)}
    today = datetime.now(timezone(timedelta(hours=7))).strftime("%Y-%m-%d")
    added = 0
    for o in offers:
        if o["ref"] in existing:
            continue
        ws.append([
            today, o["ref"], o["title"], o["location"],
            o.get("posted", ""), o.get("valid_until", ""), ", ".join(o.get("targets") or []),
            ", ".join(o["partnership"]), "\n".join(o["market_keywords"]),
            o["description"][:32000], o["link"],
        ])
        r = ws.max_row
        n = len(EXCEL_COLUMNS)
        for col in range(1, n + 1):
            ws.cell(r, col).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(r, n).hyperlink = o["link"]
        ws.cell(r, n).font = Font(color="0B5394", underline="single")
        added += 1
    ws.auto_filter.ref = f"A1:{ws.cell(1, len(EXCEL_COLUMNS)).column_letter}{ws.max_row}"

    tmp = EXCEL_FILE + ".tmp.xlsx"
    wb.save(tmp)
    os.replace(tmp, EXCEL_FILE)  # atomic - file is never half-written
    print(f"Excel: added {added} row(s), {ws.max_row - 1} offers in total")


def collect_cards():
    all_cards = []
    for page in range(MAX_PAGES):
        url = LIST_URL if page == 0 else f"{LIST_URL}?page={page}"
        print(f"Listing page {page + 1}: {url}")
        cards = parse_listing(get(url))
        if not cards:
            print("  no cards found - stopping")
            break
        all_cards.extend(cards)
        if min(c["age_days"] for c in cards) > LOOKBACK_DAYS:
            break
        time.sleep(DELAY)
    return all_cards


def short_text(o, limit=420):
    """EEN's own Short Summary if we have it, otherwise the start of the description."""
    text = clean(o.get("summary") or o.get("description") or "")
    if len(text) <= limit:
        return text
    cut = text[:limit]
    end = max(cut.rfind(". "), cut.rfind("! "), cut.rfind("? "))
    return cut[:end + 1] if end > limit * 0.5 else cut.rsplit(" ", 1)[0] + "..."


def keyword_name(k):
    """'07003001 - Wine and liquors' -> 'Wine and liquors'"""
    return re.sub(r"^\s*\d+\s*-\s*", "", k).strip()


SHEET_ID = os.environ.get("EEN_SHEET_ID", "").strip()
SHEET_KEY = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()


def offer_row(o, today):
    return [
        today, o["ref"], o["title"], o["location"],
        o.get("posted", ""), o.get("valid_until", ""), ", ".join(o.get("targets") or []),
        ", ".join(o["partnership"]), "\n".join(o["market_keywords"]),
        o["description"][:30000], o["link"],
    ]


def save_to_sheet(offers):
    """Add new offers to the Google Sheet. Skips references already there, so a re-run never doubles rows.
    On the very first use it also copies over everything already in the Excel file."""
    import gspread
    gc = gspread.service_account_from_dict(json.loads(SHEET_KEY))
    ws = gc.open_by_key(SHEET_ID).sheet1

    header = ws.row_values(1)
    if not header:
        ws.update(range_name="A1", values=[EXCEL_COLUMNS])
        ws.format("A1:K1", {"textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
                            "backgroundColor": {"red": 0.04, "green": 0.33, "blue": 0.58}})
        ws.freeze(rows=1)
        if os.path.exists(EXCEL_FILE):  # bring over the offers collected before the switch
            old = [list(r) for r in load_workbook(EXCEL_FILE).active.iter_rows(min_row=2, values_only=True)]
            old = [[("" if v is None else str(v)) for v in r] for r in old if r and r[1]]
            if old and len(old[0]) == len(EXCEL_COLUMNS):
                ws.append_rows(old, value_input_option="RAW")
                print(f"Sheet: copied {len(old)} offers from the Excel file")

    existing = set(ws.col_values(2))
    today = datetime.now(timezone(timedelta(hours=7))).strftime("%Y-%m-%d")
    rows = [offer_row(o, today) for o in offers if o["ref"] not in existing]
    if rows:
        ws.append_rows(rows, value_input_option="RAW")
    print(f"Sheet: added {len(rows)} row(s)")


def save_offers(offers):
    """Google Sheet if it is set up, otherwise the Excel file in the repo."""
    if SHEET_ID and SHEET_KEY:
        save_to_sheet(offers)
    else:
        save_to_excel(offers)


def excel_link():
    """Link for the email: the Google Sheet if set up, otherwise the Excel file in the GitHub repo."""
    if SHEET_ID and SHEET_KEY:
        return f"https://docs.google.com/spreadsheets/d/{SHEET_ID}"
    repo = os.environ.get("GITHUB_REPOSITORY")
    branch = os.environ.get("GITHUB_REF_NAME", "main")
    return f"https://github.com/{repo}/raw/{branch}/{EXCEL_FILE}" if repo else ""


FONT = "font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;"  # no quotes - styles sit inside '...'
NAVY = "#0b2e59"


def facts_box(o):
    """Posted / Valid until / Target market - only the ones we have."""
    e = html.escape
    label = f"style='{FONT}font-size:12px;line-height:18px;color:#6b7686;padding:2px 12px 2px 0;white-space:nowrap;vertical-align:top'"
    value = f"style='{FONT}font-size:13px;line-height:18px;color:#2b2f36;padding:2px 0;vertical-align:top'"
    rows = []
    if o.get("posted"):
        upd = f" <span style='color:#8a94a3'>(updated {e(o['updated'])})</span>" if o.get("updated") else ""
        rows.append(("Posted", e(o["posted"]) + upd))
    if o.get("valid_until"):
        rows.append(("Valid until", e(o["valid_until"])))
    targets = o.get("targets") or []
    if targets:
        shown = ", ".join(targets[:8]) + (f" +{len(targets) - 8} more" if len(targets) > 8 else "")
        rows.append(("Target market", e(shown)))
    if not rows:
        return ""
    body = "".join(f"<tr><td {label}>{k}</td><td {value}>{v}</td></tr>" for k, v in rows)
    return ("<table role='presentation' cellpadding='0' cellspacing='0' "
            "style='background:#f7f9fb;border-radius:6px;margin:0 0 12px;width:100%'>"
            f"<tr><td style='padding:8px 12px'><table role='presentation' cellpadding='0' cellspacing='0'>{body}</table></td></tr></table>")


def offer_card(o):
    e = html.escape
    pills = "".join(
        f"<span style='display:inline-block;background:#eef2f7;color:#3d4b5c;font-size:12px;"
        f"line-height:16px;padding:3px 9px;border-radius:10px;margin:0 4px 4px 0'>{e(keyword_name(k))}</span>"
        for k in o["market_keywords"][:4]
    )
    if len(o["market_keywords"]) > 4:
        pills += f"<span style='font-size:12px;color:#8a94a3'>+{len(o['market_keywords']) - 4} more</span>"
    return (
        "<table role='presentation' width='100%' cellpadding='0' cellspacing='0' "
        "style='background:#ffffff;border:1px solid #e3e7ed;border-radius:8px;margin:0 0 12px'>"
        "<tr><td style='padding:18px 20px'>"
        f"<a href='{e(o['link'])}' style='{FONT}font-size:16px;line-height:22px;font-weight:600;"
        f"color:{NAVY};text-decoration:none'>{e(o['title'])}</a>"
        f"<div style='{FONT}font-size:13px;line-height:18px;color:#6b7686;margin:6px 0 10px'>"
        f"{e(o['location'])} &nbsp;|&nbsp; {e(', '.join(o['partnership']))}</div>"
        f"{facts_box(o)}"
        f"<div style='{FONT}font-size:14px;line-height:21px;color:#2b2f36;margin:0 0 12px'>{e(short_text(o))}</div>"
        f"<div style='{FONT}margin:0 0 12px'>{pills}</div>"
        f"<a href='{e(o['link'])}' style='{FONT}display:inline-block;font-size:13px;font-weight:600;"
        f"color:{NAVY};text-decoration:none;border:1px solid {NAVY};border-radius:6px;padding:7px 14px'>"
        "View offer &rarr;</a>"
        f"<span style='{FONT}font-size:11px;color:#a0a8b4;margin-left:10px'>{e(o['ref'])}</span>"
        "</td></tr></table>"
    )


def build_email(offers):
    e = html.escape
    link = excel_link()
    today = datetime.now(timezone(timedelta(hours=7))).strftime("%d %b %Y")
    n = len(offers)
    subject = f"EEN Business Offers - {n} new - {today}"

    # Group by country, biggest groups first
    groups = {}
    for o in offers:
        groups.setdefault(o["location"] or "Other", []).append(o)
    order = sorted(groups, key=lambda c: (-len(groups[c]), c))
    by_country = "&nbsp;&middot; ".join(
        f"{e(c).replace(' ', '&nbsp;')}&nbsp;{len(groups[c])}" for c in order)

    parts = [
        f"<div style='background:#f3f5f8;padding:24px 12px;margin:0'>"
        "<table role='presentation' width='100%' cellpadding='0' cellspacing='0' "
        "style='max-width:640px;margin:0 auto'><tr><td>",
        # Header
        f"<table role='presentation' width='100%' cellpadding='0' cellspacing='0' "
        f"style='background:{NAVY};border-radius:8px;margin:0 0 20px'><tr><td style='padding:22px 24px'>"
        f"<div style='{FONT}font-size:12px;letter-spacing:1px;color:#a9bcd6;text-transform:uppercase'>"
        "EEN Business Offers</div>"
        f"<div style='{FONT}font-size:26px;line-height:32px;font-weight:700;color:#ffffff;margin:4px 0 2px'>"
        f"{n} new offer{'s' if n != 1 else ''}</div>"
        f"<div style='{FONT}font-size:14px;color:#d5deea'>{today}</div>"
        f"<div style='{FONT}font-size:13px;line-height:20px;color:#d5deea;margin-top:12px;"
        f"border-top:1px solid #2c4c78;padding-top:12px'>{by_country}</div>"
        + (f"<a href='{e(link)}' style='{FONT}display:inline-block;margin-top:14px;font-size:13px;"
           f"font-weight:600;color:{NAVY};background:#ffffff;text-decoration:none;border-radius:6px;"
           f"padding:8px 14px'>{'Open sheet' if SHEET_ID and SHEET_KEY else 'Download Excel'} - all offers &rarr;</a>"
           if link else "")
        + "</td></tr></table>",
    ]
    for c in order:
        parts.append(
            f"<div style='{FONT}font-size:12px;font-weight:700;letter-spacing:1px;color:#6b7686;"
            f"text-transform:uppercase;margin:20px 0 8px 2px'>{e(c)} &nbsp;({len(groups[c])})</div>"
        )
        parts.extend(offer_card(o) for o in groups[c])
    parts.append(
        f"<div style='{FONT}font-size:12px;line-height:18px;color:#8a94a3;text-align:center;margin:20px 0 0'>"
        + (f"Full descriptions of all offers: <a href='{e(link)}' style='color:#6b7686'>all offers sheet</a><br>"
           if link else "")
        + "Source: Enterprise Europe Network - een.ec.europa.eu/partnering-opportunities</div>"
        "</td></tr></table></div>"
    )
    return subject, "".join(parts)


def build_text(offers):
    lines = []
    for o in offers:
        lines += [o["title"], f"{o['location']} | {', '.join(o['partnership'])}",
                  short_text(o), o["link"], ""]
    return "\n".join(lines)


def send_email(subject, body, text=""):
    user = os.environ.get("GMAIL_USER")
    pwd = os.environ.get("GMAIL_APP_PASSWORD")
    to = [x.strip() for x in os.environ.get("EEN_RECIPIENTS", user or "").split(",") if x.strip()]
    if not user or not pwd:
        with open("digest_preview.html", "w", encoding="utf-8") as f:
            f.write(body)
        print("No Gmail settings - saved digest_preview.html instead of sending")
        return
    msg = MIMEMultipart("mixed")  # "mixed" so the Excel attachment shows in every mail app
    msg["Subject"] = subject
    msg["From"] = user
    msg["To"] = ", ".join(to)
    content = MIMEMultipart("alternative")
    content.attach(MIMEText(text or subject, "plain", "utf-8"))
    content.attach(MIMEText(body, "html", "utf-8"))
    msg.attach(content)  # no attachment - the email links to the Excel file in the repo instead
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
        s.login(user, pwd)
        s.sendmail(user, to, msg.as_string())
    print(f"Email sent to {', '.join(to)}")


def main():
    seen = load_seen()
    first_run = seen is None
    seen = seen or {}

    cards = collect_cards()
    offers_cards = [c for c in cards if c["type"].lower() == WANTED_TYPE]
    new_cards = [c for c in offers_cards if c["ref"] not in seen]
    if first_run:
        print("First run: emailing only offers from the last 24 hours")
        new_cards = [c for c in new_cards if c["age_days"] <= 1]
    print(f"Found {len(cards)} profiles, {len(offers_cards)} Business Offers, {len(new_cards)} new")

    offers = []
    for c in new_cards:
        print(f"  details: {c['ref']} {c['title'][:70]}")
        try:
            d = parse_detail(get(c["link"]))
        except Exception as e:
            print(f"  ! skipped for now (will retry next run): {e}")
            continue
        offers.append(add_dates({**c, **{k: v for k, v in d.items() if v}, "ref": c["ref"], "link": c["link"]}))
        time.sleep(DELAY)

    if offers:
        save_offers(offers)
        subject, body = build_email(offers)
        send_email(subject, body, build_text(offers))
    else:
        print("No new Business Offers today - no email sent")

    # Mark as seen only after the email went out
    stamp = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    for o in offers:
        seen[o["ref"]] = stamp
    if first_run:  # baseline: everything currently on the feed
        for c in offers_cards:
            seen.setdefault(c["ref"], stamp)
    save_seen(seen)
    print(f"Saved {len(seen)} seen references")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"ERROR: {e}")
        sys.exit(1)
