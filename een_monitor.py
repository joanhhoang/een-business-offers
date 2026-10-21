"""
EEN Business Offer daily digest
Source: https://een.ec.europa.eu/partnering-opportunities

Each run:
  1. Reads the newest listing pages (stops when posts are older than LOOKBACK_DAYS).
  2. Keeps only "Business Offer" profiles not seen before (by POD reference).
  3. Opens each new offer to get full details.
  4. Emails one digest: Title, Location, Type of partnership, Market keywords,
     Description, Link.
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

BASE = "https://een.ec.europa.eu"
LIST_URL = BASE + "/partnering-opportunities"
SEEN_FILE = os.environ.get("SEEN_FILE", "seen_een.json")
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
    }


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


def build_email(offers):
    today = datetime.now(timezone(timedelta(hours=7))).strftime("%d %b %Y")
    subject = f"EEN Business Offers - {len(offers)} new - {today}"
    parts = [
        "<div style='font-family:Arial,sans-serif;font-size:14px;color:#222;max-width:760px'>",
        f"<h2 style='margin:0 0 4px'>EEN Business Offers - {today}</h2>",
        f"<p style='margin:0 0 16px;color:#555'>{len(offers)} new offer(s) since the last email.</p>",
    ]
    for i, o in enumerate(offers, 1):
        desc = html.escape(o["description"]).replace("\n", "<br>")
        parts.append(
            "<div style='border-top:1px solid #ddd;padding:14px 0'>"
            f"<div style='font-size:16px;font-weight:bold;margin-bottom:6px'>{i}. "
            f"<a href='{o['link']}' style='color:#0b5394;text-decoration:none'>{html.escape(o['title'])}</a></div>"
            f"<b>Location:</b> {html.escape(o['location'])}<br>"
            f"<b>Type of partnership:</b> {html.escape(', '.join(o['partnership']))}<br>"
            f"<b>Market keywords:</b> {html.escape('; '.join(o['market_keywords']))}<br>"
            f"<b>Description:</b><br><div style='margin:4px 0 6px'>{desc}</div>"
            f"<b>Link:</b> <a href='{o['link']}'>{o['link']}</a> ({o['ref']})"
            "</div>"
        )
    parts.append("<p style='color:#888;font-size:12px'>Source: Enterprise Europe Network partnering opportunities.</p></div>")
    return subject, "".join(parts)


def send_email(subject, body):
    user = os.environ.get("GMAIL_USER")
    pwd = os.environ.get("GMAIL_APP_PASSWORD")
    to = [x.strip() for x in os.environ.get("EEN_RECIPIENTS", user or "").split(",") if x.strip()]
    if not user or not pwd:
        with open("digest_preview.html", "w", encoding="utf-8") as f:
            f.write(body)
        print("No Gmail settings - saved digest_preview.html instead of sending")
        return
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = user
    msg["To"] = ", ".join(to)
    msg.attach(MIMEText(body, "html", "utf-8"))
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
        offers.append({**c, **{k: v for k, v in d.items() if v}, "ref": c["ref"], "link": c["link"]})
        time.sleep(DELAY)

    if offers:
        subject, body = build_email(offers)
        send_email(subject, body)
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
