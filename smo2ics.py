import requests, hashlib, pytz
from datetime import datetime, date, timedelta
from dateutil import parser as dtparser
from dateutil.relativedelta import relativedelta
from icalendar import Calendar, Event

# ---- Konfiguration ----
BASE_URL = "https://login.schulmanager-online.de/api/calls"
INSTITUTION_ID = 590
BUNDLE_VERSION = "5449e30183"  # bei Bedarf aus DevTools aktualisieren
TZ = pytz.timezone("Europe/Berlin")
MONTHS_PAST = 0          # wie viele Monate rückwärts
MONTHS_AHEAD = 12        # wie viele Monate vorwärts
OUTFILE = "horststrasse.schulmanager.ics"

HEADERS = {
    "Content-Type": "application/json",
    "Accept": "application/json",
    "Origin": "https://login.schulmanager-online.de",
    "Referer": "https://login.schulmanager-online.de/",
}

def fetch_span(d_from: date, d_to: date):
    payload = {
        "bundleVersion": BUNDLE_VERSION,
        "requests": [{
            "moduleName": "calendar",
            "endpointName": "get-public-events",
            "parameters": {
                "institutionId": INSTITUTION_ID,
                "start": d_from.isoformat(),
                "end": d_to.isoformat(),
                "includeHolidays": True
            }
        }]
    }
    r = requests.post(BASE_URL, headers=HEADERS, json=payload, timeout=20)
    r.raise_for_status()
    j = r.json()
    data = j.get("results", [{}])[0].get("data", {})
    events = (data.get("nonRecurringEvents") or []) + (data.get("recurringEvents") or [])
    return events

def month_windows(start: date, months_back: int, months_fwd: int):
    first = (start.replace(day=1) - relativedelta(months=months_back))
    cursor = first
    for _ in range(months_back + months_fwd + 1):
        start_d = cursor
        end_d = (cursor + relativedelta(months=1)) - timedelta(days=1)
        yield (start_d, end_d)
        cursor = cursor + relativedelta(months=1)

def to_dt(x):
    if x is None:
        return None
    return dtparser.parse(x)

# ---- Events holen ----
seen = set()
cal = Calendar()
cal.add("prodid", "-//Schulmanager Scraper//DE")
cal.add("version", "2.0")
cal.add("X-WR-TIMEZONE", "Europe/Berlin")

from datetime import timezone

build_utc = datetime.now(timezone.utc)
build_local = build_utc.astimezone(TZ)

cal.add("X-WR-CALNAME", "Horststraße Schulkalender")
cal.add("X-WR-CALDESC", f"Stand: {build_local:%Y-%m-%d %H:%M} {build_local.tzname()} (wöchentlich)")
# optionaler Hinweis für Clients:
cal.add("X-PUBLISHED-TTL", "P7D")   # 7 Tage


today = date.today()
for (win_start, win_end) in month_windows(today, MONTHS_PAST, MONTHS_AHEAD):
    evs = fetch_span(win_start, win_end)
    for e in evs:
        title = e.get("summary") or "Ohne Titel"
        start_s = e.get("start")
        end_s = e.get("end")
        all_day = bool(e.get("allDay", False))
        location = e.get("location")
        desc = e.get("description")
        eid = str(e.get("id") or "")

        if not start_s:
            continue

        dtstart = to_dt(start_s)
        dtend = to_dt(end_s) if end_s else (dtstart + timedelta(hours=1))

        # stabile UID (Event-ID + Start + End)
        uid_src = f"{eid}|{dtstart.isoformat()}|{dtend.isoformat()}"
        uid = hashlib.md5(uid_src.encode()).hexdigest() + "@smo"
        if uid in seen:
            continue
        seen.add(uid)

        ve = Event()
        ve.add("uid", uid)
        ve.add("summary", title)

        if all_day:
            # RFC 5545: DTEND ist EXKLUSIV. Schulmanager liefert für allDay bereits exklusive Grenzen.
            ve.add("dtstart", dtstart.date())
            if end_s:
                ve.add("dtend", dtend.date())   # exklusive Grenze direkt übernehmen
            else:
                ve.add("dtend", (dtstart + timedelta(days=1)).date())  # 1-Tages-Event
        else:
            if dtstart.tzinfo is None:
                dtstart = TZ.localize(dtstart)
            if dtend.tzinfo is None:
                dtend = TZ.localize(dtend)
            ve.add("dtstart", dtstart)
            ve.add("dtend", dtend)

        if location:
            ve.add("location", location)
        if desc:
            ve.add("description", desc)

        cal.add_component(ve)

with open(OUTFILE, "wb") as f:
    f.write(cal.to_ical())

print("Geschrieben:", OUTFILE)


