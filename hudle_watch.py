#!/usr/bin/env python3
"""
hudle_watch.py — Noida pickleball availability checker for Hudle.

Reads the real Hudle booking grid and writes an Excel workbook showing which
slots are Not Booked / Booked / Filling Fast / Not Available, plus what CHANGED
since last time — so a cancellation you could grab is obvious.

Results ACCUMULATE. Every run folds into output/history.json, and the workbook
is rendered from the whole store. Checking today never erases what you already
know about Saturday; it just refreshes today. Dates in the past drop off.

Speed: Hudle shows 4 days per screen. Ask for 1-4 days (or one specific date in
that range) and nothing has to page forward — that is the fast, reliable path.

How it works
------------
It drives a headless browser the same way you would: open the venue page, pick
the sport, pick the court, read the grid. No login, no booking, no payment, no
API keys — only the public page any visitor sees.

Setup (once)
------------
    pip install playwright openpyxl
    python -m playwright install chromium

Run
---
    python hudle_watch.py                       # today, every venue (default)
    python hudle_watch.py --date 14             # just the 14th
    python hudle_watch.py --date 14 --time "7:00 PM"    # one slot on the 14th
    python hudle_watch.py --days 4              # today + next 3, still no paging
    python hudle_watch.py --only "sky-nine,pickle-pros"  # a few venues
    python hudle_watch.py --headed              # watch it work

Output lands in ./output/ :
    latest.xlsx                   always the newest workbook
    courts_YYYY-MM-DD_HHMM.xlsx   timestamped copy
    history.json                  the accumulated store — do not delete
    last_run.txt                  plain-text summary
"""

import argparse
import asyncio
import datetime as dt
import json
import os
import re
import sys
import traceback

OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")

# ---------------------------------------------------------------------------
# YOUR VENUES — edit this list freely. Order does not matter.
# ---------------------------------------------------------------------------
VENUE_URLS = [
    "https://hudle.in/venues/arena-44/800489",
    "https://hudle.in/venues/playall-bsic-noida-46/818792",
    "https://hudle.in/venues/zyng-gallant-sports-arena/316748",
    "https://hudle.in/venues/downtown-pickleball-sector-37/912373",
    "https://hudle.in/venues/noida-pickleball-courts/351430",
    "https://hudle.in/venues/playall-box-cricket-arena-sector-105-noida/570479",
    "https://hudle.in/venues/spodium-arena/607735",
    "https://hudle.in/venues/crickdome/421878",
    "https://hudle.in/venues/gd-goenka-sports-arena/299781",
    "https://hudle.in/venues/pickle-pros-noida/902747",
    "https://hudle.in/venues/noida-indoor-stadium/889915",
    "https://hudle.in/venues/anandmay-sports-club-sector-128/334093",
    "https://hudle.in/venues/playall-badminton-arena/840648",
    "https://hudle.in/venues/dinking-zone-one29/423071",
    "https://hudle.in/venues/play-all-modern-school-noida/930734",
    "https://hudle.in/venues/spuddy-multisports-academy-noida/235636",
    "https://hudle.in/venues/coplay-rise-sec-73-noida/874313",
    "https://hudle.in/venues/pickit-vijay-sri-academy/405616",
    "https://hudle.in/venues/pickleflow-social/597820",
    "https://hudle.in/venues/sky-nine-rackets/929811",
    "https://hudle.in/venues/pickleplay-121/588355",
    "https://hudle.in/venues/rally-n-roast-sector-141/479971",
    "https://hudle.in/venues/pitch-please-box-cricket-football/116864",
    "https://hudle.in/venues/the-box-cage/628181",
    "https://hudle.in/venues/heat-sports-arena-150/803998",
    "https://hudle.in/venues/playlife/397418",
]

# Which activities to check. Anything whose Hudle activity name contains one of
# these words (case-insensitive) gets scanned. Pickleball is the focus; add
# "padel" here (or pass --sports pickleball,padel) if you ever want it too.
DEFAULT_SPORTS = ["pickleball"]

# Equipment-rental entries that masquerade as facilities — skipped.
RENTAL_RE = re.compile(r"\b(ball|racquet|racket)\b", re.I)

# ---------------------------------------------------------------------------
# In-page helpers. These selectors were verified against Hudle's live venue
# pages. If Hudle redesigns and a run returns 0 courts, these are what to check.
# ---------------------------------------------------------------------------

JS_VENUE_META = """
() => {
  try {
    const el = document.getElementById('__NEXT_DATA__');
    const v = JSON.parse(el.textContent).props.pageProps.venueDetails;
    return {
      ok: true,
      name: v.name,
      activities: (v.activities || []).map(a => ({
        name: a.name,
        facilities: (a.facilities || []).map(f => ({
          name: f.name,
          slot_length: f.slot_length,
          price: f.price,
          start_time: f.start_time,
          end_time: f.end_time
        }))
      }))
    };
  } catch (e) { return { ok: false, err: String(e) }; }
}
"""

JS_LIST_KIND = """
() => {
  const bodies = [...document.querySelectorAll('.card-body')];
  const sec = bodies.find(c => c.querySelector('[class*="style_bookingCard__"]'));
  if (!sec) return null;
  const head = sec.innerText.slice(0, 60);
  if (/Choose an Activity/i.test(head)) return 'activity';
  if (/Choose a Facility/i.test(head)) return 'facility';
  return 'unknown';
}
"""

JS_CARD_LABELS = """
() => [...document.querySelectorAll('[class*="style_bookingCard__"]')]
        .map(c => (c.innerText.split('\\n').map(s => s.trim()).filter(Boolean)[0] || '?'))
"""

JS_SELECTED = """
() => {
  // Reads back what the wizard says is currently selected, e.g.
  //   "2  Selected a Facility  Court 3  price Rs 225 onwards  Change"
  const out = { activity: null, facility: null };
  for (const c of document.querySelectorAll('.card-body')) {
    const lines = c.innerText.split('\\n').map(s => s.trim()).filter(Boolean);
    for (let i = 0; i < lines.length; i++) {
      if (/^Selected an Activity$/i.test(lines[i]) && lines[i + 1]) out.activity = lines[i + 1];
      if (/^Selected a Facility$/i.test(lines[i]) && lines[i + 1]) out.facility = lines[i + 1];
    }
  }
  return out;
}
"""

JS_READ_GRID = """
() => {
  const tbl = document.querySelector('table[class*="style_table__"]') || document.querySelector('table');
  if (!tbl || tbl.rows.length < 2) return null;
  const head = [...tbl.rows[0].cells].slice(1)
      .map(c => c.innerText.replace(/\\s+/g, ' ').trim());
  const cells = [];
  for (const r of [...tbl.rows].slice(1)) {
    const cs = [...r.cells];
    const time = cs[0].innerText.replace(/\\s+/g, ' ').trim();
    if (!time) continue;
    cs.slice(1).forEach((c, i) => {
      const si = c.querySelector('[class*="slot-item"]');
      if (!si) { cells.push({ day: head[i] || ('c' + i), time, raw: '', text: '' }); return; }
      cells.push({
        day: head[i] || ('c' + i),
        time,
        raw: si.className.toString(),
        text: si.innerText.replace(/\\n/g, ' ').trim()
      });
    });
  }
  return { head, cells };
}
"""


def normalise_status(raw_class: str, text: str) -> str:
    """Map Hudle's slot CSS class onto the four states its own legend shows."""
    blob = f"{raw_class} {text}".lower()
    if not raw_class.strip():
        return "—"                      # past time, or court not offered then
    if "not" in blob and "avail" in blob:
        return "Not available"
    if "book" in blob or "sold" in blob:
        return "Booked"
    if "fill" in blob or "fast" in blob:
        return "Filling fast"
    if "avail" in blob or "left" in blob or "₹" in text:
        return "Available"
    return raw_class.replace("slot-item", "").strip() or "—"


def parse_price(text: str):
    m = re.search(r"₹\s?([\d,]+)", text or "")
    return int(m.group(1).replace(",", "")) if m else None


def parse_seats(text: str):
    m = re.search(r"(\d+)\s*left", text or "")
    return int(m.group(1)) if m else None


def parse_day(day_label: str, today: dt.date):
    """'13 Sun' / 'Today 13 Sun' -> a real date, resolving month rollover."""
    m = re.search(r"\b(\d{1,2})\b", day_label or "")
    if not m:
        return None
    dom = int(m.group(1))
    for delta in range(0, 40):
        cand = today + dt.timedelta(days=delta)
        if cand.day == dom:
            return cand
    return None


# ---------------------------------------------------------------------------
# Scraping
# ---------------------------------------------------------------------------

async def click_book_card(page, index: int) -> bool:
    cards = page.locator('[class*="style_bookingCard__"]')
    if await cards.count() <= index:
        return False
    btn = cards.nth(index).locator("button", has_text=re.compile(r"^book$", re.I))
    if await btn.count() == 0:
        return False
    await btn.first.click()
    await page.wait_for_timeout(1100)
    return True


async def ensure_list(page, kind: str, tries: int = 5) -> bool:
    """Get the wizard showing the activity list or the facility list."""
    for _ in range(tries):
        if await page.evaluate(JS_LIST_KIND) == kind:
            return True
        changes = page.locator("button", has_text=re.compile(r"^change$", re.I))
        n = await changes.count()
        idx = 0 if kind == "activity" else 1
        if n <= idx:
            return await page.evaluate(JS_LIST_KIND) == kind
        await changes.nth(idx).click()
        await page.wait_for_timeout(900)
    return await page.evaluate(JS_LIST_KIND) == kind


def head_dates(grid, today: dt.date):
    """The real dates shown in a grid's column headers, in order."""
    out = []
    for h in grid.get("head", []):
        d = parse_day(h, today)
        if d:
            out.append(d)
    return out


async def wait_for_window(page, after: dt.date, today: dt.date, seconds: float = 16.0):
    """Wait until the grid shows a date window starting strictly after `after`.

    Comparing real dates — not the header string — is what makes paging
    reliable. A string compare can be satisfied by a half-rendered table, which
    is how later courts silently ended up with only the first 4 days.
    """
    deadline = seconds / 0.35
    attempt = 0
    while attempt < deadline:
        attempt += 1
        await page.wait_for_timeout(350)
        g = await page.evaluate(JS_READ_GRID)
        if not g or not g["cells"]:
            continue
        dates = head_dates(g, today)
        if not dates:
            continue
        if after is None or min(dates) > after:
            return g, min(dates)
    return None, None


async def read_facility(page, wanted, today: dt.date):
    """Read the slot grid for the currently-selected facility.

    `wanted` is the set of dates we actually care about. Hudle shows 4 days per
    window, so when every wanted date is inside the first window this returns
    without ever touching the (unreliable) next arrow.

    Returns (slots, dates_covered, reason).
    """
    wanted = set(wanted)
    last_wanted = max(wanted)
    # The page loads showing Today, so only reset when we may have paged away.
    if len(wanted) > 4 or max(wanted) - today >= dt.timedelta(days=4):
        today_btn = page.locator('button[class*="style_today__"]')
        if await today_btn.count():
            try:
                await today_btn.first.click()
                await page.wait_for_timeout(700)
            except Exception:
                pass

    collected, seen = [], set()
    last_first = None
    reason = "complete"

    window_span = None
    for _window in range(6):
        grid, first = await wait_for_window(page, last_first, today)
        if grid is None:
            reason = "grid did not render"
            break

        # If the window jumped further than one step, dates were skipped and
        # this court's data has a hole in it. Say so rather than hide it.
        if last_first is not None:
            step = (first - last_first).days
            if window_span is None:
                window_span = step
            elif step > window_span:
                reason = "skipped a window"
        last_first = first

        for c in grid["cells"]:
            key = (c["day"], c["time"])
            if key in seen:
                continue
            seen.add(key)
            date = parse_day(c["day"], today)
            if date is None or date not in wanted:
                continue
            collected.append({
                "date": date.isoformat(),
                "dow": date.strftime("%a"),
                "time": c["time"],
                "status": normalise_status(c["raw"], c["text"]),
                "price": parse_price(c["text"]),
                "seats_left": parse_seats(c["text"]),
            })

        covered = {dt.date.fromisoformat(c["date"]) for c in collected}
        if wanted <= covered:
            break                                   # got everything asked for
        # Paged past the last date we want? Nothing more worth fetching.
        if last_first and last_first > last_wanted:
            reason = "past horizon"
            break

        nxt = page.locator('button[class*="style_next__"]')
        if not await nxt.count():
            reason = "no next button"
            break

        # Is Hudle actually offering more dates for THIS court? A disabled
        # arrow means the venue only publishes this far ahead on this court —
        # that is a fact about the venue, not a failure to scrape.
        try:
            if await nxt.first.is_disabled():
                reason = "venue limit"
                break
        except Exception:
            pass
        try:
            cls = (await nxt.first.get_attribute("class")) or ""
            aria = (await nxt.first.get_attribute("aria-disabled")) or ""
            if "disabled" in cls.lower() or aria.lower() == "true":
                reason = "venue limit"
                break
        except Exception:
            pass

        # Advance ONE window. The header before the click is the reference:
        # if it changes at all, the click landed and we must NOT click again —
        # a second click skips a whole window and loses those days for good.
        before = await page.evaluate(JS_READ_GRID)
        before_head = ",".join(before["head"]) if before else ""

        advanced = False
        for attempt in range(2):
            try:
                await nxt.first.click()
            except Exception:
                await page.wait_for_timeout(500)
                continue

            # Generous wait: a slow render is not a failed click.
            for _ in range(40):                     # up to ~14s
                await page.wait_for_timeout(350)
                g = await page.evaluate(JS_READ_GRID)
                if g and g["cells"] and ",".join(g["head"]) != before_head:
                    advanced = True
                    break
            if advanced:
                break

        if not advanced:
            reason = "next click had no effect"
            break

    return collected, len({c["date"] for c in collected}), reason


def same_name(a: str, b: str) -> bool:
    norm = lambda s: re.sub(r"[^a-z0-9]", "", (s or "").lower())
    return norm(a) == norm(b)


async def select_court(page, url, act_idx, fac_idx, want_court, log, fresh=True):
    """Select one court and CONFIRM the wizard really landed on it.

    Without this check a mis-fired click silently leaves the previous court
    selected, and the grid would be read twice under two different names —
    which looks like 'every court has identical availability'.
    """
    for attempt in (1, 2):
        # Reload only when asked. Staying on the page is far faster, and is safe
        # whenever no date paging is involved (every wanted date is in the first
        # window) — there is no stale date window to inherit. The caller asks for
        # a reload for the first court of a venue, and after any failure.
        # ALWAYS reload. Switching courts in-page leaves the PREVIOUS court's
        # slot table in the DOM; the new court's name is confirmed, the grid is
        # not, and the old numbers get filed under the new court. A reload wipes
        # the table so there is nothing stale to misread. `fresh` is kept for
        # callers but is no longer an opt-out.
        await page.goto(url, wait_until="domcontentloaded", timeout=60000)
        await page.wait_for_timeout(700 if attempt == 1 else 1600)

        if await ensure_list(page, "activity"):
            await click_book_card(page, act_idx)
        await page.wait_for_timeout(350)

        # A venue with a single court skips the facility list entirely — Hudle
        # selects the only court for you. Take the win rather than demanding a
        # list that will never appear.
        sel = await page.evaluate(JS_SELECTED)
        if sel.get("facility") and same_name(sel["facility"], want_court):
            return True, sel["facility"]

        if await ensure_list(page, "facility"):
            await click_book_card(page, fac_idx)
            await page.wait_for_timeout(350)

        sel = await page.evaluate(JS_SELECTED)
        got = sel.get("facility")
        if got and same_name(got, want_court):
            return True, got
        if got and attempt == 1:
            log(f"      (wizard showed {got!r} when {want_court!r} was asked for; retrying)")

    sel = await page.evaluate(JS_SELECTED)
    return False, sel.get("facility")


async def scrape_venue(page, url: str, sports, wanted, today: dt.date, log):
    await page.goto(url, wait_until="domcontentloaded", timeout=60000)
    await page.wait_for_timeout(700)

    meta = await page.evaluate(JS_VENUE_META)
    if not meta.get("ok"):
        log(f"    ! could not read venue data: {meta.get('err')}")
        return None

    vname = meta["name"]
    matching_acts = [
        (i, a) for i, a in enumerate(meta["activities"])
        if any(s.lower() in a["name"].lower() for s in sports)
    ]
    if not matching_acts:
        log(f"    {vname}: no matching sport, skipped")
        return {"venue": vname, "url": url, "courts": []}

    out = {"venue": vname, "url": url, "courts": [], "roster": []}
    fresh_needed = True          # first court of a venue always reloads

    for act_idx, act in matching_acts:
        facilities = [f for f in act["facilities"] if not RENTAL_RE.search(f["name"])]
        if not facilities:
            continue
        for fac_idx, fac in enumerate(act["facilities"]):
            if RENTAL_RE.search(fac["name"]):
                continue
            # Record every court we INTEND to read. A court that fails must still
            # show up as a column, otherwise it silently disappears and the sheet
            # looks complete when it isn't.
            if fac["name"] not in out["roster"]:
                out["roster"].append(fac["name"])
            try:
                slots, covered, reason, landed, ok = [], 0, "not attempted", None, False

                # Up to 3 goes. The first reuses the open page when the previous
                # court worked (fast path); every retry forces a full reload,
                # because an empty grid is almost always a render race.
                for attempt in range(3):
                    want_fresh = fresh_needed or attempt > 0
                    ok, landed = await select_court(page, url, act_idx, fac_idx,
                                                    fac["name"], log, fresh=want_fresh)
                    if not ok:
                        fresh_needed = True
                        continue
                    slots, covered, reason = await read_facility(page, wanted, today)
                    if covered >= len(wanted) or reason in LIMIT_REASONS:
                        fresh_needed = False
                        break
                    fresh_needed = True          # something went wrong; reload next
                    if attempt < 2:
                        log(f"      (retrying {fac['name']} — {reason})")

                if not ok:
                    log(f"    ! {vname} / {fac['name']}: could not confirm this court "
                        f"was selected (wizard showed {landed!r}) — SKIPPED rather than "
                        f"reporting another court's slots under this name")
                    fresh_needed = True
                    continue

                avail = sum(1 for s in slots if s["status"] == "Available")
                if covered >= len(wanted):
                    short = ""
                elif reason in ("venue limit", "no next button"):
                    short = f"  (venue only opens bookings {covered} days ahead here)"
                else:
                    short = f"  << only {covered}/{len(wanted)} dates [{reason}]"
                if not slots:
                    log(f"    ! {vname} / {fac['name']}: no slot grid returned "
                        f"(venue may be closed for the week, or the page did not load)")
                else:
                    log(f"    {vname} / {act['name']} / {fac['name']}: "
                        f"{len(slots)} slots, {avail} available{short}")

                out["courts"].append({
                    "sport": act["name"],
                    "court": fac["name"],
                    "slot_length": fac.get("slot_length"),
                    "list_price": fac.get("price"),
                    "days_covered": covered,
                    "coverage_reason": reason,
                    "slots": slots,
                })
            except Exception as e:
                fresh_needed = True
                log(f"    ! {vname} / {fac['name']}: {e}")
                try:
                    await page.goto(url, wait_until="domcontentloaded", timeout=60000)
                    await page.wait_for_timeout(900)
                except Exception:
                    pass

    # Sanity check: identical courts at a quiet venue are normal, but if EVERY
    # court matches exactly it is worth saying so out loud.
    if len(out["courts"]) > 1:
        prints = {}
        for c in out["courts"]:
            sig = "|".join(f"{s['date']}{s['time']}{s['status']}" for s in c["slots"])
            prints.setdefault(sig, []).append(c["court"])
        if len(prints) == 1:
            log(f"    note: all {len(out['courts'])} courts at {vname} have identical "
                f"availability. Normal at a quiet venue — worth a spot-check on the "
                f"site if it persists when the venue is busy.")
    return out


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------

LIMIT_REASONS = ("venue limit", "no next button")

# Plain words, as Vijay's example workbook uses them.
WORD = {
    "Available": "Not Booked",
    "Filling fast": "Filling Fast",
    "Booked": "Booked",
    "Not available": "Not Available",
    "—": "",
}


def parse_dates_arg(spec: str, days: int, today: dt.date):
    """'--date 14,2026-09-20' or '--days 3' -> the concrete dates to fetch."""
    if not spec.strip():
        return [today + dt.timedelta(days=i) for i in range(max(1, days))]
    out = []
    for raw in spec.split(","):
        raw = raw.strip()
        if not raw:
            continue
        try:
            out.append(dt.date.fromisoformat(raw))
            continue
        except ValueError:
            pass
        if raw.isdigit():                            # bare day-of-month
            dom = int(raw)
            for delta in range(0, 62):
                cand = today + dt.timedelta(days=delta)
                if cand.day == dom:
                    out.append(cand)
                    break
    return sorted(set(out)) or [today]


def parse_times_arg(spec: str):
    """'7:00 PM, 19:30' -> the grid's own '07:00 PM' label format."""
    out = []
    for raw in spec.split(","):
        raw = raw.strip()
        if not raw:
            continue
        for fmt in ("%I:%M %p", "%I%p", "%I %p", "%H:%M", "%I:%M%p"):
            try:
                t = dt.datetime.strptime(raw.upper().replace(".", ""), fmt)
                out.append(t.strftime("%I:%M %p"))
                break
            except ValueError:
                continue
    return out


def _time_key(t: str):
    """Sort '06:00 AM' style labels chronologically."""
    try:
        return dt.datetime.strptime(t.strip().upper(), "%I:%M %p")
    except ValueError:
        return dt.datetime.min


def time_range(start_label: str, minutes: int) -> str:
    """'06:00 AM' + 30  ->  '6:00 am to 6:30 am'."""
    try:
        t0 = dt.datetime.strptime(start_label.strip().upper(), "%I:%M %p")
    except ValueError:
        return start_label
    t1 = t0 + dt.timedelta(minutes=minutes or 30)
    fmt = lambda d: f"{d.hour % 12 or 12}:{d.minute:02d} {'am' if d.hour < 12 else 'pm'}"
    return f"{fmt(t0)} to {fmt(t1)}"


def _coverage_line(data, days):
    """Separate 'the venue does not open bookings that far' from a real failure."""
    limited, failed = [], []
    for v in data:
        for c in v["courts"]:
            cov = c.get("days_covered", days)
            if cov >= days:
                continue
            entry = f"{v['venue']} / {c['court']} ({cov}d)"
            (limited if c.get("coverage_reason") in LIMIT_REASONS else failed).append(entry)

    if not limited and not failed:
        return f"Coverage: complete — every court returned all {days} days."

    parts = []
    if limited:
        head = "; ".join(limited[:5]) + (f"; +{len(limited)-5} more" if len(limited) > 5 else "")
        parts.append(f"{len(limited)} court(s) only open bookings a few days ahead, "
                     f"so a shorter window is all there is: {head}.")
    if failed:
        head = "; ".join(failed[:5]) + (f"; +{len(failed)-5} more" if len(failed) > 5 else "")
        parts.append(f"{len(failed)} court(s) FAILED to return the full window "
                     f"and may be missing dates: {head}.")
    return "Coverage: " + " ".join(parts)


def write_workbook(data, changes, path, today, days, sports, time_filter=None):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    F = "Arial"
    hdr_fill = PatternFill("solid", fgColor="1F3864")
    hdr = Font(name=F, bold=True, color="FFFFFF", size=10)
    body = Font(name=F, size=10)
    bold = Font(name=F, size=10, bold=True)
    title = Font(name=F, size=14, bold=True, color="1F3864")
    note = Font(name=F, size=9, italic=True, color="666666")
    thin = Side(style="thin", color="D9D9D9")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)

    FILLS = {
        "Available":    PatternFill("solid", fgColor="C6EFCE"),
        "Filling fast": PatternFill("solid", fgColor="FFEB9C"),
        "Booked":       PatternFill("solid", fgColor="F8CBCB"),
        "Not available": PatternFill("solid", fgColor="E8E8E8"),
        "—":            PatternFill("solid", fgColor="FFFFFF"),
    }
    SHORT = {"Available": "OPEN", "Filling fast": "FAST",
             "Booked": "BOOKED", "Not available": "n/a", "—": ""}

    wb = Workbook()

    # ---- Run Info -----------------------------------------------------------
    ws = wb.active
    ws.title = "Run Info"
    ws.column_dimensions["A"].width = 100
    n_courts = sum(len(v["courts"]) for v in data)
    n_slots = sum(len(c["slots"]) for v in data for c in v["courts"])
    n_open = sum(1 for v in data for c in v["courts"]
                 for s in c["slots"] if s["status"] == "Available")
    info = [
        ("Hudle court availability — Noida", title),
        ("", body),
        (f"Checked: {dt.datetime.now():%a %d %b %Y, %H:%M}", bold),
        (f"Horizon: {today:%d %b} to {today + dt.timedelta(days=days-1):%d %b} ({days} days)", body),
        (f"Sports: {', '.join(sports)}", body),
        (f"Venues read: {len(data)}   Courts read: {n_courts}   Slots read: {n_slots:,}", body),
        (f"Slots currently OPEN: {n_open:,}", bold),
        ("", body),
        (_coverage_line(data, days), note),
        ("", body),
        ("Sheets", bold),
        ("Summary     — open slots per venue per day. Start here to pick where to play.", body),
        ("Open Slots  — every bookable slot, soonest first. This is the one to act on.", body),
        ("Changes     — what opened up or got taken since the previous run.", body),
        ("One sheet per venue — Date, Day, Slot Time, then a column per court.", body),
        ("", body),
        ("Colour key", bold),
        ("green Not Booked = free   ·   amber Filling Fast   ·   red Booked   ·   grey Not Available", body),
        ("Blank means the slot is in the past or the court is closed then.", body),
        ("", body),
        ("A 30-minute court needs two consecutive OPEN slots for an hour's game.", note),
        ("Read directly from hudle.in public venue pages. No login, no booking, no payment.", note),
    ]
    for i, (t, f) in enumerate(info, start=1):
        c = ws.cell(row=i, column=1, value=t)
        c.font = f

    # ---- Open Slots ---------------------------------------------------------
    ws = wb.create_sheet("Open Slots")
    heads = ["Date", "Day", "Time", "Venue", "Sport", "Court", "Slot", "Price", "Seats left"]
    widths = [12, 7, 11, 32, 22, 24, 8, 10, 11]
    for i, (h, w) in enumerate(zip(heads, widths), start=1):
        c = ws.cell(row=1, column=i, value=h)
        c.fill = hdr_fill; c.font = hdr; c.border = box
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    open_rows = []
    for v in data:
        for c in v["courts"]:
            for s in c["slots"]:
                if s["status"] in ("Available", "Filling fast"):
                    open_rows.append((s["date"], s["dow"], s["time"], v["venue"],
                                      c["sport"], c["court"],
                                      f"{c.get('slot_length') or ''} min",
                                      s["price"] or c.get("list_price"),
                                      s["seats_left"], s["status"]))
    open_rows.sort(key=lambda r: (r[0], r[2], r[3]))
    r = 2
    for row in open_rows:
        for i, val in enumerate(row[:9], start=1):
            c = ws.cell(row=r, column=i, value=val)
            c.font = body; c.border = box
            c.alignment = Alignment(horizontal="center" if i in (1, 2, 3, 7, 8, 9) else "left")
            if i == 8 and isinstance(val, int):
                c.number_format = '₹#,##0'
        ws.cell(row=r, column=1).fill = FILLS.get(row[9], FILLS["—"])
        r += 1
    if r > 2:
        ws.auto_filter.ref = f"A1:I{r-1}"
    else:
        ws.cell(row=2, column=1, value="Nothing open in this window.").font = note

    # ---- Changes ------------------------------------------------------------
    ws = wb.create_sheet("Changes")
    for i, (h, w) in enumerate(zip(
            ["What changed", "Date", "Day", "Time", "Venue", "Sport", "Court", "Was", "Now"],
            [16, 12, 7, 11, 32, 22, 24, 14, 14]), start=1):
        c = ws.cell(row=1, column=i, value=h)
        c.fill = hdr_fill; c.font = hdr; c.border = box
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    r = 2
    if changes is None:
        ws.cell(row=2, column=1, value="First run — nothing to compare against yet.").font = note
    elif not changes:
        ws.cell(row=2, column=1, value="No changes since the previous run.").font = note
    else:
        for ch in changes:
            vals = [ch["kind"], ch["date"], ch["dow"], ch["time"], ch["venue"],
                    ch["sport"], ch["court"], ch["was"], ch["now"]]
            for i, val in enumerate(vals, start=1):
                c = ws.cell(row=r, column=i, value=val)
                c.font = bold if i == 1 else body
                c.border = box
                c.alignment = Alignment(horizontal="center" if i in (1, 2, 3, 4, 8, 9) else "left")
            ws.cell(row=r, column=1).fill = (
                FILLS["Available"] if ch["kind"] == "OPENED UP" else FILLS["Booked"])
            r += 1
        ws.auto_filter.ref = f"A1:I{r-1}"

    # ---- Summary: one row per venue, one column per date --------------------
    ws = wb.create_sheet("Summary")
    all_dates = sorted({s["date"] for v in data for c in v["courts"] for s in c["slots"]})
    ws.cell(row=1, column=1, value="Open slots per venue, per day").font = title
    ws.cell(row=2, column=1,
            value="How many slots are currently Not Booked, counting every court at that "
                  "venue. Bigger number = easier to get a game.").font = note

    hrow = 4
    heads = ["Venue", "Courts"] + [f"{dt.date.fromisoformat(d):%d %b}\n{dt.date.fromisoformat(d):%a}"
                                   for d in all_dates] + ["Week total"]
    for i, h in enumerate(heads, start=1):
        c = ws.cell(row=hrow, column=i, value=h)
        c.fill = hdr_fill; c.font = hdr; c.border = box
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = 34 if i == 1 else (8 if i == 2 else 10)
    ws.row_dimensions[hrow].height = 30
    ws.freeze_panes = ws.cell(row=hrow + 1, column=2)

    rows_summary = []
    for v in data:
        if not v["courts"]:
            continue
        per_day = []
        for d in all_dates:
            per_day.append(sum(1 for c in v["courts"] for s in c["slots"]
                               if s["date"] == d and s["status"] == "Available"))
        rows_summary.append((v["venue"], len(v["courts"]), per_day))
    rows_summary.sort(key=lambda r: -sum(r[2]))

    r = hrow + 1
    for vname, ncourts, per_day in rows_summary:
        ws.cell(row=r, column=1, value=vname).font = bold
        ws.cell(row=r, column=1).border = box
        ws.cell(row=r, column=2, value=ncourts).border = box
        ws.cell(row=r, column=2).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=2).font = body
        for j, n in enumerate(per_day):
            c = ws.cell(row=r, column=3 + j, value=n)
            c.font = body; c.border = box
            c.alignment = Alignment(horizontal="center")
            if n == 0:
                c.fill = FILLS["Booked"]
            elif n < 10:
                c.fill = FILLS["Filling fast"]
            else:
                c.fill = FILLS["Available"]
        first = get_column_letter(3)
        last = get_column_letter(2 + len(all_dates))
        tc = ws.cell(row=r, column=3 + len(all_dates), value=f"=SUM({first}{r}:{last}{r})")
        tc.font = bold; tc.border = box
        tc.alignment = Alignment(horizontal="center")
        r += 1
    ws.cell(row=r + 1, column=1,
            value="Red = nothing open that day · amber = under 10 slots · green = plenty."
            ).font = note

    # ---- Slot Watch: one chosen time, every court, every date ---------------
    if time_filter:
        ws = wb.create_sheet("Slot Watch")
        ws.cell(row=1, column=1,
                value=f"Slot watch — {', '.join(time_filter)}").font = title
        ws.cell(row=2, column=1,
                value="Status of just this slot across every court you track. "
                      "One row per court, one column per date.").font = note

        # One row per court PER TIME — collapsing times onto one row would let a
        # later time overwrite an earlier one and report the wrong status.
        pairs = []
        for v in data:
            for c in v["courts"]:
                for t in sorted({s["time"] for s in c["slots"]}, key=_time_key):
                    pairs.append((v["venue"], c, t))
        wdates = sorted({s["date"] for v in data for c in v["courts"] for s in c["slots"]})
        hrow = 4
        heads = ["Venue", "Court", "Slot Time"] + [f"{dt.date.fromisoformat(d):%d %b}\n"
                                                   f"{dt.date.fromisoformat(d):%a}" for d in wdates]
        widths = [32, 22, 22] + [14] * len(wdates)
        for i, (h, w) in enumerate(zip(heads, widths), start=1):
            c = ws.cell(row=hrow, column=i, value=h)
            c.fill = hdr_fill; c.font = hdr; c.border = box
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.row_dimensions[hrow].height = 30
        ws.freeze_panes = ws.cell(row=hrow + 1, column=3)

        r = hrow + 1
        for vname, crt, tlabel in pairs:
            by_date = {s["date"]: s for s in crt["slots"] if s["time"] == tlabel}
            ws.cell(row=r, column=1, value=vname).font = body
            ws.cell(row=r, column=2, value=crt["court"]).font = body
            ws.cell(row=r, column=3,
                    value=time_range(tlabel, crt.get("slot_length") or 30)).font = body
            for i in (1, 2, 3):
                ws.cell(row=r, column=i).border = box
            for j, d in enumerate(wdates):
                s = by_date.get(d)
                st = s["status"] if s else "—"
                c = ws.cell(row=r, column=4 + j, value=WORD.get(st, st))
                c.font = body; c.border = box
                c.fill = FILLS.get(st, FILLS["—"])
                c.alignment = Alignment(horizontal="center", vertical="center")
            r += 1
        if r > hrow + 1:
            ws.auto_filter.ref = f"A{hrow}:{get_column_letter(3 + len(wdates))}{r-1}"

    # ---- One sheet per venue, laid out court-by-court ------------------------
    used = set()
    for v in data:
        if not v["courts"]:
            continue
        base = re.sub(r"[\\/*?:\[\]]", "-", v["venue"]).strip()[:28].strip()
        name, k = base, 2
        while name.lower() in used:
            name = f"{base[:25]}_{k}"; k += 1
        used.add(name.lower())
        ws = wb.create_sheet(name)

        dates = sorted({s["date"] for c in v["courts"] for s in c["slots"]})
        times = sorted({s["time"] for c in v["courts"] for s in c["slots"]},
                       key=_time_key)
        lookup = {(c["court"], s["date"], s["time"]): s
                  for c in v["courts"] for s in c["slots"]}
        courts = [c["court"] for c in v["courts"]]
        slot_len = {c["court"]: (c.get("slot_length") or 30) for c in v["courts"]}

        ws.cell(row=1, column=1, value=v["venue"]).font = title
        ws.cell(row=2, column=1,
                value=f"{len(courts)} court(s) · "
                      f"{', '.join(sorted({c['sport'] for c in v['courts']}))} · "
                      f"filter the Date column to pick a day").font = note

        hrow = 4
        heads = ["Date", "Day", "Slot Time"] + courts + ["Last checked"]
        widths = [12, 7, 24] + [17] * len(courts) + [14]
        for i, (h, w) in enumerate(zip(heads, widths), start=1):
            c = ws.cell(row=hrow, column=i, value=h)
            c.fill = hdr_fill; c.font = hdr; c.border = box
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.row_dimensions[hrow].height = 30
        ws.freeze_panes = ws.cell(row=hrow + 1, column=4)

        row = hrow + 1
        for d in dates:
            dd = dt.date.fromisoformat(d)
            for t in times:
                cells = [lookup.get((crt, d, t)) for crt in courts]
                # A row where every court is blank is a past time or a closed
                # hour — nothing to decide about, so leave it out.
                if not any(c and c["status"] != "—" for c in cells):
                    continue
                mins = slot_len.get(courts[0], 30)
                ws.cell(row=row, column=1, value=f"{dd:%d %b %Y}").font = body
                ws.cell(row=row, column=2, value=f"{dd:%a}").font = body
                ws.cell(row=row, column=3, value=time_range(t, mins)).font = body
                for i in (1, 2, 3):
                    ws.cell(row=row, column=i).border = box
                    ws.cell(row=row, column=i).alignment = Alignment(horizontal="center")
                checked = ""
                for j, crt in enumerate(courts):
                    s = lookup.get((crt, d, t))
                    st = s["status"] if s else "—"
                    if s and s.get("checked"):
                        checked = s["checked"]
                    c = ws.cell(row=row, column=4 + j, value=WORD.get(st, st))
                    c.font = body
                    c.fill = FILLS.get(st, FILLS["—"])
                    c.border = box
                    c.alignment = Alignment(horizontal="center", vertical="center")
                cc = ws.cell(row=row, column=4 + len(courts), value=checked)
                cc.font = note; cc.border = box
                cc.alignment = Alignment(horizontal="center")
                row += 1
        if row > hrow + 1:
            ws.auto_filter.ref = f"A{hrow}:{get_column_letter(4 + len(courts))}{row-1}"

    wb.save(path)
    return n_open, len(open_rows)


# ---------------------------------------------------------------------------
# Change detection
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# Accumulating store. A run only refreshes what it looked at; everything else
# stays exactly as it was, so checking today never wipes what you know about
# Saturday.
# ---------------------------------------------------------------------------

HISTORY = "history.json"

# How many past days to keep in the archive. Past dates are not bookable, but
# they show how fast each slot fills, which is what tells you when to book.
KEEP_DAYS = 90

# Key separator. It must be something no venue or court name can contain.
# "|" was a bad choice: a dozen venues are named like "Spodium Arena | Sector 29",
# which split into too many pieces and were silently dropped from the workbook.
SEP = "\x1f"


def _migrate_keys(slots):
    """Upgrade old pipe-separated keys without losing their data."""
    if not any("|" in k and SEP not in k for k in slots):
        return slots
    fixed = {}
    for k, v in slots.items():
        if SEP in k:
            fixed[k] = v
            continue
        # date and time are the last two fields, so split from the RIGHT and the
        # venue keeps any pipes it legitimately contains.
        parts = k.rsplit("|", 4)
        fixed[SEP.join(parts) if len(parts) == 5 else k] = v
    return fixed


HIST_DIR = "history"


def _day_files(out_dir):
    d = os.path.join(out_dir, HIST_DIR)
    if not os.path.isdir(d):
        return []
    return sorted(os.path.join(d, f) for f in os.listdir(d) if f.endswith(".json"))


def load_history(out_dir):
    """Read every day-file back into one store.

    One file per day matters: a run only rewrites TODAY, so past days are
    written once and never touched again. That is what makes 90 days of
    history affordable when we commit to git every 30 minutes.
    """
    history = {"slots": {}, "rosters": {}}

    # Legacy single-file store — fold it in once, then it can be deleted.
    legacy = os.path.join(out_dir, HISTORY)
    if os.path.exists(legacy):
        try:
            with open(legacy, encoding="utf-8") as fh:
                old = json.load(fh)
            history["slots"].update(_migrate_keys(old.get("slots", {}) or {}))
            history["rosters"].update(old.get("rosters", {}) or {})
        except Exception:
            pass

    for path in _day_files(out_dir):
        try:
            with open(path, encoding="utf-8") as fh:
                day = json.load(fh)
            history["slots"].update(_migrate_keys(day.get("slots", {}) or {}))
            history["rosters"].update(day.get("rosters", {}) or {})
        except Exception:
            continue
    return history


def save_history(out_dir, history, today):
    """Write each day to its own file and drop days past the retention window."""
    d = os.path.join(out_dir, HIST_DIR)
    os.makedirs(d, exist_ok=True)

    by_day = {}
    for key, rec in history.get("slots", {}).items():
        parts = key.split(SEP)
        if len(parts) != 5:
            continue
        by_day.setdefault(parts[3], {})[key] = rec

    rosters = history.get("rosters", {})
    for date, slots in by_day.items():
        with open(os.path.join(d, f"{date}.json"), "w", encoding="utf-8") as fh:
            json.dump({"slots": slots, "rosters": rosters}, fh)

    cutoff = today - dt.timedelta(days=KEEP_DAYS)
    for path in _day_files(out_dir):
        stem = os.path.splitext(os.path.basename(path))[0]
        try:
            if dt.date.fromisoformat(stem) < cutoff:
                os.remove(path)
        except Exception:
            continue

    # The old single file is now redundant.
    legacy = os.path.join(out_dir, HISTORY)
    if os.path.exists(legacy):
        try:
            os.remove(legacy)
        except Exception:
            pass
    return len(by_day)


def merge_into_history(history, data, today, keep_past_days=KEEP_DAYS):
    """Fold this run's readings into the store, then drop stale past dates."""
    stamp = dt.datetime.now().strftime("%d %b %H:%M")
    slots = history.setdefault("slots", {})
    for v in data:
        for ci, c in enumerate(v["courts"]):
            for s in c["slots"]:
                key = SEP.join([v["venue"], c["sport"], c["court"], s["date"], s["time"]])

                # Once a slot's time has passed, Hudle greys it out and we read
                # it as "—". Do NOT let that erase what we already knew about
                # it. Keep the last real status instead, so the day's record
                # stays intact instead of dissolving into blanks as it ages.
                prev = slots.get(key)
                if s["status"] == "—" and prev and prev.get("status") not in ("—", None):
                    continue

                slots[key] = {
                    "order": ci,
                    "status": s["status"],
                    "price": s.get("price"),
                    "seats_left": s.get("seats_left"),
                    "slot_length": c.get("slot_length") or 30,
                    "url": v.get("url", ""),
                    "checked": stamp,
                }
    rosters = history.setdefault("rosters", {})
    for v in data:
        if v.get("roster"):
            rosters[v["venue"]] = v["roster"]

    cutoff = today - dt.timedelta(days=keep_past_days)
    for key in [k for k in slots if len(k.split(SEP)) == 5]:
        try:
            if dt.date.fromisoformat(key.split(SEP)[3]) < cutoff:
                del slots[key]
        except Exception:
            pass
    return history


def history_to_data(history, rosters=None):
    """Rebuild the venue/court/slot structure from the accumulated store."""
    venues = {}
    for key, rec in history.get("slots", {}).items():
        parts = key.split(SEP)
        if len(parts) != 5:
            continue
        vname, sport, court, date, time = parts
        v = venues.setdefault(vname, {"venue": vname, "url": rec.get("url", ""), "_courts": {}})
        if rec.get("url"):
            v["url"] = rec["url"]
        c = v["_courts"].setdefault((sport, court), {
            "sport": sport, "court": court,
            "slot_length": rec.get("slot_length", 30),
            "list_price": rec.get("price"), "slots": [],
            "days_covered": 0, "coverage_reason": "complete",
            "order": rec.get("order", 999),
        })
        c["slots"].append({
            "date": date,
            "dow": dt.date.fromisoformat(date).strftime("%a"),
            "time": time,
            "status": rec.get("status", "—"),
            "price": rec.get("price"),
            "seats_left": rec.get("seats_left"),
            "checked": rec.get("checked", ""),
        })
    rosters = rosters if rosters is not None else history.get("rosters", {})
    out = []
    for v in venues.values():
        courts = list(v.pop("_courts").values())
        # Re-insert courts the venue has but this run never managed to read.
        have = {c["court"] for c in courts}
        for i, cname in enumerate(rosters.get(v["venue"], [])):
            if cname not in have:
                courts.append({"sport": "", "court": cname, "slot_length": 30,
                               "list_price": None, "slots": [], "days_covered": 0,
                               "coverage_reason": "not read this run",
                               "order": i, "missing": True})
        for c in courts:
            c["days_covered"] = len({s["date"] for s in c["slots"]})
        # Hudle's own order (Outdoor 1-4, then Indoor 1-3) reads far better than
        # alphabetical, which interleaves them.
        v["courts"] = sorted(courts, key=lambda c: (c.get("order", 999), c["court"]))
        out.append(v)
    return sorted(out, key=lambda v: v["venue"])


def flatten(data):
    out = {}
    for v in data:
        for c in v["courts"]:
            for s in c["slots"]:
                out[SEP.join([v["venue"], c["sport"], c["court"], s["date"], s["time"]])] = s["status"]
    return out


def diff(prev, curr, data):
    meta = {}
    for v in data:
        for c in v["courts"]:
            for s in c["slots"]:
                meta[SEP.join([v["venue"], c["sport"], c["court"], s["date"], s["time"]])] = (
                    v["venue"], c["sport"], c["court"], s["date"], s["dow"], s["time"])
    changes = []
    for key, now in curr.items():
        was = prev.get(key)
        if was is None or was == now:
            continue
        opened = now in ("Available", "Filling fast") and was in ("Booked", "Not available")
        taken = was in ("Available", "Filling fast") and now in ("Booked", "Not available")
        if not (opened or taken):
            continue
        venue, sport, court, date, dow, time = meta[key]
        changes.append({"kind": "OPENED UP" if opened else "got taken",
                        "venue": venue, "sport": sport, "court": court,
                        "date": date, "dow": dow, "time": time,
                        "was": was, "now": now})
    changes.sort(key=lambda c: (c["kind"] != "OPENED UP", c["date"], c["time"]))
    return changes


def notify(text):
    """Best-effort desktop toast on Windows; harmless elsewhere."""
    if sys.platform != "win32":
        return
    try:
        import subprocess
        safe = text.replace("'", "").replace('"', "")[:220]
        ps = (
            "[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications,"
            "ContentType=WindowsRuntime] > $null;"
            "$t=[Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent("
            "[Windows.UI.Notifications.ToastTemplateType]::ToastText02);"
            f"$t.GetElementsByTagName('text')[0].AppendChild($t.CreateTextNode('Hudle courts'))>$null;"
            f"$t.GetElementsByTagName('text')[1].AppendChild($t.CreateTextNode('{safe}'))>$null;"
            "[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("
            "'Hudle Watch').Show([Windows.UI.Notifications.ToastNotification]::new($t));"
        )
        subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                       timeout=20, capture_output=True)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

async def main():
    ap = argparse.ArgumentParser(
        description="Check Hudle court availability. Results accumulate in "
                    "output/history.json — checking today never erases what you "
                    "already know about later dates.")
    ap.add_argument("--days", type=int, default=1,
                    help="days ahead starting today (default 1 = today only). "
                         "1-4 needs no page-forward and is much faster.")
    ap.add_argument("--date", default="",
                    help="specific date(s) instead of --days: '14', '2026-09-14', "
                         "or a comma-separated list. Day-of-month picks the next "
                         "occurrence.")
    ap.add_argument("--time", default="",
                    help="only report these slot start times, e.g. '7:00 PM' or "
                         "'19:00'. Comma-separated for several. Adds a Slot Watch "
                         "sheet showing just those times everywhere.")
    ap.add_argument("--sports", default=",".join(DEFAULT_SPORTS),
                    help="comma-separated sport keywords (default: pickleball)")
    ap.add_argument("--only", default="",
                    help="comma-separated substrings; only matching venue URLs are checked")
    ap.add_argument("--headed", action="store_true", help="show the browser window")
    ap.add_argument("--workers", type=int, default=4,
                    help="how many venues to read at the same time (default 4). "
                         "Higher is faster but heavier on your PC and on Hudle; "
                         "use 1 for the old one-at-a-time behaviour.")
    args = ap.parse_args()

    sports = [s.strip() for s in args.sports.split(",") if s.strip()]
    today = dt.date.today()
    wanted_dates = parse_dates_arg(args.date, args.days, today)
    time_filter = parse_times_arg(args.time)
    if args.time and not time_filter:
        print(f"Could not understand --time {args.time!r}. Try '7:00 PM' or '19:00'.")
        return
    urls = VENUE_URLS
    if args.only:
        needles = [n.strip().lower() for n in args.only.split(",") if n.strip()]
        urls = [u for u in urls if any(n in u.lower() for n in needles)]

    os.makedirs(OUT_DIR, exist_ok=True)
    lines = []

    def log(msg):
        print(msg, flush=True)
        lines.append(msg)

    span = (f"{wanted_dates[0]:%d %b}" if len(wanted_dates) == 1
            else f"{wanted_dates[0]:%d %b}-{wanted_dates[-1]:%d %b}")
    log(f"Hudle watch — {dt.datetime.now():%d %b %Y %H:%M} — {len(urls)} venue(s), "
        f"{', '.join(sports)}, dates: {span}"
        + (f", times: {', '.join(time_filter)}" if time_filter else ""))
    if len(wanted_dates) <= 4:
        log("(within Hudle's 4-day window — no page-forward needed, so this is quick)")

    from playwright.async_api import async_playwright

    data = []
    # Venues are independent of each other, so several browser tabs can work
    # through them at once. Each worker gets its OWN page — they never share
    # one, which is what keeps a court from reading another court's grid.
    workers = max(1, min(args.workers, len(urls)))
    buckets = [urls[i::workers] for i in range(workers)]
    done = {"n": 0}

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=not args.headed)
        ctx = await browser.new_context(
            viewport={"width": 1400, "height": 900},
            user_agent=("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                        "(KHTML, like Gecko) Chrome/126.0 Safari/537.36"))

        if workers > 1:
            log(f"Running {workers} venues at a time.")

        async def run_bucket(bucket):
            page = await ctx.new_page()
            try:
                for url in bucket:
                    slug = url.rsplit("/", 2)[-2]
                    done["n"] += 1
                    log(f"[{done['n']}/{len(urls)}] {slug}")
                    try:
                        v = await scrape_venue(page, url, sports, wanted_dates,
                                               today, log)
                        if v:
                            data.append(v)
                    except Exception as e:
                        log(f"    ! {slug} failed: {e}")
                    await page.wait_for_timeout(250)   # be gentle with Hudle
            finally:
                try:
                    await page.close()
                except Exception:
                    pass

        await asyncio.gather(*(run_bucket(b) for b in buckets))
        await browser.close()

    # Changes are measured against what we knew about THESE dates before.
    history = load_history(OUT_DIR)
    prev = {k: v["status"] for k, v in history.get("slots", {}).items()}
    curr = flatten(data)
    changes = diff(prev, curr, data) if prev else None

    # Fold this run in; everything we did not look at this time survives intact.
    history = merge_into_history(history, data, today)
    n_days = save_history(OUT_DIR, history, today)

    # The whole store, past days included — this is the archive.
    archive = history_to_data(history)

    # The live views show only what you can still BOOK. Yesterday's slots would
    # only pad the sheets with rows nobody can act on.
    full = []
    for v in archive:
        courts = []
        for c in v["courts"]:
            keep = [s for s in c["slots"]
                    if dt.date.fromisoformat(s["date"]) >= today]
            if keep or c.get("missing"):
                courts.append({**c, "slots": keep})
        if courts:
            full.append({**v, "courts": courts})
    if time_filter:
        for v in full:
            for c in v["courts"]:
                c["slots"] = [s for s in c["slots"] if s["time"] in time_filter]
            v["courts"] = [c for c in v["courts"] if c["slots"]]
        full = [v for v in full if v["courts"]]

    all_dates_known = sorted({s["date"] for v in full for c in v["courts"] for s in c["slots"]})
    stamp = dt.datetime.now().strftime("%Y-%m-%d_%H%M")
    xlsx = os.path.join(OUT_DIR, f"courts_{stamp}.xlsx")
    n_open, n_rows = write_workbook(full, changes, xlsx, today,
                                    max(1, len(all_dates_known)), sports,
                                    time_filter=time_filter)

    # Flat CSV as well as the workbook. A CSV in a public repo can be pulled
    # straight into Google Sheets with =IMPORTDATA(url) — no credentials, no
    # service account, and it updates itself.
    csv_path = os.path.join(OUT_DIR, "latest.csv")
    try:
        import csv as _csv
        rows = []
        for v in full:
            for c in v["courts"]:
                for sl in c["slots"]:
                    if sl["status"] == "—":
                        continue
                    rows.append([
                        sl["date"], sl["dow"],
                        time_range(sl["time"], c.get("slot_length") or 30),
                        v["venue"], c["court"],
                        WORD.get(sl["status"], sl["status"]),
                        sl.get("price") or c.get("list_price") or "",
                        sl.get("checked", ""),
                    ])
        rows.sort(key=lambda r: (r[0], _time_key(r[2].split(" to ")[0].upper()
                                                 .replace("AM", " AM").replace("PM", " PM")), r[3]))
        with open(csv_path, "w", encoding="utf-8", newline="") as fh:
            w = _csv.writer(fh)
            w.writerow(["Date", "Day", "Slot Time", "Venue", "Court",
                        "Status", "Price", "Last checked"])
            w.writerows(rows)
        log(f"CSV rows: {len(rows):,}  ->  {csv_path}")
    except Exception as e:
        log(f"(csv export skipped: {e})")

    # Grid-shaped CSV: a row per time, a column per court — the same picture as
    # the Excel venue sheets, but every venue side by side in one table.
    try:
        import csv as _csv2
        pairs, labels = [], []
        for v in sorted(full, key=lambda x: x["venue"]):
            for c in v["courts"]:
                pairs.append((v["venue"], c["court"]))
                labels.append(f"{v['venue']} · {c['court']} "
                              f"({c.get('slot_length') or 30}m)")
        look = {}
        for v in full:
            for c in v["courts"]:
                for sl in c["slots"]:
                    look[(v["venue"], c["court"], sl["date"], sl["time"])] = sl["status"]

        gdates = sorted({sl["date"] for v in full for c in v["courts"]
                         for sl in c["slots"]})
        gtimes = sorted({sl["time"] for v in full for c in v["courts"]
                         for sl in c["slots"]}, key=_time_key)

        grid_path = os.path.join(OUT_DIR, "grid.csv")
        with open(grid_path, "w", encoding="utf-8", newline="") as fh:
            w = _csv2.writer(fh)
            w.writerow(["Date", "Day", "Time"] + labels)
            n = 0
            for d in gdates:
                dd = dt.date.fromisoformat(d)
                for t in gtimes:
                    row = [look.get((vn, cn, d, t), "—") for vn, cn in pairs]
                    if not any(x != "—" for x in row):
                        continue                      # past or closed everywhere
                    w.writerow([f"{dd:%d %b %Y}", f"{dd:%a}", t.lower()]
                               + [WORD.get(x, x) for x in row])
                    n += 1
        log(f"Grid rows: {n}  x {len(labels)} courts  ->  {grid_path}")
    except Exception as e:
        log(f"(grid export skipped: {e})")

    # One CSV per venue, laid out exactly like that venue's Excel sheet:
    # Date | Day | Slot Time | Court 1 | Court 2 | ... | Last checked
    # Each becomes its own tab in Google Sheets via IMPORTDATA.
    try:
        import csv as _csv3
        vdir = os.path.join(OUT_DIR, "venues")
        os.makedirs(vdir, exist_ok=True)
        index = []
        for v in full:
            if not v["courts"]:
                continue
            slug = re.sub(r"[^a-z0-9]+", "-", v["venue"].lower()).strip("-")[:60]
            courts = [c["court"] for c in v["courts"]]
            missing_courts = {c["court"] for c in v["courts"] if c.get("missing")}
            mins = next((c.get("slot_length") for c in v["courts"]
                         if not c.get("missing")), 30) or 30
            look = {(c["court"], sl["date"], sl["time"]): sl
                    for c in v["courts"] for sl in c["slots"]}
            vdates = sorted({sl["date"] for c in v["courts"] for sl in c["slots"]})
            vtimes = sorted({sl["time"] for c in v["courts"] for sl in c["slots"]},
                            key=_time_key)
            path = os.path.join(vdir, slug + ".csv")
            n = 0
            with open(path, "w", encoding="utf-8", newline="") as fh:
                w = _csv3.writer(fh)
                w.writerow(["Date", "Day", "Slot Time"] + courts + ["Last checked"])
                for d in vdates:
                    dd = dt.date.fromisoformat(d)
                    for t in vtimes:
                        cells = [look.get((crt, d, t)) for crt in courts]
                        if not any(c and c["status"] != "—" for c in cells):
                            continue
                        checked = next((c["checked"] for c in cells
                                        if c and c.get("checked")), "")
                        vals = []
                        for crt, c in zip(courts, cells):
                            if c:
                                vals.append(WORD.get(c["status"], c["status"]))
                            elif crt in missing_courts:
                                vals.append("not checked")
                            else:
                                vals.append("")
                        w.writerow([f"{dd:%d %b %Y}", f"{dd:%a}", time_range(t, mins)]
                                   + vals + [checked])
                        n += 1
            index.append((v["venue"], f"output/venues/{slug}.csv", len(courts), n))

        with open(os.path.join(vdir, "_index.csv"), "w", encoding="utf-8",
                  newline="") as fh:
            w = _csv3.writer(fh)
            w.writerow(["Venue", "File to use in IMPORTDATA", "Courts", "Rows"])
            for row in sorted(index):
                w.writerow(row)
        log(f"Per-venue sheets: {len(index)} files -> {vdir}")
    except Exception as e:
        log(f"(per-venue export skipped: {e})")

    # Archive: every day still held in history, oldest first. Use this to see
    # how early a given slot tends to disappear.
    try:
        import csv as _csv4
        arows = []
        for v in archive:
            for c in v["courts"]:
                for sl in c["slots"]:
                    if sl["status"] == "—":
                        continue
                    arows.append([
                        sl["date"], sl["dow"],
                        time_range(sl["time"], c.get("slot_length") or 30),
                        v["venue"], c["court"],
                        WORD.get(sl["status"], sl["status"]),
                        sl.get("price") or "", sl.get("checked", ""),
                    ])
        # Google Sheets struggles with very large IMPORTDATA files, so the
        # rolling archive covers the last 14 days. Older days stay available as
        # individual files under output/history/.
        recent = (today - dt.timedelta(days=14)).isoformat()
        arows = [r for r in arows if r[0] >= recent]
        arows.sort(key=lambda r: (r[0], r[3], r[4], r[2]))
        apath = os.path.join(OUT_DIR, "archive.csv")
        with open(apath, "w", encoding="utf-8", newline="") as fh:
            w = _csv4.writer(fh)
            w.writerow(["Date", "Day", "Slot Time", "Venue", "Court",
                        "Status", "Price", "Last checked"])
            w.writerows(arows)
        adays = len({r[0] for r in arows})
        log(f"Archive: {len(arows):,} rows across {adays} day(s) -> {apath}")
    except Exception as e:
        log(f"(archive export skipped: {e})")

    latest = os.path.join(OUT_DIR, "latest.xlsx")
    try:
        import shutil
        shutil.copyfile(xlsx, latest)
    except Exception:
        pass

    log("")
    log(f"{n_open:,} slots open across {sum(len(v['courts']) for v in full)} courts "
        f"({len(all_dates_known)} date(s) held in history).")
    if changes:
        opened = [c for c in changes if c["kind"] == "OPENED UP"]
        taken = [c for c in changes if c["kind"] != "OPENED UP"]
        log(f"{len(opened)} slot(s) opened up, {len(taken)} got taken since last run.")
        for c in opened[:12]:
            log(f"  OPEN  {c['date']} {c['dow']} {c['time']}  {c['venue']} · {c['court']}")
        if opened:
            first = opened[0]
            notify(f"{len(opened)} slot(s) freed — e.g. {first['date']} {first['time']} "
                   f"{first['venue']}")
    elif changes == []:
        log("No changes since last run.")
    log(f"Workbook: {xlsx}")

    with open(os.path.join(OUT_DIR, "last_run.txt"), "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines))


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\nStopped. Partial results were not saved — re-run to get a full sweep.")
        sys.exit(0)
