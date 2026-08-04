# -*- coding: utf-8 -*-
"""
Offline Hours Summary
=====================
Drag an Outlook calendar CSV export onto the window and this builds the monthly
Offline Hours Report (.md + .pdf) next to the CSV.

Export the CSV from Outlook with:
    File -> Save As -> Comma Separated Values (.csv)
    (or File -> Open & Export -> Import/Export -> Export to a file -> CSV)
Make sure the date range covers the whole month.

Author note: rules encoded here were agreed with Herman Ras, July 2026.
  * "Busy" and "Working Elsewhere" both COUNT as offline hours.
  * "Out of Office" means Leave and does NOT count.
  * All-day events (e.g. Leave) do not count.
  * A meeting sitting inside a longer block (e.g. during a site visit) is
    counted once - elapsed time, never double-counted.
"""

import os
import re
import csv
import io
import sys
import json
import shutil
import datetime as dt
import subprocess
import collections

# --------------------------------------------------------------------------
# CONFIG - safe to edit
# --------------------------------------------------------------------------

# 'Show time as' codes that mean Leave / not working. In Herman's export the
# all-day "Leave" event carries code 4, so that is the one excluded.
LEAVE_STATUS_CODES = {"4"}

# Everything else (2 = Busy, 3 = Working Elsewhere, ...) counts as offline hours.

# Treat all-day events as non-working days (Leave, public holidays).
ALL_DAY_IS_LEAVE = True

# Category buckets - first keyword that matches the subject wins.
CATEGORIES = [
    ("Site visits & prep",     ["site visit", "site prep", "sitevisit", "depot", "huis"]),
    ("Module & kit conversion", ["module", "conversion", "converting", "3three", "6five",
                                 "cpu", "charging", "rebuilding"]),
    ("Testing, QA & rework",   ["testing", "test setup", "test rig", "kit test", "ssr",
                                "conformal", "reflow", "plugin", "debug"]),
    ("Meetings",               ["meeting", "all-hands", "all hands", "review", "training"]),
    ("Client & tech support",  ["support", "rma", "printer", "call"]),
    ("Loom work",              ["loom", "repin"]),
    ("Admin, office & stores", ["kitchen", "organis", "office", "stock", "reel",
                                "container", "pickup", "packing", "pack", "soldering",
                                "cleanup", "clean up"]),
]
FALLBACK_CATEGORY = "Other"

MONTHS = ["January", "February", "March", "April", "May", "June", "July",
          "August", "September", "October", "November", "December"]


# --------------------------------------------------------------------------
# Remembered settings (the employee name)
# --------------------------------------------------------------------------

def settings_path():
    """
    Per-user, deliberately NOT next to the script - the script itself may live
    in a shared OneDrive folder, so one person's name must not follow it around.
    """
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    return os.path.join(base, "OfflineHoursSummary", "settings.json")


def load_settings():
    try:
        with io.open(settings_path(), "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_settings(data):
    """Best effort - a report is still worth having if this fails."""
    try:
        path = settings_path()
        d = os.path.dirname(path)
        if d and not os.path.isdir(d):
            os.makedirs(d)
        with io.open(path, "w", encoding="utf-8") as fh:
            fh.write(json.dumps(data, indent=2, ensure_ascii=False))
        return True
    except Exception:
        return False


def load_employee():
    return (load_settings().get("employee") or "").strip()


def save_employee(name):
    data = load_settings()
    data["employee"] = name
    return save_settings(data)


# --------------------------------------------------------------------------
# CSV parsing
# --------------------------------------------------------------------------

def _split_date(s):
    return [p for p in re.split(r"[/\-.]", s.strip()) if p != ""]


def detect_date_order(date_strings):
    """Work out whether the CSV uses y/m/d, d/m/y or m/d/y."""
    first_is_year = False
    max0 = max1 = 0
    for s in date_strings:
        p = _split_date(s)
        if len(p) != 3:
            continue
        if len(p[0]) == 4:
            first_is_year = True
            continue
        try:
            max0 = max(max0, int(p[0]))
            max1 = max(max1, int(p[1]))
        except ValueError:
            pass
    if first_is_year:
        return "ymd"
    if max0 > 12:
        return "dmy"
    if max1 > 12:
        return "mdy"
    return "dmy"          # South African default when genuinely ambiguous


def parse_date(s, order):
    p = _split_date(s)
    if len(p) != 3:
        raise ValueError("unrecognised date %r" % s)
    a, b, c = (int(x) for x in p)
    if order == "ymd":
        y, m, d = a, b, c
    elif order == "mdy":
        m, d, y = a, b, c
    else:
        d, m, y = a, b, c
    if y < 100:
        y += 2000
    return dt.date(y, m, d)


def parse_time(s):
    s = (s or "").strip()
    if not s:
        return dt.time(0, 0)
    ampm = None
    m = re.search(r"\s*([AaPp])\.?[Mm]\.?\s*$", s)
    if m:
        ampm = m.group(1).lower()
        s = s[:m.start()]
    parts = [int(x) for x in re.split(r"[:.]", s.strip()) if x != ""]
    while len(parts) < 3:
        parts.append(0)
    h, mi, sec = parts[0], parts[1], parts[2]
    if ampm == "p" and h < 12:
        h += 12
    if ampm == "a" and h == 12:
        h = 0
    if h >= 24:
        h, mi = 23, 59
    return dt.time(h, mi, sec)


def col(row, *names):
    """Fetch a column tolerantly (Outlook varies the header wording)."""
    for n in names:
        for k in row:
            if k and k.strip().lower() == n.lower():
                return (row[k] or "").strip()
    return ""


def read_events(path):
    with io.open(path, "r", encoding="utf-8-sig", newline="") as fh:
        rows = list(csv.DictReader(fh))
    if not rows:
        raise ValueError("The CSV contains no rows.")
    if not any(k and "subject" in k.lower() for k in rows[0]):
        raise ValueError("This does not look like an Outlook calendar export "
                         "(no 'Subject' column found).")

    order = detect_date_order([col(r, "Start Date") for r in rows] +
                              [col(r, "End Date") for r in rows])

    events, skipped = [], 0
    for r in rows:
        subj = " ".join(col(r, "Subject").split()) or "(no subject)"
        sd, st = col(r, "Start Date"), col(r, "Start Time")
        ed, et = col(r, "End Date"), col(r, "End Time")
        if not sd:
            skipped += 1
            continue
        try:
            s_date = parse_date(sd, order)
            e_date = parse_date(ed, order) if ed else s_date
        except ValueError:
            skipped += 1
            continue
        allday = col(r, "All day event").strip().lower() in ("true", "yes", "1")
        status = col(r, "Show time as")
        s = dt.datetime.combine(s_date, parse_time(st))
        e = dt.datetime.combine(e_date, parse_time(et))
        if e < s:
            e = s
        events.append(dict(subj=subj, s=s, e=e, allday=allday, status=status))
    return events, order, skipped


# --------------------------------------------------------------------------
# Cleaning: duplicates + dropped recurring instances
# --------------------------------------------------------------------------

def remove_duplicates(events):
    """Outlook sometimes writes the same recurring instance several times."""
    seen, keep, dups = set(), [], []
    for ev in events:
        k = (ev["subj"], ev["s"], ev["e"], ev["allday"])
        if k in seen:
            dups.append(ev)
        else:
            seen.add(k)
            keep.append(ev)
    return keep, dups


MIN_SERIES_INSTANCES = 3   # a real weekly series must appear at least this often
MAX_GAPS_TO_RESTORE = 2    # more holes than this and we only warn, never invent


def find_missing_recurrences(events):
    """
    Outlook's CSV export can drop instances of a weekly recurring series (and
    pile duplicate copies onto one date). Detect series that run on the same
    weekday, at the same time, for the same duration, and fill in the holes.

    Deliberately conservative - it is far worse to invent an event than to miss
    one. A group only qualifies if it already appears at least
    MIN_SERIES_INSTANCES times with at most MAX_GAPS_TO_RESTORE holes, which is
    what separates a genuine recurring meeting from an ad-hoc task that happens
    to repeat (Kitchen Management, say, which drifts around the morning).
    Anything gappier is reported as a warning instead.
    """
    working_days = set(ev["s"].date() for ev in events)
    leave_days = set(ev["s"].date() for ev in events
                     if ALL_DAY_IS_LEAVE and ev["allday"])
    subj_days = set((ev["subj"], ev["s"].date()) for ev in events)

    groups = collections.defaultdict(list)
    for ev in events:
        if ev["allday"]:
            continue
        key = (ev["subj"], ev["s"].weekday(), ev["s"].time(), ev["e"] - ev["s"])
        groups[key].append(ev)

    missing, suspicious = [], []
    for (subj, _wd, tm, dur), items in groups.items():
        dates = sorted(set(ev["s"].date() for ev in items))
        if len(dates) < MIN_SERIES_INSTANCES:
            continue
        first, last = dates[0], dates[-1]
        span_weeks = int((last - first).days / 7) + 1
        gaps = span_weeks - len(dates)
        if gaps <= 0:
            continue

        holes = []
        d = first
        have = set(dates)
        while d <= last:
            if d not in have:
                holes.append(d)
            d += dt.timedelta(days=7)

        if gaps > MAX_GAPS_TO_RESTORE:
            suspicious.append((subj, holes))
            continue

        for d in holes:
            # never invent an event on a day off, a day with no calendar at
            # all, or a day where this subject already appears at another time
            if d in leave_days or d not in working_days or (subj, d) in subj_days:
                suspicious.append((subj, [d]))
                continue
            s = dt.datetime.combine(d, tm)
            missing.append(dict(subj=subj, s=s, e=s + dur,
                                allday=False, status="2", restored=True))
    return missing, suspicious


# --------------------------------------------------------------------------
# Hour allocation
# --------------------------------------------------------------------------

def is_leave(ev):
    if ALL_DAY_IS_LEAVE and ev["allday"]:
        return True
    return ev["status"] in LEAVE_STATUS_CODES


def subtract(interval, covered):
    """Return the parts of `interval` not already inside `covered`."""
    parts = [interval]
    for cs, ce in covered:
        nxt = []
        for ps, pe in parts:
            if ce <= ps or cs >= pe:
                nxt.append((ps, pe))
                continue
            if ps < cs:
                nxt.append((ps, min(pe, cs)))
            if pe > ce:
                nxt.append((max(ps, ce), pe))
        parts = nxt
        if not parts:
            break
    return parts


def allocate_day(day_events):
    """
    Give each event the time it actually adds to the day. A block fully inside
    a longer one gets 0.00 (marked 'concurrent'); a partial overlap gets the
    uncovered remainder. Sum of allocations == elapsed time for the day.
    """
    timed = [ev for ev in day_events if not is_leave(ev) and ev["e"] > ev["s"]]
    timed.sort(key=lambda ev: (ev["s"], -(ev["e"] - ev["s"]).total_seconds()))
    covered, out = [], {}
    for ev in timed:
        parts = subtract((ev["s"], ev["e"]), covered)
        hours = sum((b - a).total_seconds() for a, b in parts) / 3600.0
        out[id(ev)] = hours
        covered.append((ev["s"], ev["e"]))
        covered.sort()
    return out


def categorise(subject):
    t = subject.lower()
    for name, keys in CATEGORIES:
        for k in keys:
            if k in t:
                return name
    return FALLBACK_CATEGORY


# --------------------------------------------------------------------------
# Report building
# --------------------------------------------------------------------------

def build_report(csv_path, restore_missing=True, log=print):
    events, order, skipped = read_events(csv_path)
    log("Read %d rows (date format detected: %s)." % (len(events), order))
    if skipped:
        log("  %d row(s) skipped - unreadable dates." % skipped)

    events, dups = remove_duplicates(events)
    if dups:
        log("Removed %d duplicate event(s):" % len(dups))
        for d in dups[:10]:
            log("   - %s  %s" % (d["s"].strftime("%a %d %b %H:%M"), d["subj"]))

    # keep only the dominant month
    counter = collections.Counter((ev["s"].year, ev["s"].month) for ev in events)
    (year, month), _ = counter.most_common(1)[0]
    outside = [ev for ev in events if (ev["s"].year, ev["s"].month) != (year, month)]
    events = [ev for ev in events if (ev["s"].year, ev["s"].month) == (year, month)]
    log("Month detected: %s %d  (%d events)." % (MONTHS[month - 1], year, len(events)))
    if outside:
        log("  %d event(s) outside that month were ignored." % len(outside))

    restored, suspicious = [], []
    if restore_missing:
        restored, suspicious = find_missing_recurrences(events)
        if restored:
            log("Restored %d recurring instance(s) Outlook dropped:" % len(restored))
            for r in restored:
                log("   + %s  %s" % (r["s"].strftime("%a %d %b %H:%M"), r["subj"]))
            events.extend(restored)
        if suspicious:
            log("Possible gaps left alone (check these if the total looks low):")
            for subj, holes in suspicious[:8]:
                log("   ? %s - %s" % (subj, ", ".join(d.strftime("%a %d %b")
                                                      for d in holes)))

    for ev in events:
        ev.setdefault("restored", False)
    events.sort(key=lambda ev: (ev["s"], ev["e"]))

    byday = collections.OrderedDict()
    for ev in events:
        byday.setdefault(ev["s"].date(), []).append(ev)

    alloc, cats = {}, collections.Counter()
    for day, items in byday.items():
        a = allocate_day(items)
        alloc.update(a)
        for ev in items:
            h = a.get(id(ev), 0.0)
            if h > 0:
                cats[categorise(ev["subj"])] += h

    # group days into weeks
    weeks, seen = [], set()
    for day in sorted(byday):
        key = day.isocalendar()[:2]
        if key not in seen:
            seen.add(key)
            weeks.append((key, []))
        weeks[-1][1].append(day)

    grand = sum(alloc.values())
    leave_days = [d for d, items in byday.items() if all(is_leave(ev) for ev in items)]

    return dict(year=year, month=month, byday=byday, alloc=alloc, cats=cats,
                weeks=weeks, grand=grand, events=events, dups=dups,
                restored=restored, suspicious=suspicious,
                leave_days=leave_days, csv_path=csv_path)


def month_title(rep):
    return "%s %d" % (MONTHS[rep["month"] - 1], rep["year"])


def notes_for(rep):
    n = [("Source", "Built from the Outlook calendar export `%s`. Times and titles "
                    "are taken exactly as exported - nothing is estimated."
                    % os.path.basename(rep["csv_path"])),
         ("What counts", "Busy and Working Elsewhere both count as offline hours. "
                         "Out of Office is treated as Leave and is excluded, as are "
                         "all-day events.")]
    if rep["dups"]:
        n.append(("Duplicates removed",
                  "Outlook exported %d event(s) more than once; the extra copies were "
                  "discarded." % len(rep["dups"])))
    if rep["restored"]:
        lst = ", ".join(sorted(set(r["s"].strftime("%a %d %b") for r in rep["restored"])))
        n.append(("Recurring instances restored",
                  "Outlook's export dropped %d instance(s) of a weekly recurring "
                  "series (%s). They were added back to match the rest of the series."
                  % (len(rep["restored"]), lst)))
    n.append(("Overlaps", "Where a shorter event sits inside a longer one (a meeting "
                          "during a site visit, say) it is shown but marked "
                          "'concurrent' and adds no hours, so elapsed time is only "
                          "counted once."))
    if rep["leave_days"]:
        lst = ", ".join(d.strftime("%a %d %b") for d in sorted(rep["leave_days"]))
        n.append(("Leave", "%s - no hours counted." % lst))
    return n


def write_markdown(rep, path, employee):
    L = ["# Offline Hours Report - %s" % month_title(rep), "",
         "**Employee:** %s  " % employee,
         "**Period:** %s  " % month_title(rep),
         "**Source:** `%s`" % os.path.basename(rep["csv_path"]), "", "---", ""]
    for i, (_key, days) in enumerate(rep["weeks"], 1):
        lo, hi = min(days), max(days)
        label = "Week %d: %d-%d %s" % (i, lo.day, hi.day, MONTHS[rep["month"] - 1])
        L += [("## " + label), "", "| Date | Activity | Time | Hours |",
              "|------|----------|------|------:|"]
        wtot = 0.0
        for d in days:
            dtot = 0.0
            for ev in rep["byday"][d]:
                ds = d.strftime("%a ") + str(d.day)
                if is_leave(ev):
                    L.append("| %s | %s%s | - | 0.00 |"
                             % (ds, ev["subj"], " (all day)" if ev["allday"] else " (out of office)"))
                    continue
                h = rep["alloc"].get(id(ev), 0.0)
                dtot += h
                tm = ev["s"].strftime("%H:%M") + "-" + ev["e"].strftime("%H:%M")
                mark = " *(restored)*" if ev.get("restored") else ""
                val = "*concurrent*" if h == 0 else "%.2f" % h
                L.append("| %s | %s%s | %s | %s |" % (ds, ev["subj"], mark, tm, val))
            L.append("| **%s Total** | | | **%.2f** |"
                     % (d.strftime("%a ") + str(d.day), dtot))
            wtot += dtot
        L += ["", "**Week %d Total: %.2f hours**" % (i, wtot), "", "---", ""]

    L += ["## Monthly Summary", "", "| Week | Hours |", "|------|------:|"]
    for i, (_key, days) in enumerate(rep["weeks"], 1):
        wt = sum(rep["alloc"].get(id(ev), 0.0) for d in days for ev in rep["byday"][d])
        L.append("| Week %d — %d-%d %s | %.2f |"
                 % (i, min(days).day, max(days).day, MONTHS[rep["month"] - 1], wt))
    L += ["| **Grand Total** | **%.2f** |" % rep["grand"], "", "---", "",
          "## Activity Breakdown", "", "| Category | Hours |", "|----------|------:|"]
    for name, h in rep["cats"].most_common():
        L.append("| %s | %.2f |" % (name, h))
    L += ["| **Total** | **%.2f** |" % sum(rep["cats"].values()), "", "---", "",
          "## Notes", ""]
    for head, body in notes_for(rep):
        L.append("- **%s:** %s" % (head, body))
    io.open(path, "w", encoding="utf-8").write("\n".join(L) + "\n")


CSS = ("body{font-family:Calibri,Arial,sans-serif;font-size:12px;color:#1a1a1a;margin:28px}"
       "h1{font-size:20px;margin-bottom:4px}"
       "h2{font-size:15px;margin-top:22px;border-bottom:1px solid #999;padding-bottom:3px}"
       "p.meta{margin:2px 0}table{border-collapse:collapse;width:100%;margin:8px 0 4px}"
       "th,td{border:1px solid #bbb;padding:4px 8px;text-align:left;font-size:11.5px}"
       "th{background:#2f4f6f;color:#fff}tr.total td{font-weight:bold;background:#eef2f6}"
       "tr.grand td{font-weight:bold;background:#2f4f6f;color:#fff}"
       "td.num,th.num{text-align:right}tr.conc td{color:#777;font-style:italic}"
       "tr.leave td{color:#7a4a7a;font-style:italic}"
       "hr{border:none;border-top:1px solid #ccc;margin:16px 0}"
       "dl.note{font-size:11px;color:#333}dt{font-weight:bold;margin-top:6px}"
       "dd{margin:0 0 0 14px}")


def esc(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def write_html(rep, path, employee):
    H = ["<!DOCTYPE html><html><head><meta charset='utf-8'><title>Offline Hours - %s</title>"
         "<style>%s</style></head><body>" % (month_title(rep), CSS),
         "<h1>Offline Hours Report &mdash; %s</h1>" % month_title(rep),
         "<p class='meta'><b>Employee:</b> %s</p>" % esc(employee),
         "<p class='meta'><b>Period:</b> %s</p>" % month_title(rep),
         "<p class='meta'><b>Source:</b> <code>%s</code></p><hr>"
         % esc(os.path.basename(rep["csv_path"]))]
    for i, (_key, days) in enumerate(rep["weeks"], 1):
        lo, hi = min(days), max(days)
        H.append("<h2>Week %d: %d&ndash;%d %s</h2><table><tr><th>Date</th><th>Activity</th>"
                 "<th>Time</th><th class='num'>Hours</th></tr>"
                 % (i, lo.day, hi.day, MONTHS[rep["month"] - 1]))
        wtot = 0.0
        for d in days:
            dtot = 0.0
            ds = d.strftime("%a ") + str(d.day)
            for ev in rep["byday"][d]:
                if is_leave(ev):
                    tag = " (all day)" if ev["allday"] else " (out of office)"
                    H.append("<tr class='leave'><td>%s</td><td>%s%s</td><td>-</td>"
                             "<td class='num'>0.00</td></tr>" % (ds, esc(ev["subj"]), tag))
                    continue
                h = rep["alloc"].get(id(ev), 0.0)
                dtot += h
                tm = ev["s"].strftime("%H:%M") + "-" + ev["e"].strftime("%H:%M")
                mark = " (restored)" if ev.get("restored") else ""
                cls = " class='conc'" if h == 0 else ""
                val = "concurrent" if h == 0 else "%.2f" % h
                H.append("<tr%s><td>%s</td><td>%s%s</td><td>%s</td>"
                         "<td class='num'>%s</td></tr>"
                         % (cls, ds, esc(ev["subj"]), mark, tm, val))
            H.append("<tr class='total'><td colspan='3'>%s Total</td>"
                     "<td class='num'>%.2f</td></tr>" % (ds, dtot))
            wtot += dtot
        H.append("</table><p><b>Week %d Total: %.2f hours</b></p>" % (i, wtot))
    H.append("<hr><h2>Monthly Summary</h2><table><tr><th>Week</th><th class='num'>Hours</th></tr>")
    for i, (_key, days) in enumerate(rep["weeks"], 1):
        wt = sum(rep["alloc"].get(id(ev), 0.0) for d in days for ev in rep["byday"][d])
        H.append("<tr><td>Week %d &mdash; %d&ndash;%d %s</td><td class='num'>%.2f</td></tr>"
                 % (i, min(days).day, max(days).day, MONTHS[rep["month"] - 1], wt))
    H.append("<tr class='grand'><td>Grand Total</td><td class='num'>%.2f</td></tr></table>"
             % rep["grand"])
    H.append("<hr><h2>Activity Breakdown</h2><table><tr><th>Category</th>"
             "<th class='num'>Hours</th></tr>")
    for name, h in rep["cats"].most_common():
        H.append("<tr><td>%s</td><td class='num'>%.2f</td></tr>" % (esc(name), h))
    H.append("<tr class='grand'><td>Total</td><td class='num'>%.2f</td></tr></table>"
             % sum(rep["cats"].values()))
    H.append("<hr><h2>Notes</h2><dl class='note'>")
    for head, body in notes_for(rep):
        H.append("<dt>%s</dt><dd>%s</dd>" % (esc(head), esc(body).replace("`", "")))
    H.append("</dl></body></html>")
    io.open(path, "w", encoding="utf-8").write("\n".join(H))


def find_browser():
    cands = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]
    for c in cands:
        if os.path.isfile(c):
            return c
    return shutil.which("msedge") or shutil.which("chrome")


def html_to_pdf(html_path, pdf_path, log=print):
    browser = find_browser()
    if not browser:
        log("No Edge/Chrome found - PDF skipped (HTML written instead).")
        return False
    tmp = os.path.join(os.environ.get("TEMP", "."), "_offline_hours_tmp.pdf")
    for f in (tmp,):
        try:
            os.remove(f)
        except OSError:
            pass
    cmd = [browser, "--headless", "--disable-gpu", "--no-pdf-header-footer",
           "--print-to-pdf=" + tmp, "file:///" + html_path.replace("\\", "/")]
    try:
        subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=120)
    except Exception as exc:
        log("PDF render failed: %s" % exc)
        return False
    if not os.path.isfile(tmp):
        log("PDF render produced no file.")
        return False
    try:
        shutil.copyfile(tmp, pdf_path)
        os.remove(tmp)
        return True
    except (IOError, OSError) as exc:
        log("Could not write the PDF: %s" % exc)
        log("  -> If it is open in Adobe Acrobat, close it and run again.")
        return False


def generate(csv_path, employee, restore_missing=True, log=print):
    employee = (employee or "").strip()
    if not employee:
        raise ValueError("No name given - the report needs an employee name.")
    rep = build_report(csv_path, restore_missing, log)
    outdir = os.path.dirname(os.path.abspath(csv_path))
    stem = "Offline_Hours_%s%d" % (MONTHS[rep["month"] - 1], rep["year"])
    md = os.path.join(outdir, stem + ".md")
    html = os.path.join(outdir, stem + ".html")
    pdf = os.path.join(outdir, stem + ".pdf")

    write_markdown(rep, md, employee)
    write_html(rep, html, employee)
    ok = html_to_pdf(html, pdf, log)
    try:
        os.remove(html)
    except OSError:
        pass

    log("")
    log("-" * 46)
    for i, (_k, days) in enumerate(rep["weeks"], 1):
        wt = sum(rep["alloc"].get(id(ev), 0.0) for d in days for ev in rep["byday"][d])
        log("  Week %d (%d-%d)%s%6.2f h"
            % (i, min(days).day, max(days).day, " " * 10, wt))
    log("  %-28s %6.2f h" % ("GRAND TOTAL", rep["grand"]))
    log("-" * 46)
    log("")
    log("Saved: %s" % os.path.basename(md))
    if ok:
        log("Saved: %s" % os.path.basename(pdf))
    return rep, md, (pdf if ok else None), outdir


# --------------------------------------------------------------------------
# GUI
# --------------------------------------------------------------------------

def run_gui(initial=None, name=None):
    import tkinter as tk
    from tkinter import filedialog, ttk

    try:
        from tkinterdnd2 import TkinterDnD, DND_FILES
        root = TkinterDnD.Tk()
        has_dnd = True
    except Exception:
        root = tk.Tk()
        has_dnd = False

    root.title("Offline Hours Summary")
    root.geometry("640x620")
    root.configure(bg="#f4f6f8")

    state = {"pdf": None, "outdir": None, "csv": None}

    tk.Label(root, text="Offline Hours Summary", bg="#f4f6f8",
             font=("Segoe UI", 16, "bold")).pack(pady=(16, 2))
    tk.Label(root, text="Turn an Outlook calendar CSV export into the monthly report",
             bg="#f4f6f8", fg="#555", font=("Segoe UI", 9)).pack()

    remembered = (name or "").strip() or load_employee()
    who = tk.Frame(root, bg="#f4f6f8")
    who.pack(fill="x", padx=22, pady=(14, 0))
    tk.Label(who, text="Your name", bg="#f4f6f8", font=("Segoe UI", 9, "bold"),
             anchor="w").pack(anchor="w")
    name_var = tk.StringVar(value=remembered)
    name_entry = tk.Entry(who, textvariable=name_var, font=("Segoe UI", 11),
                          relief="solid", bd=1)
    name_entry.pack(fill="x", pady=(2, 2))
    tk.Label(who, text="Appears as the Employee on the report. Remembered for next time.",
             bg="#f4f6f8", fg="#777", font=("Segoe UI", 8), anchor="w").pack(anchor="w")

    drop = tk.Label(root,
                    text=("Drag your calendar .CSV here\n\nor click to browse"
                          if has_dnd else "Click here to choose your calendar .CSV"),
                    bg="#ffffff", fg="#2f4f6f", font=("Segoe UI", 11),
                    relief="ridge", bd=2, width=56, height=6)
    drop.pack(pady=(12, 14), padx=20, fill="x")

    opts = tk.Frame(root, bg="#f4f6f8")
    opts.pack(fill="x", padx=22)
    restore_var = tk.BooleanVar(value=True)
    tk.Checkbutton(opts, text="Restore recurring meetings that Outlook's export dropped",
                   variable=restore_var, bg="#f4f6f8", font=("Segoe UI", 9),
                   anchor="w").pack(anchor="w")

    txt = tk.Text(root, height=13, font=("Consolas", 9), bg="#1e1e1e", fg="#dcdcdc",
                  relief="flat", wrap="word")
    txt.pack(fill="both", expand=True, padx=20, pady=(10, 6))

    btns = tk.Frame(root, bg="#f4f6f8")
    btns.pack(fill="x", padx=20, pady=(0, 14))
    make_btn = ttk.Button(btns, text="Generate report", state="disabled")
    make_btn.pack(side="left")
    open_pdf_btn = ttk.Button(btns, text="Open PDF", state="disabled")
    open_pdf_btn.pack(side="left", padx=6)
    open_dir_btn = ttk.Button(btns, text="Open folder", state="disabled")
    open_dir_btn.pack(side="left")
    ttk.Button(btns, text="Quit", command=root.destroy).pack(side="right")

    def log(msg=""):
        txt.insert("end", str(msg) + "\n")
        txt.see("end")
        root.update_idletasks()

    def run(_e=None):
        path = state["csv"]
        name = name_var.get().strip()
        txt.delete("1.0", "end")
        state["pdf"] = state["outdir"] = None
        open_pdf_btn.config(state="disabled")
        open_dir_btn.config(state="disabled")
        if not path:
            log("Choose your calendar .CSV first.")
            return
        if not name:
            log("Enter your name above - it goes on the report as the Employee.")
            name_entry.focus_set()
            return
        if name != load_employee():
            if not save_employee(name):
                log("(Could not save your name for next time - carrying on.)")
        log("Processing %s" % os.path.basename(path))
        log("Employee: %s" % name)
        log("")
        try:
            _rep, _md, pdf, outdir = generate(path, name, restore_var.get(), log)
        except Exception as exc:
            log("")
            log("ERROR: %s" % exc)
            return
        state["pdf"], state["outdir"] = pdf, outdir
        if pdf:
            open_pdf_btn.config(state="normal")
        open_dir_btn.config(state="normal")

    def select(path):
        """Take a file, then run straight away if we already know the name."""
        if not path or not os.path.isfile(path):
            log("No file selected.")
            return
        if not path.lower().endswith(".csv"):
            log("That is not a .csv file.")
            log("Export from Outlook:  File > Save As > Comma Separated Values")
            return
        state["csv"] = path
        drop.config(text=os.path.basename(path))
        make_btn.config(state="normal")
        if name_var.get().strip():
            run()
        else:
            log("Ready: %s" % os.path.basename(path))
            log("Enter your name above, then click 'Generate report'.")
            name_entry.focus_set()

    def browse(_e=None):
        select(filedialog.askopenfilename(title="Choose the Outlook calendar CSV",
                                         filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]))

    make_btn.config(command=run)
    name_entry.bind("<Return>", run)
    drop.bind("<Button-1>", browse)
    if has_dnd:
        def on_drop(event):
            raw = event.data.strip()
            files = re.findall(r"\{([^}]*)\}", raw) or raw.split()
            select(files[0] if files else "")
        drop.drop_target_register(DND_FILES)
        drop.dnd_bind("<<Drop>>", on_drop)

    open_pdf_btn.config(command=lambda: state["pdf"] and os.startfile(state["pdf"]))
    open_dir_btn.config(command=lambda: state["outdir"] and os.startfile(state["outdir"]))

    if not has_dnd:
        log("(Drag-and-drop needs 'tkinterdnd2'.  Install with:")
        log("     pip install tkinterdnd2")
        log(" Clicking the box to browse works fine in the meantime.)")
        log("")

    if not remembered:
        name_entry.focus_set()
        log("Enter your name above - it goes on the report as the Employee.")
        log("It is remembered afterwards, and you can change it any time.")
        log("")

    if initial:
        root.after(200, lambda: select(initial))

    root.mainloop()


def parse_cli(argv):
    """Pull out file arguments and --name (which takes a value) separately."""
    files, name, i = [], None, 0
    while i < len(argv):
        a = argv[i]
        if a.startswith("--name="):
            name = a.split("=", 1)[1]
        elif a == "--name":
            i += 1
            if i < len(argv):
                name = argv[i]
        elif not a.startswith("-"):
            files.append(a)
        i += 1
    return files, name


def main():
    args, cli_name = parse_cli(sys.argv[1:])
    if "--nogui" in sys.argv:
        if not args:
            print("usage: offline_hours.py <calendar.csv> --nogui "
                  "[--name \"Your Name\"] [--no-restore]")
            return 2
        if not os.path.isfile(args[0]):
            print("ERROR: no such file - %s" % args[0])
            return 2
        name = (cli_name or load_employee()).strip()
        if not name:
            print("ERROR: no name available. Pass --name \"Your Name\" once and it "
                  "is remembered afterwards.")
            return 2
        if cli_name and name != load_employee():
            save_employee(name)
        try:
            generate(args[0], name, "--no-restore" not in sys.argv)
        except Exception as exc:
            print("ERROR: %s" % exc)
            return 1
        return 0
    run_gui(args[0] if args else None, cli_name)
    return 0


if __name__ == "__main__":
    sys.exit(main())
