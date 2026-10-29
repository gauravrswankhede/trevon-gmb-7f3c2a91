"""Build data.json for the Trevon Google Business Profile dashboard.

Usage: python3 -I build_gmb.py <folder with the .xlsx files> [data.json]

Reads the agency (Digital Dhanu) master workbooks, one per property, plus the
UC-* delegation tracker, and writes weekly metrics, traffic-light signals and
action items. Nothing private (passwords, phone numbers, GST) is copied out.
"""
import json
import os
import re
import sys
from datetime import datetime, timedelta, timezone

import openpyxl

IST = timezone(timedelta(hours=5, minutes=30))
TODAY = datetime.now(IST).replace(tzinfo=None)

PROPERTIES = {
    "ranchi": {"name": "Trevon Hotel Gujarat Centrale", "city": "Ranchi"},
    "haldwani": {"name": "Trevon The Shire Inn", "city": "Haldwani"},
    "gurugram": {"name": "Trevon Inn Gurugram", "city": "Gurugram"},
    "delhi": {"name": "Trevon Delhi", "city": "Delhi"},
    "goa": {"name": "Trevon Goa", "city": "Goa"},
    "dhanbad": {"name": "Trevon Dhanbad", "city": "Dhanbad"},
    "chakrata": {"name": "Trevon Chakrata", "city": "Chakrata"},
    "corbett": {"name": "Trevon Pine Corbett", "city": "Corbett"},
}
CITY_ALIASES = {"gurgaon": "gurugram"}

REPORT = "06_Weekly& Monthly Report_Numbe"
ROWS = {  # row in the weekly block -> key
    11: "interactions", 12: "calls", 13: "chats", 14: "directions",
    15: "website", 16: "views", 18: "reviews_total", 19: "negative",
    20: "responses", 21: "rating", 23: "posts", 24: "events", 25: "offers",
    26: "citations", 27: "products",
}
ACTIVITY_COLS = {  # planned column in 04_Regular Activities -> task
    3: "GMB post", 7: "Product", 11: "Offer", 15: "Service", 19: "Review response",
    23: "Citation", 26: "Instagram", 29: "Facebook", 32: "LinkedIn",
    35: "Pinterest", 38: "Images", 41: "Video",
}


def num(v):
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = v.strip().replace(",", "")
        try:
            return float(s)
        except ValueError:
            return None
    return None


def as_date(v):
    if isinstance(v, datetime):
        return v
    if isinstance(v, str):
        for f in ("%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(v.strip(), f)
            except ValueError:
                pass
    return None


def which_property(title):
    t = title.lower()
    for alias, key in CITY_ALIASES.items():
        if alias in t:
            return key
    for key in PROPERTIES:
        if key in t:
            return key
    return None


def weekly(ws):
    weeks = []
    for c in range(4, 60):
        d = as_date(ws.cell(3, c).value)
        if not d:
            break
        rec = {"week": c - 4, "start": d.strftime("%Y-%m-%d")}
        for r, k in ROWS.items():
            v = ws.cell(r, c).value
            rec[k] = num(v)
            if isinstance(v, str) and "not updated" in v.lower():
                rec.setdefault("not_updated", []).append(k)
        rec["ranks"] = [num(ws.cell(r, c).value) for r in range(5, 10)]
        weeks.append(rec)
    # the newest week column whose report was due (weekly figures go in on the Monday after)
    due = [w["start"] for w in weeks if datetime.strptime(w["start"], "%Y-%m-%d") + timedelta(days=1) <= TODAY]
    keywords = [str(ws.cell(r, 3).value or "").strip() for r in range(5, 10)]
    # keep weeks that have any performance number filled in
    filled = [w for w in weeks if any(w[k] is not None for k in ("views", "interactions", "reviews_total"))
              or any(x is not None for x in w["ranks"])]
    return keywords, filled, (due[-1] if due else None)


def activities(ws):
    per = {}
    for pc, name in ACTIVITY_COLS.items():
        planned = done = ontime = 0
        for r in range(6, ws.max_row + 1):
            p = as_date(ws.cell(r, pc).value)
            if not p or p.date() > TODAY.date():
                continue
            a = ws.cell(r, pc + 1).value
            if isinstance(a, str) and a.strip().lower() in ("holiday", "na", "n/a", "not applicable"):
                continue
            planned += 1
            ad = as_date(a)
            if ad or (isinstance(a, str) and a.strip() and not a.startswith("#")):
                done += 1
                if ad and ad.date() <= (p + timedelta(days=1)).date():
                    ontime += 1
        if planned:
            per[name] = {"planned": planned, "done": done, "ontime": ontime}
    return per


def pms(ws):
    get = lambda a: ws[a].value
    cit = num(get("N3"))
    return {
        "start": as_date(get("C6")).strftime("%Y-%m-%d") if as_date(get("C6")) else None,
        "age_days": num(get("C22")),
        "setup_owner": get("C16"),
        "owner": get("C17") or None,
        "citation_owner": get("C18"),
        "status": get("C20"),
        "citation_pct": round(cit * 100, 1) if cit is not None else None,
    }


# ---------- traffic lights ----------
G, A, R, NA = "green", "amber", "red", "grey"


def pct_change(cur, prev):
    if cur is None or prev is None or prev < 10:  # tiny week-0 baselines give silly percentages
        return None
    return (cur - prev) / prev * 100


def trend_status(ch):
    if ch is None:
        return NA
    return G if ch >= -5 else (A if ch >= -20 else R)


def avg(xs):
    xs = [x for x in xs if x is not None]
    return sum(xs) / len(xs) if xs else None


def signals_for(p):
    w = p["weeks"]
    cur = w[-1] if w else {}
    prev = w[-2] if len(w) > 1 else {}
    out = []

    def prev_val(key):
        for x in reversed(w[:-1]):
            if x.get(key) is not None:
                return x[key]
        return None

    def add(key, label, value, fmt, status, detail, change=None):
        out.append({"key": key, "label": label, "value": value, "fmt": fmt,
                    "status": status, "detail": detail, "change": change})

    days = (TODAY - datetime.strptime(cur["start"], "%Y-%m-%d")).days if cur else None
    missing = cur.get("not_updated", []) if cur else []
    due = p.get("expected_week")
    if days is None:
        st = R
    else:
        st = G if days <= 8 else (A if days <= 14 else R)
    if cur and due and due > cur["start"]:
        st = R
        missing = ["week of " + due] + missing
    if missing and st == G:
        st = A
    det = f"Latest week starts {cur.get('start')}, {days} days ago" if cur else "No weekly data"
    if missing:
        det += "; missing: " + ", ".join(missing)
    add("fresh", "Report freshness", days, "days", st, det)

    for key, label in (("views", "Profile views"), ("interactions", "Interactions"), ("calls", "Calls"),
                       ("directions", "Direction requests")):
        ch = pct_change(cur.get(key), prev_val(key))
        pv = prev_val(key)
        if cur.get(key) is None:
            st = NA
        elif ch is None and pv is not None:
            st = G if cur[key] >= pv else A
        else:
            st = trend_status(ch)
        det = ("vs previous week" if ch is not None else
               (f"up from {pv:g} (baseline too small for a %)" if pv is not None else "no comparison week yet"))
        add(key, label, cur.get(key), "int", st, det, ch)

    ranks = cur.get("ranks", [])
    pranks = prev.get("ranks", [])
    a, pa = avg(ranks), avg(pranks)
    if a is None:
        st = NA
    else:
        st = G if a <= 10 else (A if a <= 20 else R)
        if pa is not None and a - pa > 3 and st == G:
            st = A
    first = next((avg(x["ranks"]) for x in w if avg(x["ranks"]) is not None), None)
    det = f"average Google position of the 5 target keywords (lower is better); started at {first:.0f}" if first else ""
    add("rank", "Avg keyword position", round(a, 1) if a else None, "rank", st, det,
        (pa - a) if (a is not None and pa is not None) else None)
    top10 = sum(1 for x in ranks if x is not None and x <= 10)
    add("top10", "Keywords in top 10", top10, "of5", G if top10 >= 3 else (A if top10 >= 1 else R),
        "target keywords ranking on Google page 1")

    rt = cur.get("rating")
    add("rating", "Google rating", rt, "rating", NA if rt is None else (G if rt >= 4.5 else (A if rt >= 4.0 else R)),
        f"{int(cur.get('reviews_total') or 0)} reviews in total")
    new = (cur.get("reviews_total") or 0) - (prev.get("reviews_total") or 0) if prev else None
    add("new_reviews", "New reviews this week", new, "int",
        NA if new is None else (G if new >= 5 else (A if new >= 2 else R)), "target 5+ a week")
    resp = cur.get("responses")
    if new:
        ratio = (resp or 0) / new
        st = G if ratio >= 1 else (A if ratio >= 0.5 else R)
        det = f"{int(resp or 0)} replies to {int(new)} new reviews"
    else:
        st, det = (G if resp else NA), f"{int(resp or 0)} replies this week"
    add("responses", "Review replies", resp, "int", st, det)
    neg = cur.get("negative")
    add("negative", "Negative reviews", neg, "int", NA if neg is None else (G if neg == 0 else (A if neg == 1 else R)),
        "this week")
    posts = sum((cur.get(k) or 0) for k in ("posts", "events", "offers")) if cur else None
    add("posts", "GMB posts this week", posts, "int", NA if posts is None else (G if posts >= 2 else (A if posts == 1 else R)),
        "updates + events + offers; target 2+")

    acts = p["activities"]
    pl = sum(v["planned"] for v in acts.values())
    ot = sum(v["ontime"] for v in acts.values())
    dn = sum(v["done"] for v in acts.values())
    rate = ot / pl * 100 if pl else None
    add("ontime", "Agency tasks on time", round(rate) if rate is not None else None, "pct",
        NA if rate is None else (G if rate >= 90 else (A if rate >= 70 else R)),
        f"{ot} of {pl} scheduled tasks on time, {dn} done in total")
    cp = p["pms"]["citation_pct"]
    add("citations", "Citations built", cp, "pct", NA if cp is None else (G if cp >= 80 else (A if cp >= 50 else R)),
        "share of planned directory listings completed")
    return out


ACTION_TEXT = {
    "fresh": "Ask {owner} to fill this week's numbers in the weekly report (due every Monday 12 PM){extra}.",
    "views": "Profile views fell {chg}. Post 2 fresh photo updates and check the profile for edits or suspensions.",
    "interactions": "Interactions fell {chg}. Add an offer post and make sure the booking link and phone number are live.",
    "calls": "Calls fell {chg}. Confirm the listed phone number rings through and is answered 24×7.",
    "directions": "Direction requests fell {chg}. Check the map pin and address match the website.",
    "rank": "Keywords average position {val}. Work keywords into the description, services and posts; build local citations.",
    "top10": "Only {val} of 5 target keywords on page 1. Prioritise the closest keyword with posts and Q&A that use it.",
    "rating": "Rating is {val}. Ask happy guests for reviews at check-out (QR card at reception) and reply to every low rating.",
    "new_reviews": "Only {val} new reviews this week. Send review links to checked-out guests daily.",
    "responses": "Not every new review got a reply. Reply to all reviews within 48 hours.",
    "negative": "{val} negative review(s) this week. Reply within 24 hours and flag the issue to the hotel manager.",
    "posts": "Only {val} GMB posts this week. Publish at least 2 updates (rooms, food, offers).",
    "ontime": "Agency delivered only {val}% of scheduled tasks on time. Raise it in the weekly call with the agency.",
    "citations": "Citations at {val}%. Ask {cit_owner} to finish the remaining directory listings.",
}


def actions_for(p):
    acts = []
    for s in p["signals"]:
        if s["status"] not in (R, A):
            continue
        t = ACTION_TEXT.get(s["key"])
        if not t:
            continue
        chg = f"{abs(s['change']):.0f}% week on week" if isinstance(s.get("change"), (int, float)) else ""
        val = s["value"]
        val = f"{val:g}" if isinstance(val, float) else val
        extra = ""
        if s["key"] == "fresh" and "missing" in s["detail"]:
            extra = " (" + s["detail"].split("missing: ")[1] + ")"
        acts.append({
            "property": p["city"], "signal": s["label"], "priority": "High" if s["status"] == R else "Medium",
            "status": s["status"],
            "text": t.format(owner=p["pms"]["owner"] or p["pms"]["setup_owner"] or "the agency",
                             cit_owner=p["pms"]["citation_owner"] or "the agency", chg=chg, val=val, extra=extra),
        })
    acts.sort(key=lambda a: 0 if a["priority"] == "High" else 1)
    return acts


def tasks(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb["Master"]
    out = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[2]:
            continue
        due = as_date(row[7]) or as_date(row[6]) or as_date(row[5]) or as_date(row[4])
        status = (row[11] or "Pending").strip()
        text = str(row[2]).strip()
        prop = which_property(text) or "All"
        overdue = status != "Completed" and due is not None and due.date() < TODAY.date()
        out.append({"doer": str(row[1] or "").strip(), "task": text, "due": due.strftime("%Y-%m-%d") if due else None,
                    "status": status, "priority": row[9] or "Medium",
                    "property": PROPERTIES[prop]["city"] if prop in PROPERTIES else "All",
                    "assigned_by": row[19], "overdue": overdue})
    return out


CRITICAL = {"rank", "rating", "fresh", "views"}


def score(signals):
    """0-100 health score; critical signals count double."""
    pts = {G: 2, A: 1, R: 0}
    got = tot = 0
    for s in signals:
        if s["status"] in pts:
            wt = 2 if s["key"] in CRITICAL else 1
            got += pts[s["status"]] * wt
            tot += 2 * wt
    return round(got / tot * 100) if tot else None


def main():
    folder = sys.argv[1]
    out_path = sys.argv[2] if len(sys.argv) > 2 else "data.json"
    props, task_list = [], []
    for f in sorted(os.listdir(folder)):
        if not f.lower().endswith(".xlsx"):
            continue
        path = os.path.join(folder, f)
        wb = openpyxl.load_workbook(path, data_only=True)
        if "Master" in wb.sheetnames and REPORT not in wb.sheetnames:
            task_list += tasks(path)
            continue
        if REPORT not in wb.sheetnames:
            continue
        title = str(wb[REPORT]["G1"].value or f)
        key = which_property(title) or which_property(f)
        if not key:
            continue
        keywords, weeks, expected = weekly(wb[REPORT])
        p = {"key": key, **PROPERTIES[key], "keywords": keywords, "weeks": weeks, "expected_week": expected,
             "pms": pms(wb["01_PMS Dashboard"]),
             "activities": activities(wb["04_Regular Activities"]) if "04_Regular Activities" in wb.sheetnames else {}}
        p["signals"] = signals_for(p)
        p["score"] = score(p["signals"])
        p["overall"] = G if p["score"] >= 75 else (A if p["score"] >= 45 else R)
        # a red on a critical signal never shows as green overall
        if p["overall"] == G and any(s["status"] == R for s in p["signals"] if s["key"] in CRITICAL):
            p["overall"] = A
        p["actions"] = actions_for(p)
        props.append(p)
    props.sort(key=lambda p: p["score"] or 0)
    data = {"generated_at": datetime.now(IST).strftime("%Y-%m-%d %H:%M IST"),
            "properties": props, "tasks": task_list}
    json.dump(data, open(out_path, "w"), ensure_ascii=False, indent=1, default=str)
    print(f"wrote {len(props)} properties, {len(task_list)} tasks -> {out_path}")
    for p in props:
        print(p["city"], p["score"], p["overall"], [(s["key"], s["status"]) for s in p["signals"]])


if __name__ == "__main__":
    main()
