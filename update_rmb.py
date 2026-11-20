#!/usr/bin/env python3
"""Merge a fresh RateMyBiz snapshot into rmb.json.

Usage: python3 update_rmb.py snapshot.json rmb.json

snapshot.json has the same shape as rmb.json's "hotels" list:
  {"checked_at": "<ISO time>", "hotels": [ {...}, ... ]}
Each hotel needs "key", "name", "city", "status"; connected hotels also need
rating, reviews, new_month, month, goal_rating, goal_needed, unreplied,
unreplied_note, local_rank, local_of, review_score, maps{top3,top10,top20,
beyond,best,keyword,radius_km,data_month,updated}, ai, complaint.

The script replaces "hotels" and "checked_at", and adds (or replaces) one
history entry for the snapshot's date so the weekly trend charts grow.
"""
import json
import sys

NUM_KEYS = ("rating", "reviews", "new_month", "local_rank", "local_of", "review_score")
MAP_KEYS = ("top3", "top10", "top20", "beyond")


def check(h):
    for k in ("key", "name", "city", "status"):
        if not h.get(k):
            raise SystemExit(f"hotel missing {k}: {h}")
    if h["status"] != "connected":
        return
    for k in NUM_KEYS:
        if not isinstance(h.get(k), (int, float)):
            raise SystemExit(f"{h['key']}: {k} must be a number")
    if not 1 <= h["rating"] <= 5:
        raise SystemExit(f"{h['key']}: rating out of range")
    m = h.get("maps") or {}
    if sum(int(m.get(k, 0)) for k in MAP_KEYS) not in (0, 9):
        raise SystemExit(f"{h['key']}: maps grid must add up to 9 points")
    u = h.get("unreplied")
    if u is not None and not isinstance(u, int):
        raise SystemExit(f"{h['key']}: unreplied must be a whole number or null")


def main(snap_path, out_path):
    snap = json.load(open(snap_path))
    hotels = snap["hotels"]
    for h in hotels:
        check(h)
    if not any(h["status"] == "connected" for h in hotels):
        raise SystemExit("no connected hotels in snapshot; refusing to overwrite")
    data = json.load(open(out_path))
    data["checked_at"] = snap["checked_at"]
    data["hotels"] = hotels
    day = snap["checked_at"][:10]
    entry = {"date": day, "hotels": {}}
    for h in hotels:
        if h["status"] == "connected":
            m = h.get("maps") or {}
            entry["hotels"][h["key"]] = {
                "rating": h["rating"], "reviews": h["reviews"],
                "unreplied": h.get("unreplied"),
                "maps_top10": int(m.get("top3", 0)) + int(m.get("top10", 0)),
            }
    hist = [e for e in data.get("history", []) if e["date"] != day]
    hist.append(entry)
    hist.sort(key=lambda e: e["date"])
    data["history"] = hist[-104:]
    with open(out_path, "w") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    n = len(entry["hotels"])
    print(f"wrote {n} connected hotels, {len(hotels)} total, {len(data['history'])} weeks of history")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit(__doc__)
    main(sys.argv[1], sys.argv[2])
