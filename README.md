# Trevon Hotels — Google Business Profile dashboard

Static page (`index.html` + `rmb.json`) served by GitHub Pages.

* Data comes from RateMyBiz (get.ratemybiz.xyz): Google rating, reviews, unreplied
  reviews, review-score rank against nearby hotels, and the Google Maps rank tracker.
* A scheduled Claude task reads RateMyBiz every Monday, writes a snapshot and runs
  `python3 update_rmb.py snapshot.json rmb.json`, which replaces the hotel figures and
  adds one point to `history` for the weekly trend charts.
