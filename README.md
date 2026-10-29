# Trevon Hotels — Google Business Profile dashboard

Static page (`index.html` + `data.json`) served by GitHub Pages.

* `build_gmb.py <folder of .xlsx> data.json` rebuilds the data from the agency
  workbooks in the Google Drive folder "Trevon Hotels GMB data".
* A scheduled Claude task runs it every morning and pushes the new `data.json`.
* Traffic-light rules are in `build_gmb.py` (`signals_for`) and shown on the page.
