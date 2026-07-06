# Logistics Network Map

Turns raw shipment export files (Excel) into a single interactive HTML map of a
transport network: routes between warehouses and destination cities, tonnage,
top customers, sales-office breakdowns, and month-by-month filtering — all in
one self-contained file you can open in a browser, no server required.

## How it works

`generate_map.py` reads one or more `.xlsx` shipment exports, normalizes city
names and postal codes, groups shipments into routes, fetches real road paths
from the public [OSRM](http://project-osrm.org/) routing API (falling back to
straight lines if OSRM is unreachable, and caching results in `osrm_cache.json`
so repeat runs don't re-fetch), and writes everything into
`logistics_network_map.html`.

Required columns in the input file:

| Column | Meaning |
|---|---|
| `cityName` / `zipcode` | Origin city and postal code |
| `to_cityName` / `to_zipcode` | Destination city and postal code |
| `waga` | Shipment weight, in kg |
| `załadunek` | Warehouse / loading point label |

Optional columns the script picks up if present: `NrSekcji = nr trasy` (route
number), `NrZlecenia` (order number, used for counts), `TripDistance`,
`CzasZaładunku` (loading timestamp, enables month filtering), `UworzonePrzez`
(salesperson — mapped to a sales office via `BH_handlowcy.xlsx` if that file
exists alongside the script), or a direct `BiuroHandlowe` column as a
shortcut. Column names are matched flexibly (case-insensitive, common
aliases), so slightly different exports usually work without edits.

## Try it with demo data

The real shipment files used in production contain live customer data and
aren't in this repo. `demo_data/` has a small synthetic dataset with the same
column layout, plus the script that generated it (`generate_demo_data.py`),
so you can see exactly how it was built. `demo_data/logistics_network_map_demo.html`
is the pre-built output — open it directly in a browser to see the program
working without installing anything.

To regenerate it yourself:

```bash
pip install pandas numpy openpyxl
python generate_map.py demo_data/demo_data.xlsx
```

This writes `logistics_network_map.html` in the script directory. Open it in
a browser.

## Using it with your own data

Drop your `.xlsx` export(s) next to `generate_map.py` and run the script with
no arguments — it picks up every `.xlsx` file in the directory automatically
(add specific filenames as arguments to restrict it to just those files):

```bash
python generate_map.py
```

If your `UworzonePrzez` names need mapping to sales offices, add a
`BH_handlowcy.xlsx` file with `Oddział` and `Name, surname` columns next to
the script, or fix an existing export in place with:

```bash
python "napraw_biura_handlowe (1).py" your_data.xlsx
```

## Notes

- `safe_version.py` is an earlier snapshot of `generate_map.py`, kept for
  reference.
- If you see a `UnicodeEncodeError` on the console when the script finishes
  (Windows-only, non-UTF-8 terminals), the HTML output was still written
  successfully — it's just the final status print failing to render an emoji.
  Run with `set PYTHONIOENCODING=utf-8` first if you want a clean run.
- Real shipment `.xlsx` files and the HTML generated from them are excluded
  from this repo via `.gitignore` — see `demo_data/` to evaluate the tool
  without needing them.
