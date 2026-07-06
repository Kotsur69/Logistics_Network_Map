#!/usr/bin/env python3
"""Generates a small synthetic dataset matching the schema generate_map.py expects.
All company/customer/salesperson names are fictional demo data."""
import pandas as pd
import random
from datetime import datetime, timedelta

random.seed(42)

WAREHOUSES = [
    ("WARSZAWA", "00-841"),
    ("WROCŁAW", "50-001"),
    ("POZNAŃ", "60-001"),
]

DESTINATIONS = [
    ("KRAKÓW", "31-001"), ("GDAŃSK", "80-001"), ("GDYNIA", "81-001"),
    ("ŁÓDŹ", "90-001"), ("KATOWICE", "40-001"), ("SZCZECIN", "70-001"),
    ("LUBLIN", "20-001"), ("BIAŁYSTOK", "15-001"), ("BYDGOSZCZ", "85-001"),
    ("OLSZTYN", "10-001"), ("OPOLE", "45-001"), ("KIELCE", "25-001"),
    ("RADOM", "26-001"), ("TORUŃ", "87-001"), ("RZESZÓW", "35-001"),
]

SALESPEOPLE = [
    ("Jan Kowalski", "BH Warszawa"),
    ("Anna Nowak", "BH Wrocław"),
    ("Piotr Wiśniewski", "BH Poznań"),
    ("Katarzyna Wójcik", "BH Warszawa"),
]

MONTHS = ["2025-01", "2025-02", "2025-03"]

rows = []
trasa_no = 1
zlecenie_no = 100000

for month in MONTHS:
    for wh_city, wh_zip in WAREHOUSES:
        n_trips = random.randint(10, 16)
        for _ in range(n_trips):
            dest_city, dest_zip = random.choice(DESTINATIONS)
            salesperson, biuro = random.choice(SALESPEOPLE)
            waga_kg = random.choice([500, 1200, 3000, 5400, 8000, 12000, 18000, 24000])
            distance = random.randint(80, 620)
            day = random.randint(1, 27)
            hour = random.randint(6, 18)
            load_dt = datetime.strptime(f"{month}-{day:02d}", "%Y-%m-%d") + timedelta(hours=hour)

            rows.append({
                "cityName": wh_city,
                "zipcode": wh_zip,
                "to_cityName": dest_city,
                "to_zipcode": dest_zip,
                "waga": waga_kg,
                "załadunek": f"Magazyn {wh_city.title()}",
                "NrSekcji = nr trasy": f"TRASA-{trasa_no:04d}",
                "NrZlecenia": f"ZL-{zlecenie_no}",
                "TripDistance": distance,
                "UworzonePrzez": salesperson,
                "BiuroHandlowe": biuro,
                "CzasZaładunku": load_dt.strftime("%d.%m.%Y %H:%M:%S"),
            })
            trasa_no += 1
            zlecenie_no += 1

df = pd.DataFrame(rows)
out_path = "demo_data.xlsx"
df.to_excel(out_path, index=False, sheet_name="dane")
print(f"Wrote {len(df)} rows to {out_path}")
