#!/usr/bin/env python3
"""
Napraw BiuroHandlowe w pliku Excel
──────────────────────────────────
Mapuje UworzonePrzez → prawidłowe BiuroHandlowe na podstawie BH_handlowcy.xlsx.
Obsługuje: odwróconą kolejność imię/nazwisko, zmianę nazwiska (Kolumna1),
podwójne nazwiska z myślnikiem, czeskie znaki, brak diakrytyków.

Użycie:
  python napraw_biura_handlowe.py dane.xlsx
  python napraw_biura_handlowe.py dane.xlsx --mapping BH_handlowcy.xlsx
  python napraw_biura_handlowe.py dane.xlsx --output dane_poprawione.xlsx
"""
import pandas as pd
import unicodedata
import os, sys, glob

# ── Argumenty ──
data_file = None; mapping_file = None; output_file = None
args = sys.argv[1:]
i = 0
while i < len(args):
    if args[i] == '--mapping' and i+1 < len(args): mapping_file = args[i+1]; i += 2
    elif args[i] == '--output' and i+1 < len(args): output_file = args[i+1]; i += 2
    elif not args[i].startswith('-'): data_file = args[i]; i += 1
    else: i += 1

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

if not data_file:
    xlsx = [f for f in glob.glob(os.path.join(SCRIPT_DIR, '*.xlsx'))
            if 'BH_handlowcy' not in f and 'poprawione' not in f and not os.path.basename(f).startswith('~$')]
    if xlsx: data_file = xlsx[0]; print(f"Auto: {os.path.basename(data_file)}")
    else: print("ERROR: Podaj plik Excel"); sys.exit(1)

if not mapping_file:
    for c in ['BH_handlowcy.xlsx','bh_handlowcy.xlsx']:
        for d in [SCRIPT_DIR, os.path.dirname(data_file)]:
            fp = os.path.join(d, c)
            if os.path.exists(fp): mapping_file = fp; break
        if mapping_file: break

if not mapping_file: print("ERROR: Brak BH_handlowcy.xlsx"); sys.exit(1)
if not output_file:
    base, ext = os.path.splitext(data_file)
    output_file = f"{base}_poprawione{ext}"

print(f"Dane:    {os.path.basename(data_file)}")
print(f"Mapping: {os.path.basename(mapping_file)}")
print(f"Wynik:   {os.path.basename(output_file)}")

# ── Funkcje normalizacji ──
def strip_diacritics(s):
    """Remove all diacritics: ą→a, ó→o, č→c, á→a etc."""
    nfkd = unicodedata.normalize('NFKD', s)
    return ''.join(c for c in nfkd if not unicodedata.combining(c))

def tokenize(name):
    """Split name into normalized tokens (handles hyphens, spaces)."""
    s = str(name).strip().lower()
    s = ' '.join(s.split())
    # Split on spaces AND hyphens
    tokens = set()
    for part in s.split():
        tokens.add(part)
        for sub in part.split('-'):
            if sub: tokens.add(sub)
    # Also add ascii versions
    for t in list(tokens):
        tokens.add(strip_diacritics(t))
    return tokens

def name_variants(name):
    """Generate all matching variants of a name."""
    s = str(name).strip().lower()
    s = ' '.join(s.split())
    sa = strip_diacritics(s)
    parts = s.split()
    parts_a = sa.split()
    
    variants = {s, sa}
    # Reversed order
    if len(parts) == 2:
        variants.add(f"{parts[1]} {parts[0]}")
    if len(parts_a) == 2:
        variants.add(f"{parts_a[1]} {parts_a[0]}")
    # Handle hyphenated: "Zagała-Szymkiewicz Anna" → also try "Zagała Anna"
    for p in parts:
        if '-' in p:
            for sub in p.split('-'):
                for other in parts:
                    if other != p:
                        variants.add(f"{sub} {other}")
                        variants.add(f"{other} {sub}")
                        variants.add(f"{strip_diacritics(sub)} {strip_diacritics(other)}")
                        variants.add(f"{strip_diacritics(other)} {strip_diacritics(sub)}")
    return variants

# ── Buduj mapping ──
print("\nBuduję mapping handlowiec → biuro handlowe...")
bh_df = pd.read_excel(mapping_file)
bh_df = bh_df.dropna(subset=['Oddział', 'Name, surname'])

name_to_bh = {}  # exact variant → BH
token_to_bh = {} # frozenset of tokens → BH (for fuzzy)

for _, row in bh_df.iterrows():
    oddz = str(row['Oddział']).strip()
    names_to_add = [str(row['Name, surname']).strip()]
    # Also add Kolumna1 (alternative name / name after marriage)
    if 'Kolumna1' in row.index and pd.notna(row.get('Kolumna1')) and str(row['Kolumna1']).strip() not in ['?','']:
        names_to_add.append(str(row['Kolumna1']).strip())
    
    for name in names_to_add:
        for v in name_variants(name):
            name_to_bh[v] = oddz
        # Token-based matching (for fuzzy)
        toks = tokenize(name)
        if len(toks) >= 2:
            token_to_bh[frozenset(toks)] = oddz

print(f"  {len(bh_df)} handlowców, {len(name_to_bh)} wariantów nazw, {len(token_to_bh)} zestawów tokenów")

def find_bh(name):
    """Find BH with multi-level matching."""
    if pd.isna(name): return None
    # Level 1: Exact variant match
    for v in name_variants(name):
        if v in name_to_bh: return name_to_bh[v]
    # Level 2: Token overlap (at least 2 common tokens including ascii)
    query_tokens = tokenize(name)
    best_match = None
    best_overlap = 0
    for known_tokens, bh in token_to_bh.items():
        overlap = len(query_tokens & known_tokens)
        if overlap >= 2 and overlap > best_overlap:
            best_overlap = overlap
            best_match = bh
    return best_match

# ── Wczytaj dane ──
print(f"\nWczytuję {os.path.basename(data_file)}...")
xls = pd.ExcelFile(data_file)
df = None; sheet_name = None
for sheet in xls.sheet_names:
    candidate = pd.read_excel(xls, sheet_name=sheet)
    if len(candidate.columns) >= 6 and (df is None or len(candidate.columns) > len(df.columns)):
        df = candidate; sheet_name = sheet
print(f"  Arkusz: '{sheet_name}' ({len(df)} wierszy)")

uzp_col = None
for c in df.columns:
    if c.lower().strip() in ['utworzoneprzez','uworzoneprzez']: uzp_col = c; break
if not uzp_col: print("ERROR: Brak kolumny UworzonePrzez!"); sys.exit(1)

bh_col = None
for c in df.columns:
    if c.lower().strip() in ['biurohandlowe','biuro handlowe']: bh_col = c; break

# ── Mapuj ──
print("Mapuję handlowców...")
df['_bh_new'] = df[uzp_col].apply(find_bh)

matched = df['_bh_new'].notna().sum()
total = df[uzp_col].notna().sum()
print(f"  Dopasowano: {matched}/{total} ({matched/total*100:.1f}%)")

unmatched = df.loc[df[uzp_col].notna() & df['_bh_new'].isna(), uzp_col].unique()
if len(unmatched) > 0:
    print(f"  Niedopasowani ({len(unmatched)}):")
    for u in unmatched[:15]:
        cnt = (df[uzp_col] == u).sum()
        print(f"    '{u}' ({cnt} wierszy)")

if bh_col:
    has_both = df[bh_col].notna() & df['_bh_new'].notna()
    same = df.loc[has_both, bh_col].str.strip().str.upper() == df.loc[has_both, '_bh_new'].str.strip().str.upper()
    print(f"\n  Zgodne: {same.sum()} ({same.mean()*100:.1f}%) | Zmienione: {(~same).sum()} ({(~same).mean()*100:.1f}%)")
    changes = df.loc[has_both & ~same].groupby([bh_col,'_bh_new']).size().reset_index(name='n').sort_values('n',ascending=False)
    for _,r in changes.head(10).iterrows():
        print(f"    {str(r[bh_col]):25s} → {r['_bh_new']:25s}: {r['n']:>5}")

# ── Zastosuj ──
if bh_col:
    original = df[bh_col].copy()
    mask = df['_bh_new'].notna()
    df.loc[mask, bh_col] = df.loc[mask, '_bh_new']
    df.loc[df[bh_col].isna(), bh_col] = original.loc[df[bh_col].isna()]
else:
    df['BiuroHandlowe'] = df['_bh_new'].fillna('NIEPRZYPISANE')
    bh_col = 'BiuroHandlowe'
df.drop(columns=['_bh_new'], inplace=True)

print(f"\n{'='*60}")
print(f"NOWY STAN '{bh_col}':")
for bh_name, cnt in df[bh_col].value_counts().items():
    ton = df.loc[df[bh_col]==bh_name,'waga'].sum()/1000 if 'waga' in df.columns else 0
    print(f"  {str(bh_name):25s}: {cnt:>6} wierszy | {ton:>8,.0f} t")

print(f"\nZapisuję: {output_file}")
df.to_excel(output_file, index=False, sheet_name=sheet_name or 'Arkusz1')
print(f"✅ Gotowe! {len(df)} wierszy, '{bh_col}' poprawione.")
