#!/usr/bin/env python3
"""
Logistics Flow Map Generator — Multi-file version
Loads all .xlsx files from the script directory (or given as arguments),
generates ONE interactive HTML with tab-switching between datasets.
"""
import pandas as pd, numpy as np, json, random, os, sys, glob, re

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

ZIP2 = {
    '00':(52.230,21.010),'01':(52.250,20.950),'02':(52.200,20.930),'03':(52.270,21.050),
    '04':(52.200,21.080),'05':(52.150,20.850),'06':(52.730,20.100),'07':(52.550,21.500),
    '08':(51.650,21.850),'09':(52.430,19.800),'10':(53.780,20.480),'11':(54.050,20.500),
    '12':(53.800,20.950),'13':(53.550,19.700),'14':(53.650,20.150),'15':(53.130,23.150),
    '16':(53.550,23.500),'17':(52.950,22.750),'18':(53.450,22.500),'19':(53.850,22.350),
    '20':(51.250,22.550),'21':(51.450,22.050),'22':(50.800,22.550),'23':(50.700,23.300),
    '24':(51.050,21.800),'25':(50.870,20.630),'26':(51.350,20.550),'27':(50.950,21.420),
    '28':(50.650,21.350),'29':(51.050,20.150),'30':(50.070,19.950),'31':(50.100,20.050),
    '32':(49.950,19.850),'33':(49.850,20.550),'34':(49.700,19.450),'35':(50.000,22.000),
    '36':(50.100,22.550),'37':(50.550,21.550),'38':(49.700,20.850),'39':(50.350,21.800),
    '40':(50.250,19.000),'41':(50.300,18.950),'42':(50.300,19.150),'43':(49.800,19.050),
    '44':(50.100,19.300),'45':(50.680,17.930),'46':(50.650,17.550),'47':(50.450,18.150),
    '48':(50.550,18.550),'49':(50.750,17.550),'50':(51.100,17.030),'51':(51.050,17.050),
    '52':(51.150,16.850),'53':(51.100,16.950),'54':(51.100,17.100),'55':(51.300,16.850),
    '56':(51.200,16.500),'57':(50.750,16.450),'58':(50.700,16.650),'59':(51.450,16.150),
    '60':(52.400,16.900),'61':(52.450,16.850),'62':(52.300,17.150),'63':(51.750,17.800),
    '64':(52.400,16.550),'65':(52.000,15.500),'66':(52.400,15.050),'67':(51.950,15.500),
    '68':(51.600,15.200),'69':(51.950,14.800),'70':(53.430,14.550),'71':(53.350,14.650),
    '72':(53.500,15.300),'73':(53.750,15.350),'74':(53.600,14.700),'75':(54.180,15.600),
    '76':(54.200,16.200),'77':(53.750,16.100),'78':(54.150,16.550),'80':(54.350,18.650),
    '81':(54.520,18.500),'82':(54.050,18.550),'83':(54.250,18.200),'84':(54.450,17.750),
    '85':(53.120,18.000),'86':(53.400,18.600),'87':(53.050,18.600),'88':(53.350,18.150),
    '89':(53.250,17.550),'90':(51.770,19.450),'91':(51.800,19.500),'92':(51.700,19.400),
    '93':(51.750,19.500),'94':(51.750,19.400),'95':(51.600,19.250),'96':(51.900,19.800),
    '97':(51.550,19.950),'98':(51.400,19.100),'99':(51.400,19.950),
}
ZIP3 = {
    '002':(52.230,21.010),'010':(52.250,20.900),'015':(52.260,20.920),
    '024':(52.210,20.940),'031':(52.290,21.060),'041':(52.200,21.090),
    '300':(50.060,19.940),'310':(50.070,19.900),'317':(50.060,20.010),
    '400':(50.260,19.020),'410':(50.290,18.970),'413':(50.320,18.950),
    '500':(51.110,17.040),'600':(52.410,16.930),'800':(54.360,18.640),
    '810':(54.520,18.530),'900':(51.780,19.460),
}
CITY_COORDS = {
    'WARSZAWA':(52.230,21.012),'KRAKÓW':(50.065,19.945),'GDAŃSK':(54.352,18.647),
    'GDYNIA':(54.519,18.531),'WROCŁAW':(51.110,17.032),'POZNAŃ':(52.407,16.930),
    'ŁÓDŹ':(51.775,19.461),'KATOWICE':(50.265,19.023),'SZCZECIN':(53.429,14.553),
    'LUBLIN':(51.248,22.568),'BIAŁYSTOK':(53.132,23.169),'BYDGOSZCZ':(53.122,18.001),
    'OLSZTYN':(53.778,20.480),'OPOLE':(50.672,17.926),'KIELCE':(50.866,20.629),
    'RADOM':(51.402,21.147),'TORUŃ':(53.014,18.598),'SOSNOWIEC':(50.287,19.104),
    'RZESZÓW':(50.042,21.999),'GLIWICE':(50.294,18.671),'ZABRZE':(50.325,18.786),
    'BIELSKO-BIAŁA':(49.822,19.046),'RYBNIK':(50.102,18.541),'TYCHY':(50.132,18.978),
    'DĄBROWA GÓRNICZA':(50.332,19.200),'ELBLĄG':(54.152,19.404),'PŁOCK':(52.547,19.706),
    'TARNÓW':(50.013,20.986),'KOSZALIN':(54.194,16.172),'KALISZ':(51.761,18.093),
    'LEGNICA':(51.210,16.156),'GRUDZIĄDZ':(53.484,18.754),'SŁUPSK':(54.464,17.029),
    'JAWORZNO':(50.205,19.274),'JASTRZĘBIE ZDRÓJ':(49.948,18.593),
    'SIEDLCE':(52.167,22.290),'MYSŁOWICE':(50.232,19.167),'KONIN':(52.223,18.251),
    'PIŁA':(53.151,16.738),'SUWAŁKI':(54.112,22.931),'INOWROCŁAW':(52.796,18.260),
    'LUBIN':(51.401,16.202),'MIELEC':(50.287,21.425),'STALOWA WOLA':(50.583,22.054),
    'TARNOBRZEG':(50.573,21.679),'TARNOWSKIE GÓRY':(50.445,18.862),'CHORZÓW':(50.297,18.955),
    'LĘBORK':(54.539,17.750),'ZAMOŚĆ':(50.723,23.252),'WŁOCŁAWEK':(52.660,19.068),
    'EŁK':(53.828,22.345),'ŻORY':(50.049,18.702),'PRZEMYŚL':(49.784,22.768),
    'NOWY SĄCZ':(49.618,20.714),'RUDA ŚLĄSKA':(50.260,18.856),'OŚWIĘCIM':(50.034,19.211),
    'SKARŻYSKO-KAMIENNA':(51.113,20.863),'PRUSZCZ GDAŃSKI':(54.262,18.636),
    'KĘDZIERZYN-KOŹLE':(50.349,18.207),'KOŁOBRZEG':(54.176,15.583),
    'WIELUŃ':(51.221,18.570),'STARGARD':(53.337,15.049),'LESZNO':(51.842,16.575),
    'BYTOM':(50.348,18.912),'BĘDZIN':(50.325,19.128),'ZAWIERCIE':(50.490,19.435),
    'OLKUSZ':(50.281,19.564),'WYSOKIE MAZOWIECKIE':(52.920,22.520),
    'NAREW':(52.898,23.518),'ZAGÓRÓW':(52.166,17.889),'BUDZYŃ':(52.895,16.991),
    'POŁCHOWO':(54.668,18.268),'BYSŁAW':(53.608,18.202),'LIPSKO':(51.162,21.654),
    'SKAWINA':(49.976,19.828),'MEŁGIEW':(51.175,22.845),'KOŚCIERZYNA':(54.122,17.979),
    # Czech/Slovak cities
    'OSTRAVA':(49.821,18.263),'FRÝDEK-MÍSTEK':(49.688,18.353),'FRYČOVICE':(49.683,18.224),
    'OSTRAVA - KUNČICE':(49.802,18.281),'ŠENOV U OSTRAVY':(49.791,18.373),
    'OSTRAVA - VÍTKOVICE':(49.812,18.253),'FRYŠTÁK':(49.285,17.693),
    'OLOMOUC':(49.594,17.251),'PROSTĚJOV':(49.472,17.111),'PŘEROV':(49.455,17.451),
    'ŠTERNBERK':(49.731,17.299),'OPAVA':(49.938,17.905),'BRNO':(49.195,16.608),
    'ZLÍN':(49.227,17.667),'TŘINEC':(49.678,18.672),'BOHUMÍN':(49.904,18.358),
    'NOVÝ JIČÍN':(49.594,18.010),'KARVINÁ':(49.854,18.542),'HAVÍŘOV':(49.780,18.437),
    'KROMĚŘÍŽ':(49.298,17.393),'VALAŠSKÉ MEZIŘÍČÍ':(49.472,17.971),
    'ČESKÝ TĚŠÍN':(49.746,18.628),'KOPŘIVNICE':(49.599,18.145),
    'FRÝDEK MÍSTEK':(49.688,18.353),'FRYDEK-MISTEK':(49.688,18.353),
    'ŠENOV':(49.791,18.373),'PETŘVALD':(49.828,18.386),
    'FRÝDLANT NAD OSTRAVICÍ':(49.590,18.361),'PASKOV':(49.726,18.289),
    'VRATIMOV':(49.770,18.310),'KLIMKOVICE':(49.788,18.127),
    'HLUČÍN':(49.898,18.192),'BÍLOVEC':(49.756,17.862),
    'UHERSKÉ HRADIŠTĚ':(49.070,17.460),'UHERSKÝ BROD':(49.025,17.650),
    # Czech cities (continued)
    'HRANICE':(49.548,17.735),'LHOTKA NAD BEČVOU':(49.480,18.020),
    'LITOVEL':(49.701,17.076),'OTROKOVICE':(49.210,17.531),
    'MORAVSKÉ BUDĚJOVICE':(49.052,15.808),'JESENÍK':(50.229,17.204),
    'PŘEROV':(49.455,17.451),'VELKÉ ALBRECHTICE':(49.730,17.990),
    'HUSTOPEČE':(48.940,16.738),'DĚTMAROVICE':(49.864,18.460),
    'PROSTĚJOV':(49.472,17.111),'PASKOV':(49.726,18.289),
    'ŠTERNBERK':(49.731,17.299),'KROMĚŘÍŽ':(49.298,17.393),
    'OSTRAVA - KUNČIČKY':(49.802,18.300),'UHŘÍNOV':(49.419,16.124),
    # Slovak cities
    'NÁMESTOVO':(49.408,19.484),'KRÁSNO NAD KYSUCOU':(49.394,18.830),
    'ŽILINA':(49.224,18.739),'POVAŽSKÁ BYSTRICA':(49.122,18.432),
    'PREŠOV':(48.998,21.239),'KOŠICE':(48.716,21.261),
    'VRANOV NAD TOPĽOU':(48.886,21.682),'BIDOVCE':(48.726,21.395),
    'RUŽOMBEROK':(49.078,19.307),'VEĽKÝ ŠARIŠ':(49.040,21.190),
    'STARÁ ĽUBOVŇA':(49.301,20.688),'DRIENOVEC':(48.620,20.928),
    'VYSOKÁ NAD KYSUCOU':(49.404,18.778),'TRNAVA':(48.377,17.587),
    'PIEŠŤANY':(48.595,17.826),'MALACKY':(48.436,17.024),
    'TRENČÍN':(48.895,18.044),'TRENCÍN':(48.895,18.044),
    'LIČARTOVCE':(48.950,21.335),'SVIT':(49.058,20.198),
}

def get_coords_from_zip(zipcode):
    if not isinstance(zipcode, str) or len(zipcode)<2: return None
    zr=zipcode.strip()
    # Only use Polish ZIP tables for Polish format (XX-XXX)
    import re as _re
    if _re.match(r'^\d{2}-\d{3}$',zr):
        z=zr.replace('-','')
        if z[:3] in ZIP3: return ZIP3[z[:3]]
        if z[:2] in ZIP2: return ZIP2[z[:2]]
    # Foreign zips — don't use Polish ZIP tables
    return None

def normalize_city(name):
    s=str(name).strip().upper(); s=re.sub(r'\s+',' ',s); s=re.sub(r'[.,;:!?/\\]+$','',s).strip()
    s=re.sub(r'\s+K[./].*$','',s); s=re.sub(r'\s+GM\.?.*$','',s)
    s=re.sub(r'\s+N[./].*$','',s); s=re.sub(r'\s+GN\.?.*$','',s)
    M={
        'DĄBROWA GÓRNICZA':['DABROWA GÓRNICZA','DABROWA GORNICZA','DĄBROWA GORNICZA'],
        'BIELSKO-BIAŁA':['BIELSKO BIAŁA','BIELSKO BIALA','BIELSKO - BIAŁA','BIELSKO-BIALA'],
        'KĘDZIERZYN-KOŹLE':['KĘDZIERZYN KOŹLE','KĘDZIERZYN KOZLE','KEDZIERZYN-KOZLE','KEDZIERZYN KOZLE'],
        'JELCZ-LASKOWICE':['JELCZ LASKOWICE'],'SKARŻYSKO-KAMIENNA':['SKARŻYSKO KAMIENNA'],
        'JASTRZĘBIE ZDRÓJ':['JASTRZĘBIE-ZDRÓJ','JASTRZEBIE ZDROJ','JASTRZĘBIE ZDROJ'],
        'RUDA ŚLĄSKA':['RUDA SLASKA','RUDA  ŚLĄSKA'],
        'OSTROWIEC ŚWIĘTOKRZYSKI':['OSTROWIEC SWIĘTOKRZYSKI','OSTROWIEC SWIETOKRZYSKI','OSTROWIEC'],
        'TARNOWSKIE GÓRY':['TARNOWSKIE  GÓRY','TARNOWSKIE  GORY','TARNOWSKIE GORY'],
        'GORZÓW WIELKOPOLSKI':['GORZÓW WLKP.','GORZÓW WLKP','GORZOW WIELKOPOLSKI'],
        'ŁÓDŹ':['ŁÓDZ','ŁÓDŻ','LODZ'],'CZĘSTOCHOWA':['CZESTOCHOWA'],
        'GDAŃSK':['GDANSK'],'BIAŁYSTOK':['BIALYSTOK','BIAŁSTOK'],
        'KRAKÓW':['KRAKOW'],'POZNAŃ':['POZNAN'],'WROCŁAW':['WROCLAW'],
        'TORUŃ':['TORUN'],'PŁOCK':['PLOCK'],'WŁOCŁAWEK':['WLOCLAWEK'],
        'PRUSZCZ GDAŃSKI':['PRUSZCZ GDANSKI'],
        'CZECHOWICE-DZIEDZICE':['CZECHOWICE DZIEDZICE'],
    }
    for c,vs in M.items():
        if s==c or s in vs: return c
    return s

def process_file(filepath):
    fname=os.path.basename(filepath); label=os.path.splitext(fname)[0]
    print(f"  [{label}] Loading...")
    # Try all sheets — pick the one with the most columns (skip pivot summaries)
    xls=pd.ExcelFile(filepath)
    df=None
    for sheet in xls.sheet_names:
        candidate=pd.read_excel(xls, sheet_name=sheet)
        if len(candidate.columns)>=6 and (df is None or len(candidate.columns)>len(df.columns)):
            df=candidate
            used_sheet=sheet
    if df is None:
        df=pd.read_excel(filepath)
        used_sheet=xls.sheet_names[0]
    if len(xls.sheet_names)>1:
        print(f"    Arkusze: {xls.sheet_names} → używam '{used_sheet}' ({len(df.columns)} kolumn, {len(df)} wierszy)")
    # Flexible column mapping — find required columns by name variants
    col_map={'waga':None,'NrSekcji = nr trasy':None,'cityName':None,'to_cityName':None,
             'zipcode':None,'to_zipcode':None,'załadunek':None,'NrZlecenia':None,'TripDistance':None,'UworzonePrzez':None,'BiuroHandlowe':None,'to_stateName':None}
    for needed in col_map:
        if needed in df.columns:
            col_map[needed]=needed
        else:
            # Try case-insensitive + partial match
            nl=needed.lower()
            for c in df.columns:
                cl=c.lower().strip()
                if cl==nl or cl.replace(' ','_')==nl.replace(' ','_'):
                    col_map[needed]=c; break
            if col_map[needed] is None:
                # Common aliases
                aliases={'waga':['waga [kg]','waga_kg','weight','masa'],
                         'NrSekcji = nr trasy':['nrsekcji','nr_sekcji','nr trasy','nrsekcji = nr trasy','trip_id','nrsekcji'],
                         'cityName':['cityname','city_name','miasto','from_city'],
                         'to_cityName':['to_cityname','to_city_name','miasto_docelowe','dest_city'],
                         'zipcode':['zip','kod_pocztowy','from_zipcode'],
                         'to_zipcode':['to_zip','dest_zipcode','kod_docelowy'],
                         'załadunek':['zaladunek','warehouse','magazyn','loading_point'],
                         'TripDistance':['tripdistance','trip_distance','dystans','distance'],
                         'UworzonePrzez':['utworzoneprzez','handlowiec','salesperson','created_by'],
                         'BiuroHandlowe':['biurohandlowe','biuro_handlowe','sales_office'],
                         'to_stateName':['to_statename','to_state_name','województwo','voivodeship']}
                for alias in aliases.get(needed,[]):
                    for c in df.columns:
                        if c.lower().strip()==alias:
                            col_map[needed]=c; break
                    if col_map[needed]: break
    # Rename columns to standard names
    rename={}
    for std,actual in col_map.items():
        if actual and actual!=std: rename[actual]=std
    if rename: df=df.rename(columns=rename)
    # Validate required columns
    required=['waga','cityName','to_cityName','załadunek']
    missing=[c for c in required if c not in df.columns]
    if missing:
        # Print what we have for debugging
        print(f"    Kolumny w pliku: {df.columns.tolist()}")
        raise ValueError(f"Brak kolumn: {missing}")
    df=df.dropna(subset=['waga'])
    if 'NrSekcji = nr trasy' in df.columns: df=df.dropna(subset=['NrSekcji = nr trasy'])
    df['waga_ton']=df['waga']/1000
    # ── Fill missing to_stateName from zipcode ──
    if 'to_stateName' in df.columns and 'to_zipcode' in df.columns:
        PL_ZIP2_STATE={
            '00':'Woj. Mazowieckie','01':'Woj. Mazowieckie','02':'Woj. Mazowieckie',
            '03':'Woj. Mazowieckie','04':'Woj. Mazowieckie','05':'Woj. Mazowieckie',
            '06':'Woj. Mazowieckie','07':'Woj. Mazowieckie','08':'Woj. Mazowieckie',
            '09':'Woj. Mazowieckie','10':'Woj. Warmińsko-Mazurskie','11':'Woj. Warmińsko-Mazurskie',
            '12':'Woj. Warmińsko-Mazurskie','13':'Woj. Warmińsko-Mazurskie','14':'Woj. Warmińsko-Mazurskie',
            '15':'Woj. Podlaskie','16':'Woj. Podlaskie','17':'Woj. Podlaskie',
            '18':'Woj. Podlaskie','19':'Woj. Podlaskie','20':'Woj. Lubelskie',
            '21':'Woj. Lubelskie','22':'Woj. Lubelskie','23':'Woj. Lubelskie',
            '24':'Woj. Lubelskie','25':'Woj. Świętokrzyskie','26':'Woj. Mazowieckie',
            '27':'Woj. Świętokrzyskie','28':'Woj. Świętokrzyskie','29':'Woj. Świętokrzyskie',
            '30':'Woj. Małopolskie','31':'Woj. Małopolskie','32':'Woj. Małopolskie',
            '33':'Woj. Małopolskie','34':'Woj. Małopolskie','35':'Woj. Podkarpackie',
            '36':'Woj. Podkarpackie','37':'Woj. Podkarpackie','38':'Woj. Małopolskie',
            '39':'Woj. Podkarpackie','40':'Woj. Śląskie','41':'Woj. Śląskie',
            '42':'Woj. Śląskie','43':'Woj. Śląskie','44':'Woj. Śląskie',
            '45':'Woj. Opolskie','46':'Woj. Opolskie','47':'Woj. Opolskie',
            '48':'Woj. Opolskie','49':'Woj. Opolskie','50':'Woj. Dolnośląskie',
            '51':'Woj. Dolnośląskie','52':'Woj. Dolnośląskie','53':'Woj. Dolnośląskie',
            '54':'Woj. Dolnośląskie','55':'Woj. Dolnośląskie','56':'Woj. Dolnośląskie',
            '57':'Woj. Dolnośląskie','58':'Woj. Dolnośląskie','59':'Woj. Dolnośląskie',
            '60':'Woj. Wielkopolskie','61':'Woj. Wielkopolskie','62':'Woj. Wielkopolskie',
            '63':'Woj. Wielkopolskie','64':'Woj. Wielkopolskie','65':'Woj. Lubuskie',
            '66':'Woj. Lubuskie','67':'Woj. Lubuskie','68':'Woj. Lubuskie',
            '69':'Woj. Lubuskie','70':'Woj. Zachodniopomorskie','71':'Woj. Zachodniopomorskie',
            '72':'Woj. Zachodniopomorskie','73':'Woj. Zachodniopomorskie','74':'Woj. Zachodniopomorskie',
            '75':'Woj. Zachodniopomorskie','76':'Woj. Zachodniopomorskie','77':'Woj. Wielkopolskie',
            '78':'Woj. Zachodniopomorskie','80':'Woj. Pomorskie','81':'Woj. Pomorskie',
            '82':'Woj. Pomorskie','83':'Woj. Pomorskie','84':'Woj. Pomorskie',
            '85':'Woj. Kujawsko-Pomorskie','86':'Woj. Kujawsko-Pomorskie','87':'Woj. Kujawsko-Pomorskie',
            '88':'Woj. Kujawsko-Pomorskie','89':'Woj. Kujawsko-Pomorskie','90':'Woj. Łódzkie',
            '91':'Woj. Łódzkie','92':'Woj. Łódzkie','93':'Woj. Łódzkie','94':'Woj. Łódzkie',
            '95':'Woj. Łódzkie','96':'Woj. Łódzkie','97':'Woj. Łódzkie','98':'Woj. Łódzkie',
            '99':'Woj. Łódzkie',
        }
        CZ_ZIP1_STATE={
            '1':'Jihočeský kraj','2':'Plzeňský kraj','3':'Jihočeský kraj',
            '4':'Ústecký kraj','5':'Liberecký kraj','6':'Pardubický kraj',
            '7':'Olomoucký kraj','8':'Moravskoslezský kraj','9':'Vysočina',
        }
        # Also: 7xx = Zlínský/Olomoucký, depends on exact code
        CZ_ZIP2_STATE={
            '70':'Moravskoslezský kraj','71':'Moravskoslezský kraj','72':'Zlínský kraj',
            '73':'Moravskoslezský kraj','74':'Moravskoslezský kraj','75':'Olomoucký kraj',
            '76':'Zlínský kraj','77':'Olomoucký kraj','78':'Olomoucký kraj',
            '79':'Olomoucký kraj','60':'Jihomoravský kraj','61':'Jihomoravský kraj',
            '62':'Jihomoravský kraj','63':'Vysočina','64':'Jihomoravský kraj',
            '65':'Jihomoravský kraj','66':'Jihomoravský kraj','67':'Jihomoravský kraj',
            '68':'Jihomoravský kraj','69':'Jihomoravský kraj',
        }
        empty_mask=df['to_stateName'].isna()
        if True:  # Always run — also fix WRONG states, not just empty ones
            def _zip_to_state(z_raw):
                if not isinstance(z_raw,str) or len(z_raw.strip())<2: return None
                z_raw=z_raw.strip()
                # Polish zips: ALWAYS XX-XXX (with hyphen, 5+1=6 chars)
                import re as _re
                if _re.match(r'^\d{2}-\d{3}$',z_raw):
                    z2=z_raw[:2]
                    if z2 in PL_ZIP2_STATE: return PL_ZIP2_STATE[z2]
                    return None
                # Czech zips: XXX XX (3 digits, space, 2 digits)
                z=z_raw.replace(' ','').replace('-','')
                if len(z)==5 and z.isdigit():
                    if z[:2] in CZ_ZIP2_STATE: return CZ_ZIP2_STATE[z[:2]]
                    if z[0] in CZ_ZIP1_STATE: return CZ_ZIP1_STATE[z[0]]
                # Slovak zips: XXX XX or XX-X X — treat as foreign
                return 'Zagranica'
            # Apply to ALL rows — zip-based state is more reliable than to_stateName
            all_zip_states=df['to_zipcode'].apply(_zip_to_state)
            n_empty_before=df['to_stateName'].isna().sum()
            # Override: if zip gives a state, use it (fixes wrong states like Chybie=Pomorskie)
            mask_has_zip=all_zip_states.notna()
            n_overridden=(mask_has_zip & df['to_stateName'].notna() & (df['to_stateName']!=all_zip_states)).sum()
            df.loc[mask_has_zip,'to_stateName']=all_zip_states[mask_has_zip]
            n_filled=(df['to_stateName'].notna().sum())-(len(df)-n_empty_before)
            print(f"  [{label}] Województwa: {n_filled} uzupełnionych, {n_overridden} poprawionych błędnych")
    # ── BH (Biuro Handlowe) mapping from UworzonePrzez ──
    bh_data={}
    bh_mapping_file=os.path.join(SCRIPT_DIR,'BH_handlowcy.xlsx')
    if os.path.exists(bh_mapping_file):
        try:
            import unicodedata as _ud
            _bh=pd.read_excel(bh_mapping_file).dropna(subset=['Oddział','Name, surname'])
            def _nvs(n):
                s=str(n).strip().lower();s=' '.join(s.split());parts=s.split();vs={s}
                if len(parts)==2:vs.add(f"{parts[1]} {parts[0]}")
                for v in list(vs):
                    nfkd=_ud.normalize('NFKD',v);vs.add(''.join(c for c in nfkd if not _ud.combining(c)))
                return vs
            _nm={}
            for _,_r in _bh.iterrows():
                for v in _nvs(_r['Name, surname']):_nm[v]=str(_r['Oddział']).strip()
                if pd.notna(_r.get('Kolumna1')):
                    for v in _nvs(_r['Kolumna1']):_nm[v]=str(_r['Oddział']).strip()
            if 'UworzonePrzez' in df.columns:
                def _map_bh(name):
                    if pd.isna(name):return 'NIEPRZYPISANE'
                    for v in _nvs(name):
                        if v in _nm:return _nm[v]
                    return 'NIEPRZYPISANE'
                df['bh_mapped']=df['UworzonePrzez'].apply(_map_bh)
                bh_counts=df.groupby('bh_mapped')['waga_ton'].sum().sort_values(ascending=False)
                print(f"  [{label}] BH mapping: {(df['bh_mapped']!='NIEPRZYPISANE').sum()}/{len(df)} matched, {len(bh_counts)} biur")
            else:
                df['bh_mapped']='BRAK DANYCH'
        except Exception as e:
            print(f"  [{label}] BH mapping error: {e}")
            df['bh_mapped']='BRAK DANYCH'
    elif 'BiuroHandlowe' in df.columns:
        df['bh_mapped']=df['BiuroHandlowe'].fillna('NIEPRZYPISANE').str.strip()
        print(f"  [{label}] Using BiuroHandlowe column directly")
    else:
        df['bh_mapped']='BRAK DANYCH'
    df['from_norm']=df['cityName'].apply(normalize_city)
    df['to_norm']=df['to_cityName'].apply(normalize_city)
    # Disambiguate cities with same name in different regions (Osiek, Sędziszów etc)
    df['to_zip2']=df['to_zipcode'].apply(lambda x: str(x).replace('-','').replace(' ','')[:2] if pd.notna(x) and isinstance(x,str) else '')
    _city_zip2s=df.groupby('to_norm')['to_zip2'].apply(lambda x: x.unique()).to_dict()
    _ambiguous={c for c,z in _city_zip2s.items() if len(set(z))>1}
    if _ambiguous:
        print(f"  [{label}] {len(_ambiguous)} niejednoznacznych miast: {list(_ambiguous)[:8]}...")
        df.loc[df['to_norm'].isin(_ambiguous),'to_norm']=df.loc[df['to_norm'].isin(_ambiguous)].apply(
            lambda r: r['to_norm']+' ('+r['to_zip2']+')' if r['to_zip2'] else r['to_norm'], axis=1)
    random.seed(42)
    df['wh_clean']=df['załadunek'].str.strip()
    routes=df.groupby('NrSekcji = nr trasy').agg(
        from_city=('from_norm','first'),to_city=('to_norm','first'),
        from_zip=('zipcode','first'),to_zip=('to_zipcode','first'),
        total_ton=('waga_ton','sum'),trip_dist=('TripDistance','first'),
        wh_clean=('wh_clean','first'),
        bh=('bh_mapped','first'),
        state=('to_stateName','first') if 'to_stateName' in df.columns else ('wh_clean','first'),
    ).reset_index()
    total_km=routes['trip_dist'].sum(); total_ton=routes['total_ton'].sum()
    total_orders=df['NrZlecenia'].nunique() if 'NrZlecenia' in df.columns else 0
    zd=df.groupby('to_norm')['NrZlecenia'].nunique().to_dict() if 'NrZlecenia' in df.columns else {}
    zo=df.groupby('from_norm')['NrZlecenia'].nunique().to_dict() if 'NrZlecenia' in df.columns else {}
    zf={}
    if 'NrZlecenia' in df.columns:
        tmp=df.groupby(['wh_clean','to_norm'])['NrZlecenia'].nunique().reset_index()
        zf={(r['wh_clean'],r['to_norm']):r['NrZlecenia'] for _,r in tmp.iterrows()}
    cache={}
    def geo(cn,zp):
        # Key = city + zip2 prefix (handles duplicate city names like Osiek, Sędziszów)
        z2=str(zp).replace('-','').replace(' ','')[:2] if pd.notna(zp) and isinstance(zp,str) else ''
        key=cn+'|'+z2
        if key in cache: return cache[key]
        # 1. Try zipcode-based coords first (most precise)
        c=get_coords_from_zip(str(zp))
        if c: cache[key]=c; return c
        # 2. Known city name
        if cn in CITY_COORDS: cache[key]=CITY_COORDS[cn]; return cache[key]
        # 3. ASCII-normalized city name
        RP={'Ą':'A','Ć':'C','Ę':'E','Ł':'L','Ń':'N','Ó':'O','Ś':'S','Ź':'Z','Ż':'Z',
            'Á':'A','É':'E','Í':'I','Ó':'O','Ú':'U','Ý':'Y','Č':'C','Ď':'D','Ě':'E',
            'Ň':'N','Ř':'R','Š':'S','Ť':'T','Ů':'U','Ž':'Z'}
        a=cn
        for k,v in RP.items(): a=a.replace(k,v)
        for kn,co in CITY_COORDS.items():
            ka=kn
            for k,v in RP.items(): ka=ka.replace(k,v)
            if a==ka: cache[key]=co; return co
        return None
    def bz(g):
        vc=g.value_counts(); return vc.index[0] if len(vc)>0 else ''
    dz=routes.groupby('to_city')['to_zip'].apply(bz).to_dict()
    oz=routes.groupby('from_city')['from_zip'].apply(bz).to_dict()
    for _,r in routes[['to_city','to_zip']].drop_duplicates().iterrows(): geo(r['to_city'],r['to_zip'])
    for _,r in routes[['from_city','from_zip']].drop_duplicates().iterrows(): geo(r['from_city'],r['from_zip'])
    # Dest markers — GLOBAL (for unfiltered view)
    routes['to_zip2']=routes['to_zip'].apply(lambda x: str(x).replace('-','').replace(' ','')[:2] if pd.notna(x) else '')
    routes['geo_key']=routes['to_city']+'|'+routes['to_zip2']
    da=routes.groupby(['to_city','geo_key']).agg(total_ton=('total_ton','sum'),n_routes=('total_ton','count')).reset_index()
    dd=[]
    for _,r in da.iterrows():
        co=cache.get(r['geo_key'])
        if co: dd.append({'name':r['to_city'],'lat':co[0],'lon':co[1],'ton':round(r['total_ton'],1),
            'routes':int(r['n_routes']),'orders':int(zd.get(r['to_city'],0))})
    dd.sort(key=lambda x:x['ton'],reverse=True)  # no limit — all dests included
    # Dest markers — PER WAREHOUSE (for filtered view)
    zd_wh={}
    if 'NrZlecenia' in df.columns:
        tmp2=df.groupby(['wh_clean','to_norm'])['NrZlecenia'].nunique().reset_index()
        zd_wh={(r['wh_clean'],r['to_norm']):r['NrZlecenia'] for _,r in tmp2.iterrows()}
    wh_dests={}
    for wn in routes['wh_clean'].unique():
        wr=routes[routes['wh_clean']==wn]
        wda=wr.groupby(['to_city','geo_key']).agg(total_ton=('total_ton','sum'),n_routes=('total_ton','count')).reset_index()
        wdd=[]
        for _,r in wda.iterrows():
            co=cache.get(r.get('geo_key','')) or cache.get(r['to_city']+'|')
            if co: wdd.append({'name':r['to_city'],'lat':co[0],'lon':co[1],'ton':round(r['total_ton'],1),
                'routes':int(r['n_routes']),'orders':int(zd_wh.get((wn,r['to_city']),0))})
        wdd.sort(key=lambda x:x['ton'],reverse=True)
        wh_dests[wn]=wdd[:300]
    # Origin markers
    routes['from_zip2']=routes['from_zip'].apply(lambda x: str(x).replace('-','').replace(' ','')[:2] if pd.notna(x) else '')
    routes['from_geo_key']=routes['from_city']+'|'+routes['from_zip2']
    oa=routes.groupby(['from_city','from_geo_key']).agg(total_ton=('total_ton','sum'),n_routes=('total_ton','count')).reset_index()
    od=[]
    for _,r in oa.iterrows():
        co=cache.get(r['from_geo_key'])
        if co: od.append({'name':r['from_city'],'lat':co[0],'lon':co[1],'ton':round(r['total_ton'],1),
            'routes':int(r['n_routes']),'orders':int(zo.get(r['from_city'],0))})
    # Flows grouped by (warehouse, to_city) — use WAREHOUSE coords as origin
    # Build warehouse coord lookup from data
    # Build warehouse coord lookup from ZIP CODE (more precise than city name)
    WH_COORDS={}
    for wn in routes['wh_clean'].unique():
        wr=routes[routes['wh_clean']==wn].iloc[0]
        # Try zipcode first (more precise), then city name
        zc=get_coords_from_zip(str(wr['from_zip']))
        if zc: WH_COORDS[wn]=zc
        else:
            c=cache.get(wr['from_city'])
            if c: WH_COORDS[wn]=c
    fa=routes.groupby(['wh_clean','to_city','to_zip2']).agg(
        freq=('total_ton','count'),total_ton=('total_ton','sum'),
        from_city=('from_city','first')).reset_index()
    fd=[]
    for _,r in fa.iterrows():
        # Use warehouse coords (not city) — avoids same-city filtering issue
        wc=WH_COORDS.get(r['wh_clean'])
        if wc: fc_coord=wc
        else: fc_coord=cache.get(r['from_city'])
        to_z2=str(r.get('to_zip2','')).replace('-','').replace(' ','')[:2] if 'to_zip2' in fa.columns else ''
        tc=cache.get(r['to_city']+'|'+to_z2) or cache.get(r['to_city']+'|')
        if fc_coord and tc:
            # Allow same-city flows if distance > 2km (different zip areas)
            dist_km=((fc_coord[0]-tc[0])**2+(fc_coord[1]-tc[1])**2)**0.5 * 111
            if dist_km < 2: continue  # truly same point, skip
            fd.append({'from_lat':fc_coord[0],'from_lon':fc_coord[1],'to_lat':tc[0],'to_lon':tc[1],
                'freq':int(r['freq']),'ton':round(r['total_ton'],1),
                'from_name':r['from_city'],'to_name':r['to_city'],
                'orders':int(zf.get((r['wh_clean'],r['to_city']),0)),
                'src':r['wh_clean']})
    fd.sort(key=lambda x:x['freq'],reverse=True)
    # No limit — keep ALL flows. OSRM routes top ones, rest = dashed straight lines.
    OSRM_TOP=2500  # route this many via OSRM (rest = dashed)
    # ── OSRM routing with persistent disk cache ──
    import urllib.request, time as _time
    CACHE_FILE=os.path.join(SCRIPT_DIR,'osrm_cache.json')
    try:
        with open(CACHE_FILE,'r') as cf: disk_cache=json.load(cf)
    except: disk_cache={}
    cache_size_mb=os.path.getsize(CACHE_FILE)/1024/1024 if os.path.exists(CACHE_FILE) else 0
    print(f"  [{label}] Cache: {len(disk_cache)} routes ({cache_size_mb:.1f} MB)")
    # OSRM budget per run: max new routes to fetch (prevents long hangs)
    OSRM_BUDGET=2500
    routed=0; from_cache=0; skipped=0; osrm_available=True; consecutive_fails=0
    new_routes=0
    print(f"  [{label}] Routing {len(fd)} flows (budget: {OSRM_BUDGET} nowych, ~1 req/s)...")
    for idx_f, f in enumerate(fd):
        key=f"{f['from_lon']:.3f},{f['from_lat']:.3f};{f['to_lon']:.3f},{f['to_lat']:.3f}"
        if key in disk_cache:
            f['path']=disk_cache[key]; f['routed']=1; routed+=1; from_cache+=1; continue
        if new_routes >= OSRM_BUDGET or idx_f >= OSRM_TOP or not osrm_available:
            f['path']=[[f['from_lat'],f['from_lon']],[f['to_lat'],f['to_lon']]]; f['routed']=0; skipped+=1; continue
        url=f"https://router.project-osrm.org/route/v1/driving/{f['from_lon']},{f['from_lat']};{f['to_lon']},{f['to_lat']}?overview=full&geometries=geojson"
        got_route=False
        for attempt in range(2):
            try:
                req=urllib.request.Request(url,headers={'User-Agent':'Mozilla/5.0'})
                with urllib.request.urlopen(req,timeout=8) as resp:
                    rdata=json.loads(resp.read())
                    if rdata.get('code')=='Ok' and rdata.get('routes'):
                        coords=[[round(c[1],3),round(c[0],3)] for c in rdata['routes'][0]['geometry']['coordinates']]
                        if len(coords)>25:
                            step=max(1,len(coords)//25)
                            coords=coords[::step]+[coords[-1]]
                        f['path']=coords; f['routed']=1
                        disk_cache[key]=coords
                        routed+=1; new_routes+=1; got_route=True; consecutive_fails=0; break
                    else:
                        _time.sleep(2)
            except Exception:
                if new_routes==0 and attempt>=1:
                    print(f"    ⚠ OSRM niedostępny — tryb offline")
                    osrm_available=False; break
                _time.sleep(3*(attempt+1))
        if not got_route:
            f['path']=[[f['from_lat'],f['from_lon']],[f['to_lat'],f['to_lon']]]; f['routed']=0
            skipped+=1; consecutive_fails+=1
            if consecutive_fails>=5:
                print(f"    ⚠ Rate limit po {new_routes} trasach — zatrzymuję OSRM, reszta = linie proste")
                osrm_available=False
        if (idx_f+1)%25==0 or idx_f==len(fd)-1:
            print(f"    {idx_f+1}/{len(fd)}: {from_cache} cache + {new_routes} nowych + {skipped} proste")
        if new_routes>0 and new_routes%25==0:
            with open(CACHE_FILE,'w') as cf: json.dump(disk_cache,cf)
        # 1 second between requests — slow but reliable
        if osrm_available and got_route: _time.sleep(1.0)
    # Final cache save
    if new_routes>0:
        with open(CACHE_FILE,'w') as cf: json.dump(disk_cache,cf)
        new_size=os.path.getsize(CACHE_FILE)/1024/1024
        print(f"  [{label}] Cache: {len(disk_cache)} routes ({new_size:.1f} MB)")
    road_count=sum(1 for f in fd if f.get('routed')==1)
    dash_count=len(fd)-road_count
    print(f"  [{label}] Wynik: {road_count} po drogach, {dash_count} linie proste"
          f"{' — uruchom ponownie aby dokończyć' if dash_count>0 and osrm_available else ''}")
    # Colors per warehouse — bright/neon for dark map
    PALETTE=['#ff4444','#44ff44','#4488ff','#ff8800','#ff44ff',
             '#00ddff','#ffff00','#ff6688','#44ffcc','#ff9944',
             '#88aaff','#ccff44','#ff44aa','#44ffaa','#ffcc44',
             '#aa88ff','#ff6644','#66ffff','#ddff66','#ff88cc']
    wh_tons=routes.groupby('wh_clean')['total_ton'].sum().sort_values(ascending=False)
    oc={}
    for i,wn in enumerate(wh_tons.index):
        oc[wn]=PALETTE[i] if i<len(PALETTE) else '#a0a0a0'
    for f in fd: f['color']=oc.get(f['src'],'#a0a0a0')
    # Warehouse ranking — ALL (not top 15)
    wh=df.groupby('wh_clean').agg(ton=('waga_ton','sum'),
        zlecenia=('NrZlecenia','nunique') if 'NrZlecenia' in df.columns else ('waga_ton','count'),
        trasy=('NrSekcji = nr trasy','nunique')).reset_index().sort_values('ton',ascending=False)
    wl=[]; tt=wh.iloc[0]['ton'] if len(wh)>0 else 1
    for i,(_,r) in enumerate(wh.iterrows()):
        nm=str(r['wh_clean'])
        if len(nm)>28: nm=nm[:26]+"…"
        wl.append({'name':nm,'ton':round(r['ton']),'zlecenia':int(r['zlecenia']),
            'trasy':int(r['trasy']),'pct':round(r['ton']/tt*100)})
    # Top routes by frequency
    tr=fa.sort_values('freq',ascending=False).head(15)
    trl=[]; trf=int(tr.iloc[0]['freq']) if len(tr)>0 else 1
    for i,(_,r) in enumerate(tr.iterrows()):
        fn=str(r['wh_clean']); tn=str(r['to_city'])
        if len(fn)>15: fn=fn[:13]+"…"
        if len(tn)>15: tn=tn[:13]+"…"
        trl.append({'from':fn,'to':tn,'freq':int(r['freq']),'ton':round(r['total_ton'],1),
            'orders':int(zf.get((r['wh_clean'],r['to_city']),0)),'pct':round(int(r['freq'])/trf*100)})
    # origin_colors = ALL warehouses with REAL stats for filter panel
    wh_stats_full=df.groupby('wh_clean').agg(
        ton=('waga_ton','sum'),
        zlecenia=('NrZlecenia','nunique') if 'NrZlecenia' in df.columns else ('waga_ton','count'),
        trasy=('NrSekcji = nr trasy','nunique')
    ).reset_index()
    wh_stats_dict={r['wh_clean']:r for _,r in wh_stats_full.iterrows()}
    ocl=[]
    for wn in wh_tons.index:
        st=wh_stats_dict.get(wn,{})
        ocl.append({'city':wn,'color':oc[wn],'ton':round(float(wh_tons[wn])),
            'orders':int(st.get('zlecenia',0)),'routes':int(st.get('trasy',0))})
    print(f"  [{label}] {len(dd)} dests, {len(od)} origins, {len(fd)} flows, {len(oc)} src colors")
    # BH stats for filter
    bh_stats=[]
    if 'bh_mapped' in df.columns:
        bhs=df.groupby('bh_mapped').agg(ton=('waga_ton','sum'),
            trasy=('NrSekcji = nr trasy','nunique'),
            zlecenia=('NrZlecenia','nunique') if 'NrZlecenia' in df.columns else ('waga_ton','count')
        ).reset_index().sort_values('ton',ascending=False)
        for _,r in bhs.iterrows():
            if r['ton']>0:
                bh_stats.append({'name':r['bh_mapped'],'ton':round(r['ton']),'trasy':int(r['trasy']),'zlecenia':int(r['zlecenia'])})
    # Add bh field to flows
    bh_per_route=routes.set_index('NrSekcji = nr trasy')['bh'].to_dict() if 'bh' in routes.columns else {}
    # Map bh to flows via wh_clean+to_city → find dominant bh
    bh_flow=df.groupby(['wh_clean','to_norm','bh_mapped'])['waga_ton'].sum().reset_index()
    bh_flow_best=bh_flow.sort_values('waga_ton',ascending=False).drop_duplicates(subset=['wh_clean','to_norm'])
    bh_flow_dict={(r['wh_clean'],r['to_norm']):r['bh_mapped'] for _,r in bh_flow_best.iterrows()}
    # State (województwo) per flow
    if 'to_stateName' in df.columns:
        st_flow=df.groupby(['wh_clean','to_norm','to_stateName'])['waga_ton'].sum().reset_index()
        st_flow_best=st_flow.sort_values('waga_ton',ascending=False).drop_duplicates(subset=['wh_clean','to_norm'])
        st_flow_dict={(r['wh_clean'],r['to_norm']):str(r['to_stateName']) for _,r in st_flow_best.iterrows()}
    else:
        st_flow_dict={}
    for f in fd:
        f['bh']=bh_flow_dict.get((f['src'],f['to_name']),'NIEPRZYPISANE')
        f['state']=st_flow_dict.get((f['src'],f['to_name']),'')

    # State (województwo) stats
    state_stats=[]
    if 'to_stateName' in df.columns:
        sts=df.groupby('to_stateName').agg(ton=('waga_ton','sum'),
            trasy=('NrSekcji = nr trasy','nunique')).reset_index().sort_values('ton',ascending=False)
        for _,r in sts.iterrows():
            if r['ton']>0 and pd.notna(r['to_stateName']):
                state_stats.append({'name':str(r['to_stateName']),'ton':round(r['ton']),'trasy':int(r['trasy'])})
    return {'label':label,'fname':fname,'total_km':round(total_km),'total_ton':round(total_ton),'bh_stats':bh_stats,'state_stats':state_stats,
        'total_orders':total_orders,'total_routes':len(routes),
        'dests':dd,'wh_dests':wh_dests,'origins':od,'flows':fd,'wh':wl,'top_routes':trl,'origin_colors':ocl}

# ── FIND & PROCESS FILES ────────────────────────────────────
if len(sys.argv)>1:
    xlsx_files=[f for f in sys.argv[1:] if f.lower().endswith('.xlsx') and not os.path.basename(f).startswith('~$')]
else:
    xlsx_files=sorted([f for f in glob.glob(os.path.join(SCRIPT_DIR,'*.xlsx')) if not os.path.basename(f).startswith('~$')])

if not xlsx_files:
    print("ERROR: No .xlsx files found. Usage:")
    print(f"  python {os.path.basename(__file__)} file1.xlsx file2.xlsx ..."); sys.exit(1)

print(f"Found {len(xlsx_files)} file(s):")
datasets=[]
for fp in xlsx_files:
    try: datasets.append(process_file(fp))
    except Exception as e: print(f"  ⚠ Skipping {fp}: {e}")

if not datasets: print("ERROR: No files processed."); sys.exit(1)

# ── GENERATE HTML ────────────────────────────────────────────
print(f"\nGenerating HTML ({len(datasets)} datasets)...")
adj=json.dumps(datasets,ensure_ascii=False)

html = """<!DOCTYPE html>
<html lang="pl"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Logistics Flow Map</title>
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Segoe UI',system-ui,sans-serif;background:#1a1a2e}
#map{width:100vw;height:100vh}
.info-panel{position:absolute;top:12px;left:240px;z-index:1000;background:rgba(22,28,45,.94);border-radius:12px;padding:14px 18px;box-shadow:0 4px 24px rgba(0,0,0,.4);max-width:360px;backdrop-filter:blur(10px);border:1px solid rgba(255,255,255,.08)}
.info-panel h2{font-size:14px;color:#e2e8f0;margin-bottom:4px}
.info-panel .stat{font-size:17px;font-weight:700;color:#60a5fa}
.info-panel p{font-size:11px;color:#94a3b8;line-height:1.4}
.dataset-tabs{position:absolute;top:12px;left:50%;transform:translateX(-50%);z-index:1001;display:flex;gap:4px;background:rgba(22,28,45,.94);border-radius:10px;padding:4px;box-shadow:0 4px 20px rgba(0,0,0,.4);backdrop-filter:blur(10px);border:1px solid rgba(255,255,255,.08)}
.tab-btn{padding:7px 16px;border:none;border-radius:8px;font-size:11px;font-weight:600;cursor:pointer;color:#94a3b8;background:transparent;transition:all .2s;white-space:nowrap}
.tab-btn:hover{background:rgba(255,255,255,.1)}
.tab-btn.active{background:#3b82f6;color:#fff}
.controls{position:absolute;top:60px;right:12px;z-index:1000;display:flex;flex-direction:column;gap:5px}
.ctrl-btn{background:rgba(22,28,45,.9);border:1px solid rgba(255,255,255,.1);border-radius:8px;padding:7px 13px;font-size:11px;cursor:pointer;box-shadow:0 2px 8px rgba(0,0,0,.3);color:#e2e8f0;transition:all .2s}
.ctrl-btn:hover{background:#3b82f6;color:#fff}
.ctrl-btn.active{background:#3b82f6;color:#fff}
.filter-panel{position:absolute;top:60px;left:12px;z-index:1002;background:rgba(22,28,45,.94);border-radius:12px;padding:10px 12px;box-shadow:0 4px 20px rgba(0,0,0,.4);width:220px;max-height:55vh;overflow-y:auto;backdrop-filter:blur(10px);border:1px solid rgba(255,255,255,.08)}
.filter-panel h3{font-size:11px;color:#e2e8f0;margin-bottom:6px;font-weight:700;cursor:pointer;user-select:none}
.filter-panel h3::after{content:" ▼";font-size:8px;opacity:.5}
.filter-panel.collapsed h3::after{content:" ►"}
.filter-panel.collapsed .panel-body{display:none}
.filter-btn{display:flex;align-items:center;gap:6px;width:100%;padding:5px 8px;border:none;border-radius:6px;font-size:10px;cursor:pointer;color:#94a3b8;background:transparent;transition:all .15s;text-align:left;margin-bottom:2px}
.filter-btn:hover{background:rgba(255,255,255,.08)}
.filter-btn.active{background:rgba(59,130,246,.3);color:#fff;font-weight:600}
.filter-dot{width:10px;height:10px;border-radius:50%;flex-shrink:0}
.filter-label{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.filter-count{font-size:9px;color:#64748b;flex-shrink:0}
.filter-btn.active .filter-count{color:#94a3b8}
.filter-panel::-webkit-scrollbar{width:4px}
.filter-panel::-webkit-scrollbar-thumb{background:#475569;border-radius:2px}
.legend{position:absolute;bottom:30px;right:12px;z-index:1000;background:rgba(22,28,45,.94);border-radius:10px;padding:11px 13px;box-shadow:0 2px 12px rgba(0,0,0,.3);max-width:190px;border:1px solid rgba(255,255,255,.08)}
.legend h4{font-size:10px;color:#e2e8f0;margin-bottom:5px;font-weight:700}
.legend-row{display:flex;align-items:center;gap:7px;margin-bottom:3px}
.legend-line{width:26px;height:3px;border-radius:2px}
.legend-circle{border-radius:50%}
.legend-label{font-size:9px;color:#94a3b8}
.wh-panel{position:absolute;bottom:30px;left:12px;z-index:1000;background:rgba(22,28,45,.94);border-radius:12px;padding:12px 14px;box-shadow:0 4px 20px rgba(0,0,0,.4);width:310px;max-height:42vh;overflow-y:auto;border:1px solid rgba(255,255,255,.08)}
.wh-panel h3{font-size:11px;color:#e2e8f0;margin-bottom:8px;font-weight:700;cursor:pointer;user-select:none}
.wh-panel h3::after{content:' ▼';font-size:8px;opacity:.5}
.wh-panel.collapsed h3::after{content:' ►'}
.wh-panel.collapsed .panel-body{display:none}
.rt-panel h3{font-size:11px;color:#e2e8f0;margin-bottom:8px;font-weight:700;cursor:pointer;user-select:none}
.rt-panel h3::after{content:' ▼';font-size:8px;opacity:.5}
.rt-panel.collapsed h3::after{content:' ►'}
.rt-panel.collapsed .panel-body{display:none}
.wh-row{display:flex;align-items:center;gap:7px;margin-bottom:5px}
.wh-rank{width:20px;height:20px;border-radius:50%;background:#ef4444;color:#fff;font-size:9px;font-weight:700;display:flex;align-items:center;justify-content:center;flex-shrink:0}
.wh-row:nth-child(n+5) .wh-rank{background:#f97316}
.wh-row:nth-child(n+8) .wh-rank{background:#eab308}
.wh-row:nth-child(n+12) .wh-rank{background:#64748b}
.wh-info{flex:1;min-width:0}
.wh-name{font-size:10px;font-weight:600;color:#e2e8f0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.wh-bar-bg{height:3px;background:rgba(255,255,255,.1);border-radius:2px;margin:2px 0}
.wh-bar{height:3px;background:linear-gradient(90deg,#ef4444,#f97316);border-radius:2px}
.wh-stats{font-size:8px;color:#64748b}
.routing-status{position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);z-index:2000;background:rgba(15,23,42,.95);color:#fff;padding:10px 22px;border-radius:10px;font-size:12px;display:none;border:1px solid rgba(59,130,246,.3)}
.wh-panel::-webkit-scrollbar{width:4px}
.wh-panel::-webkit-scrollbar-thumb{background:#475569;border-radius:2px}
.rt-panel{position:absolute;bottom:30px;right:210px;z-index:1000;background:rgba(22,28,45,.94);border-radius:12px;padding:12px 14px;box-shadow:0 4px 20px rgba(0,0,0,.4);width:330px;max-height:42vh;overflow-y:auto;border:1px solid rgba(255,255,255,.08)}
.rt-row{display:flex;align-items:center;gap:7px;margin-bottom:5px}
.rt-rank{width:20px;height:20px;border-radius:50%;background:#3b82f6;color:#fff;font-size:9px;font-weight:700;display:flex;align-items:center;justify-content:center;flex-shrink:0}
.rt-row:nth-child(n+5) .rt-rank{background:#6366f1}
.rt-row:nth-child(n+8) .rt-rank{background:#8b5cf6}
.rt-row:nth-child(n+12) .rt-rank{background:#64748b}
.rt-info{flex:1;min-width:0}
.rt-name{font-size:10px;font-weight:600;color:#e2e8f0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.rt-bar-bg{height:3px;background:rgba(255,255,255,.1);border-radius:2px;margin:2px 0}
.rt-bar{height:3px;background:linear-gradient(90deg,#3b82f6,#8b5cf6);border-radius:2px}
.rt-stats{font-size:8px;color:#64748b}
.rt-panel::-webkit-scrollbar{width:4px}
.rt-panel::-webkit-scrollbar-thumb{background:#475569;border-radius:2px}
</style></head><body>
<div id="map"></div>
<div class="dataset-tabs" id="datasetTabs"></div>
<div class="info-panel" id="infoPanel"></div>
<div class="controls">
  <button class="ctrl-btn active" onclick="toggleLayer('flows')">&#x1f4ca; Przepływy</button>
  <button class="ctrl-btn active" onclick="toggleLayer('dests')">&#x1f4cd; Odbiorcy</button>
  <button class="ctrl-btn active" onclick="toggleLayer('origins')">&#x1f3ed; Magazyny</button>
  <div style="margin-top:8px;padding-top:6px;border-top:1px solid rgba(255,255,255,.1)">
    <div style="font-size:9px;color:#94a3b8;margin-bottom:4px">Min. tonaż:</div>
    <button class="ctrl-btn active" data-mt="0" onclick="setMinTon(0)">Wszystko</button>
    <button class="ctrl-btn" data-mt="50" onclick="setMinTon(50)">&gt;50t</button>
    <button class="ctrl-btn" data-mt="100" onclick="setMinTon(100)">&gt;100t</button>
    <button class="ctrl-btn" data-mt="500" onclick="setMinTon(500)">&gt;500t</button>
    <button class="ctrl-btn" data-mt="1000" onclick="setMinTon(1000)">&gt;1000t</button>
  </div>
</div>
<div class="filter-panel" id="filterPanel">
  <h3 onclick="this.parentElement.classList.toggle('collapsed')">&#x1f50d; Filtruj wg magazynu</h3>
  <div class="panel-body">
  <button class="filter-btn" id="filterAll" style="color:#ef4444;font-size:9px">&#x2716; Wyczyść magazyny</button>
  <div id="filterList"></div>
  </div>
</div>
<div class="legend" id="legendPanel">
  <h4>Grubość = frekwencja trasy</h4>
  <div class="legend-row"><div class="legend-line" style="background:#888;height:1.5px"></div><span class="legend-label">1–5 tras</span></div>
  <div class="legend-row"><div class="legend-line" style="background:#888;height:3px"></div><span class="legend-label">21–50 tras</span></div>
  <div class="legend-row"><div class="legend-line" style="background:#888;height:5.5px"></div><span class="legend-label">&gt;100 tras</span></div>
  <h4 style="margin-top:7px">Kolor = magazyn źródłowy</h4>
  <div id="colorLegend"></div>
  <h4 style="margin-top:7px">Punkty</h4>
  <div class="legend-row"><div class="legend-circle" style="width:9px;height:9px;background:rgba(96,165,250,.55);border:1.5px solid rgba(96,165,250,.7)"></div><span class="legend-label">Odbiorca</span></div>
  <div class="legend-row"><div class="legend-circle" style="width:9px;height:9px;background:rgba(255,68,68,.8);border:2px solid #fff"></div><span class="legend-label">Magazyn</span></div>
</div>
<div class="filter-panel" id="statePanel" style="top:auto;bottom:30px;left:560px;width:220px;max-height:38vh">
  <h3 onclick="this.parentElement.classList.toggle('collapsed')">&#x1f5fa; Województwa</h3>
  <div class="panel-body">
  <div style="font-size:8px;color:#64748b;margin-bottom:4px">Zaznacz jedno lub wiele</div>
  <button class="filter-btn" id="stClear" style="color:#ef4444;font-size:9px">&#x2716; Wyczyść</button>
  <div id="stateList"></div>
  </div>
</div>
<div class="wh-panel" id="whPanel"></div>
<div class="rt-panel" id="rtPanel"></div>
<div class="filter-panel" id="bhPanel" style="top:auto;bottom:30px;left:240px;width:240px;max-height:38vh">
  <h3 onclick="this.parentElement.classList.toggle('collapsed')">&#x1f4bc; Biura Handlowe</h3>
  <div class="panel-body">
  <button class="filter-btn" id="bhAll" style="color:#ef4444;font-size:9px">&#x2716; Wyczyść BH</button>
  <div id="bhList"></div>
  </div>
</div>
<script>
const ALL="""  + adj + """;
let currentIdx=0, minTon=0;
let activeWH=new Set();
let activeBHs=new Set();
const map=L.map('map',{center:[51.9,19.15],zoom:6,zoomControl:false,preferCanvas:true});
L.control.zoom({position:'bottomleft'}).addTo(map);
L.tileLayer('https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png',{attribution:'&copy; OpenStreetMap &copy; CARTO',subdomains:'abcd',maxZoom:19}).addTo(map);
function fw(f){if(f<=5)return 1.5;if(f<=20)return 2;if(f<=50)return 3;if(f<=100)return 4;return 5.5}
function dr(t){if(t<50)return 4;if(t<200)return 6;if(t<500)return 8;if(t<1000)return 11;if(t<5000)return 16;if(t<10000)return 22;return 28}
const flowLayer=L.layerGroup().addTo(map),destLayer=L.layerGroup().addTo(map),originLayer=L.layerGroup().addTo(map);
const layers={flows:flowLayer,dests:destLayer,origins:originLayer},lv={flows:true,dests:true,origins:true};
function toggleLayer(n){const b=event.target;if(map.hasLayer(layers[n])){map.removeLayer(layers[n]);b.classList.remove('active');lv[n]=false}else{map.addLayer(layers[n]);b.classList.add('active');lv[n]=true}}
function setMinTon(val){
  minTon=val;
  document.querySelectorAll('.controls [data-mt]').forEach(b=>b.classList.remove('active'));
  event.target.classList.add('active');
  drawMap(ALL[currentIdx]);
}
const tabsEl=document.getElementById('datasetTabs');
ALL.forEach((ds,i)=>{const b=document.createElement('button');b.className='tab-btn'+(i===0?' active':'');b.textContent=ds.label;b.onclick=()=>{activeWH.clear();activeBHs.clear();activeStates.clear();switchDataset(i)};tabsEl.appendChild(b)});

// BH filter
let activeStates=new Set();
function toggleState(st,btn){
  if(activeStates.has(st)){activeStates.delete(st);btn.classList.remove('active')}
  else{activeStates.add(st);btn.classList.add('active')}
  // Update header with count
  const cnt=activeStates.size;
  const hdr=document.querySelector('#statePanel h3');
  hdr.textContent=cnt>0?'\ud83d\uddfa Województwa ('+cnt+')':'\ud83d\uddfa Województwa';
  drawMap(ALL[currentIdx]);
}
function clearStates(){
  activeStates.clear();
  document.querySelectorAll('#stateList .filter-btn').forEach(b=>b.classList.remove('active'));
  document.querySelector('#statePanel h3').textContent='\ud83d\uddfa Województwa';
  drawMap(ALL[currentIdx]);
}
function toggleBH(bh,btn){
  if(activeBHs.has(bh)){activeBHs.delete(bh);btn.classList.remove('active')}
  else{activeBHs.add(bh);btn.classList.add('active')}
  const cnt=activeBHs.size;
  document.getElementById('bhAll').classList.toggle('active',cnt===0);
  document.querySelector('#bhPanel h3').textContent=cnt>0?'\ud83d\udcbc BH ('+cnt+')':'\ud83d\udcbc Biura Handlowe';
  drawMap(ALL[currentIdx]);
}
function clearBHs(){
  activeBHs.clear();
  document.querySelectorAll('#bhList .filter-btn').forEach(b=>b.classList.remove('active'));
  document.getElementById('bhAll').classList.add('active');
  document.querySelector('#bhPanel h3').textContent='\ud83d\udcbc Biura Handlowe';
  drawMap(ALL[currentIdx]);
}
function drawMap(ds){
 try{
  const ocMap={};ds.origin_colors.forEach(oc=>{ocMap[oc.city]=oc.color});

  // Start with all data
  let flows=ds.flows;
  let dests=ds.dests;
  let origins=ds.origins;

  // Filter by warehouses (multi-select)
  if(activeWH.size>0){
    flows=flows.filter(f=>activeWH.has(f.src));
    origins=origins.filter(o=>activeWH.has(o.name));
  }
  // Filter by BH (multi-select)
  if(activeBHs.size>0){
    flows=flows.filter(f=>activeBHs.has(f.bh));
  }
  // Filter by state/województwo (multi-select)
  if(activeStates.size>0){
    flows=flows.filter(f=>activeStates.has(f.state));
  }
  // Filter dests to match remaining flows — with CORRECT tonnage from filtered flows
  if(activeWH.size>0||activeBHs.size>0||activeStates.size>0){
    // Build dest stats from filtered flows (not global dests)
    const destMap={};
    flows.forEach(f=>{
      if(!destMap[f.to_name])destMap[f.to_name]={ton:0,freq:0,orders:0};
      destMap[f.to_name].ton+=f.ton;
      destMap[f.to_name].freq+=f.freq;
      destMap[f.to_name].orders+=f.orders;
    });
    // Get coords from global dests
    const coordMap={};
    ds.dests.forEach(d=>{coordMap[d.name]={lat:d.lat,lon:d.lon}});
    dests=Object.keys(destMap).map(name=>{
      const c=coordMap[name]||{lat:0,lon:0};
      return {name:name,lat:c.lat,lon:c.lon,ton:Math.round(destMap[name].ton*10)/10,
              routes:destMap[name].freq,orders:destMap[name].orders};
    });
  }
  // Apply min tonnage filter
  if(minTon>0){
    flows=flows.filter(f=>f.ton>=minTon);
    dests=dests.filter(d=>d.ton>=minTon);
  }

  // Stats — compute from filtered flows, or use totals
  const anyFilter=activeWH.size>0||activeBHs.size>0||activeStates.size>0;
  let fTon,fFreq,fOrders,heading;
  if(anyFilter){
    fTon=flows.reduce((s,f)=>s+f.ton,0);
    fFreq=flows.reduce((s,f)=>s+f.freq,0);
    fOrders=flows.reduce((s,f)=>s+f.orders,0);
    const parts=[];
    if(activeWH.size>0)parts.push(activeWH.size+' mag.');
    if(activeBHs.size>0)parts.push(activeBHs.size+' BH');
    if(activeStates.size>0)parts.push(activeStates.size+' woj.');
    heading='Filtr: '+parts.join(' + ');
  }else{
    fTon=ds.total_ton;fFreq=ds.total_routes;fOrders=ds.total_orders;
    heading=ds.label;
  }

  document.getElementById('infoPanel').innerHTML=
    '<h2>&#x1f69b; '+heading+'</h2>'+
    (anyFilter?'<p style="font-size:10px;color:#60a5fa;margin-bottom:3px">'+heading+'</p>':'')+
    '<p><span class="stat">'+fOrders.toLocaleString('pl')+'</span> zleceń | '+
    '<span class="stat">'+fFreq.toLocaleString('pl')+'</span> tras | '+
    '<span class="stat">'+Math.round(fTon/1000)+'k</span> ton</p>'+
    '<p style="margin-top:3px;font-size:10px">'+ds.fname+(anyFilter?' | '+dests.length+' klientów':'')+'</p>';

  // Color legend — show active warehouse colors or all
  let cl='';
  if(activeWH.size>0){
    ds.origin_colors.forEach(oc=>{
      if(activeWH.has(oc.city))cl+='<div class="legend-row"><div class="legend-line" style="background:'+oc.color+';height:4px"></div><span class="legend-label">'+oc.city+' ('+Math.round(oc.ton/1000)+'k t)</span></div>';
    });
  }else{
    ds.origin_colors.forEach(oc=>{cl+='<div class="legend-row"><div class="legend-line" style="background:'+oc.color+';height:4px"></div><span class="legend-label">'+oc.city+' ('+Math.round(oc.ton/1000)+'k t)</span></div>'});
  }
  document.getElementById('colorLegend').innerHTML=cl;

  // Clear & redraw — flows FIRST (bottom), then markers ON TOP
  flowLayer.clearLayers();destLayer.clearLayers();originLayer.clearLayers();
  flows.forEach(f=>{
    const path=f.path||[[f.from_lat,f.from_lon],[f.to_lat,f.to_lon]];
    const isRoad=f.routed&&path.length>2;
    const op=isRoad?Math.min(0.55+f.freq/150,0.9):Math.min(0.25+f.freq/400,0.5);
    const da=isRoad?null:'6 4';
    const w=isRoad?fw(f.freq):Math.max(1,fw(f.freq)*0.7);
    const line=L.polyline(path,{color:f.color,weight:w,opacity:op,dashArray:da,smoothFactor:1,pane:'overlayPane'});
    line.bindPopup('<b>'+f.from_name+' → '+f.to_name+'</b><br>Zlecenia: <b>'+f.orders+'</b><br>Trasy: <b>'+f.freq+'</b><br>Wolumen: <b>'+f.ton.toLocaleString('pl')+' t</b>');
    flowLayer.addLayer(line);
  });
  dests.forEach(d=>{const m=L.circleMarker([d.lat,d.lon],{radius:dr(d.ton),fillColor:'#60a5fa',fillOpacity:.6,color:'rgba(96,165,250,.8)',weight:1.5,pane:'markerPane'});m.bindPopup('<b>'+d.name+'</b><br>Tonaż: <b>'+d.ton.toLocaleString('pl')+' t</b><br>Zlecenia: <b>'+d.orders+'</b><br>Trasy: <b>'+d.routes+'</b>');destLayer.addLayer(m)});
  origins.forEach(o=>{const c=ocMap[o.name]||'#ff4444';const m=L.circleMarker([o.lat,o.lon],{radius:Math.max(8,Math.min(20,5+Math.sqrt(o.ton/500))),fillColor:c,fillOpacity:.9,color:'#fff',weight:2,pane:'markerPane'});m.bindPopup('<b>'+o.name+'</b><br>Tonaż: <b>'+o.ton.toLocaleString('pl')+' t</b><br>Zlecenia: <b>'+o.orders+'</b><br>Trasy: <b>'+o.routes+'</b>');originLayer.addLayer(m)});
  Object.keys(layers).forEach(k=>{if(lv[k]&&!map.hasLayer(layers[k]))map.addLayer(layers[k]);if(!lv[k]&&map.hasLayer(layers[k]))map.removeLayer(layers[k])});
 }catch(e){console.error('drawMap error:',e);document.title='ERROR: '+e.message}
}

function switchDataset(idx){
 try{
  currentIdx=idx;
  tabsEl.querySelectorAll('.tab-btn').forEach((b,i)=>b.classList.toggle('active',i===idx));
  const ds=ALL[idx];
  // Warehouse side panels
  let wh='<h3 onclick="this.parentElement.classList.toggle(\\x27collapsed\\x27)">&#x1f3ed; TOP 15 magazynów</h3><div class="panel-body">';
  ds.wh.forEach((w,i)=>{wh+='<div class="wh-row"><div class="wh-rank">'+(i+1)+'</div><div class="wh-info"><div class="wh-name">'+w.name+'</div><div class="wh-bar-bg"><div class="wh-bar" style="width:'+w.pct+'%"></div></div><div class="wh-stats">'+w.ton.toLocaleString('pl')+' t | '+w.zlecenia.toLocaleString('pl')+' zl. | '+w.trasy.toLocaleString('pl')+' tras</div></div></div>'});
  document.getElementById('whPanel').innerHTML=wh+'</div>';
  let rt='<h3 onclick="this.parentElement.classList.toggle(\\x27collapsed\\x27)">&#x1f6e3; TOP 15 tras</h3><div class="panel-body">';
  ds.top_routes.forEach((r,i)=>{rt+='<div class="rt-row"><div class="rt-rank">'+(i+1)+'</div><div class="rt-info"><div class="rt-name">'+r.from+' → '+r.to+'</div><div class="rt-bar-bg"><div class="rt-bar" style="width:'+r.pct+'%"></div></div><div class="rt-stats">'+r.freq+' tras | '+r.orders+' zl. | '+r.ton.toLocaleString('pl')+' t</div></div></div>'});
  document.getElementById('rtPanel').innerHTML=rt+'</div>';
  // Build filter list
  const ocMap={};ds.origin_colors.forEach(oc=>{ocMap[oc.city]=oc.color});
  document.getElementById('filterList').innerHTML='';
  ds.origin_colors.forEach((oc,i)=>{
    const btn=document.createElement('button');
    btn.className='filter-btn';
    btn.dataset.city=oc.city;
    btn.innerHTML='<div class="filter-dot" style="background:'+oc.color+'"></div>'+
      '<span class="filter-label">'+oc.city+'</span>'+
      '<span class="filter-count">'+oc.routes+' tras</span>';
    btn.addEventListener('click',function(e){e.stopPropagation();toggleWarehouse(this.dataset.city,this)});
    document.getElementById('filterList').appendChild(btn);
  });
  const allBtn=document.getElementById('filterAll');
  allBtn.classList.add('active');
  allBtn.onclick=function(e){e.stopPropagation();clearWarehouses()};
  // State filter buttons
  if(ds.state_stats&&ds.state_stats.length>0){
    document.getElementById('statePanel').style.display='block';
    document.getElementById('stateList').innerHTML='';
    const sc=document.getElementById('stClear');
    const scNew=sc.cloneNode(true);sc.parentNode.replaceChild(scNew,sc);
    scNew.addEventListener('click',function(e){e.stopPropagation();clearStates()});
    ds.state_stats.forEach(st=>{
      const btn=document.createElement('button');btn.className='filter-btn';
      if(activeStates.has(st.name))btn.classList.add('active');
      btn.innerHTML='<span class="filter-label">'+st.name+'</span><span class="filter-count">'+Math.round(st.ton).toLocaleString('pl')+' t</span>';
      btn.addEventListener('click',function(e){e.stopPropagation();toggleState(st.name,this)});
      document.getElementById('stateList').appendChild(btn);
    });
  }else{document.getElementById('statePanel').style.display='none'}
  // BH filter buttons
  if(ds.bh_stats&&ds.bh_stats.length>0){
    document.getElementById('bhPanel').style.display='block';
    document.getElementById('bhList').innerHTML='';
    const ba=document.getElementById('bhAll');
    const baNew=ba.cloneNode(true);ba.parentNode.replaceChild(baNew,ba);
    baNew.addEventListener('click',function(e){e.stopPropagation();clearBHs()});
    ds.bh_stats.forEach(bh=>{
      const btn=document.createElement('button');btn.className='filter-btn';
      btn.dataset.bh=bh.name;
      btn.innerHTML='<span class="filter-label">'+bh.name+'</span><span class="filter-count">'+Math.round(bh.ton).toLocaleString('pl')+' t</span>';
      btn.addEventListener('click',function(e){e.stopPropagation();toggleBH(this.dataset.bh,this)});
      document.getElementById('bhList').appendChild(btn);
    });
  }else{document.getElementById('bhPanel').style.display='none'}
  drawMap(ALL[currentIdx]);
 }catch(e){console.error('switchDataset error:',e);document.title='SD ERROR: '+e.message}
}

function toggleWarehouse(wh,btn){
  if(activeWH.has(wh)){activeWH.delete(wh);btn.classList.remove('active')}
  else{activeWH.add(wh);btn.classList.add('active')}
  document.getElementById('filterAll').classList.toggle('active',activeWH.size===0);
  document.querySelector('#filterPanel h3').textContent=activeWH.size>0?'\ud83d\udd0d Magazyny ('+activeWH.size+')':'\ud83d\udd0d Magazyny';
  drawMap(ALL[currentIdx]);
}
function clearWarehouses(){
  activeWH.clear();
  document.querySelectorAll('#filterList .filter-btn').forEach(b=>b.classList.remove('active'));
  document.getElementById('filterAll').classList.add('active');
  document.querySelector('#filterPanel h3').textContent='\ud83d\udd0d Magazyny';
  drawMap(ALL[currentIdx]);
}
switchDataset(0);
</script></body></html>"""

output_path = os.path.join(SCRIPT_DIR, 'logistics_network_map.html')
# Clean surrogates from data (can come from Excel strings)
html = html.encode('utf-8', errors='replace').decode('utf-8')
with open(output_path, 'w', encoding='utf-8') as f:
    f.write(html)
print(f"\n✅ Saved: {output_path}")
print(f"   {len(datasets)} dataset(s). Open in browser — tabs at top to switch.")
