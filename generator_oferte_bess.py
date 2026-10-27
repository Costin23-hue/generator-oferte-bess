import json
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import tkinter as tk
from datetime import datetime
from tkinter import ttk, messagebox, filedialog

import pandas as pd
import xlwings as xw
from pptx import Presentation

# ======================================================================
# CONFIGURARE  (aici schimbi regulile, nu în restul codului)
# ======================================================================
# Folderul programului ('Ofertare BESS', comun în OneDrive): lângă .exe sau lângă scriptul .py.
# Aici stau PCS, specificațiile și template-urile; ofertele se salvează în subfolderul FOLDER_OFERTE.
# Pe Mac, aplicația e 'Generator Oferte BESS.app/Contents/MacOS/...': folderul e cel în care stă .app-ul.
# Dacă fișierele nu sunt lângă program (ex. .app-ul e în Applications), programul cere o dată folderul
# 'Ofertare BESS' și îl ține minte (vezi alege_folder_lucru).
if getattr(sys, "frozen", False):
    FOLDER_PROGRAM = os.path.dirname(sys.executable)
    if sys.platform == "darwin" and FOLDER_PROGRAM.endswith(os.path.join("Contents", "MacOS")):
        FOLDER_PROGRAM = os.path.dirname(os.path.dirname(os.path.dirname(FOLDER_PROGRAM)))
else:
    FOLDER_PROGRAM = os.path.dirname(os.path.abspath(__file__))
FOLDER_OFERTE = "Oferte generate"

# Pricing.xlsx se citește (doar citire) din folderul Purchasing sincronizat la fiecare utilizator.
# Numele folderului diferă de la un PC la altul (ex. 'Levon Çağla - Purchasing Prices Update_BESS'),
# așa că se caută în folderele din C:\Users\<utilizator> unul care se termină cu acest text.
NUME_FOLDER_PRICING = "Purchasing Prices Update_BESS"
FISIER_PRICING = "Pricing.xlsx"
# Unde ține minte programul (per utilizator) calea aleasă manual pentru Pricing.xlsx
if sys.platform == "win32":
    _FOLDER_SETARI = os.environ.get("APPDATA", os.path.expanduser("~"))
elif sys.platform == "darwin":
    _FOLDER_SETARI = os.path.join(os.path.expanduser("~"), "Library", "Application Support")
else:
    _FOLDER_SETARI = os.path.join(os.path.expanduser("~"), ".config")
FISIER_SETARI = os.path.join(_FOLDER_SETARI, "Ofertare BESS", "setari.json")

FISIER_SPECS = "Specificatii_baterii_BESS.xlsx"   # coloane: Producător / Produs | Model | Configuration | Weight
FISIER_PCS = "PCS - MV skid.xlsx"
FORMAT_DATA = "%d/%m/%y"                    # cum apare data în prezentare (ex. 15/10/26)
# Nr. proiect / contract: câmpul din fereastră pornește cu ST + anul și luna curentă + XXX (ex. ST202610XXX);
# îl completezi tu (ex. ST202606109). Apare în prezentare la [NR_PROIECT] și în numele fișierului:
# ex. ST202606109_BESS_Indicative_Proposal_Taliva_Proiect X_20MWh_10MW.pptx (+ .pdf)
NR_PROIECT_IMPLICIT = "ST{an_luna}XXX"
NUME_PREZENTARE = "{nr_proiect}_BESS_Indicative_Proposal_Taliva_{proiect}_{mwh}MWh_{mw}MW"
# Versiunea (_v1, _v2, ...) crește automat la fiecare ofertă generată pentru același nr. proiect + proiect,
# indiferent de MWh/MW; apare în numele fișierelor și în prezentare la [VERSIUNE] (ex. "Version: [VERSIUNE]").
FOAIE_BATERIE = "Battery 2027 delivery"   # foaia cu prețurile actuale (foaia "Battery" e ascunsă și veche)
FISIER_EXCEL_TEMPLATE = "Project Proposal TAMPLATE.xlsx"
FISIER_PPT_TEMPLATE = "TEMPLATE BESS.pptx"

BATERIE_IMPLICITA = "JKE-5015K-2H-LAA"
CAPACITATE_BATERIE_FALLBACK = 5.015      # MWh, folosită doar dacă nu se găsește în Excel

# Câte intrări DC ocupă o baterie pe skid (limitele Min/Max din PCS - MV skid.xlsx sunt în intrări DC)
DC_INPUTURI_PER_BATERIE = 2
DC_INPUTURI_PER_BRAND = {"Sinexcel": 1}   # excepții pe brand de skid (Sinexcel: 1 intrare DC / baterie)

# Baterii REPT: merg doar pe skidurile de mai jos. Modul de conectare pe brand de skid:
#   "rack" = rack-urile direct la PCS (TBEA): o intrare ia 1 rack (1:1) sau 2 rack-uri (1:2).
#            Ex.: 4 baterii x 12 rack-uri = 48 rack-uri -> 48 intrări la 1:1 sau 24 intrări la 1:2.
#   "dc"   = ca la celelalte baterii (DC_INPUTURI_PER_BATERIE intrări DC / baterie).
BATERIE_RACK = "REPT"
SKIDURI_PENTRU_RACK = {"TBEA": "rack", "Kehua": "dc"}
RACKURI_PER_BATERIE = 12
RACKURI_PER_INTRARE = (1, 2)

# Limitele de intrări DC pe skid (Min, Max) pentru brandurile ale căror skiduri NU sunt în 'PCS - MV skid.xlsx'.
# TBEA: rack-urile bateriei se conectează direct la PCS; un skid TBEA ia 12 sau 24 de intrări.
# Puterea skidului se ia atunci din coloana Power din Pricing. Dacă skidul apare în fișierul PCS, au prioritate datele de acolo.
INTRARI_FARA_PCS = {"TBEA": (12, 24)}

# Consum auxiliar / baterie, după capacitatea bateriei: (de la MWh, kVA), crescător.
# 5,015 MWh -> 50 kVA; 6,25 MWh -> 65 kVA. Tabelul se scrie și în Excelul de calcul (A24:B25).
KVA_AUX_PER_BATERIE = [(0, 50), (5.6, 65)]
# Trafo auxiliar: din foaia 'Auxiliary Transformer' se iau doar rândurile cu acest text în 'Detail';
# puterea (kVA) se citește din nume (ex. '1250 KVA Aux transformer'); prețul e total (cu transport).
TRAFO_FILTRU = "aux transformer"

ORE_MIN, ORE_MAX = 0.5, 8                # raport MWh/MW considerat normal
ADAOS_MIN, ADAOS_MAX = 1.0, 2.0
CURS_MIN, CURS_MAX = 0.5, 2.0
EMS_PRAG_MWH = 15          # de la această capacitate EMS se completează automat
EMS_AUTO = "3500"          # valoarea EMS (USD) pusă automat la capacitate >= prag
EMS_SUB_PRAG = "6500"      # valoarea pusă automat la capacitate sub prag (dacă nu ai modificat EMS manual)
NR_OPTIUNI_SKID = 5

# Punct de conexiune (POC): la 20 kV se folosește foaia "MV SKID" (România), ca până acum;
# la orice altă tensiune, foaia ale cărei nume conține "POC" ("MV SKID for different POC"),
# doar skidurile care au acea tensiune în nume (ex. "SES4H-5000-EX-MV 10KV SF6 switchgear").
POC_IMPLICIT_KV = 20
FOAIE_SKID = "MV SKID"
FOAIE_SKID_POC_CONTINE = "poc"

# Condiția de livrare: textul pus în prezentare la [INCOTERMS] și coloana de preț baterie folosită
INCOTERMS = {
    "DDP": "DDP Incoterms 2020",   # preț din coloana I
    "DAP": "DAP Incoterms 2020",   # preț din coloana H
}
INCOTERM_IMPLICIT = "DDP"
AUTO_SKID = "Auto (cel mai ieftin)"

# Indecși de coloană (0 = A) în foile din Pricing.xlsx
BAT_COL_DATA, BAT_COL_BRAND, BAT_COL_TYPE, BAT_COL_CAP, BAT_COL_PRET, BAT_COL_TRANSPORT = 0, 2, 3, 4, 8, 9
BAT_COL_PRET_DAP = 7                     # coloana H = 'Price / MWh ... with taxes (DAP)'
BAT_COL_GARANTIE = 20                    # coloana U = ani garanție baterie
SKID_COL_BRAND, SKID_COL_TYPE, SKID_COL_POWER = 2, 3, 4
SKID_COL_PRET, SKID_COL_TRANSPORT, SKID_COL_COM = 5, 6, 7
SKID_COL_GARANTIE = 12                   # coloana M = ani garanție skid / PCS

# Celule din Excelul de calcul citite pentru prezentare (extensie garanție, $/MWh pe an)
CELULA_EXT_GAR_BAT = "K9"
CELULA_EXT_GAR_PCS = "K11"
# Branduri care NU oferă extensie de garanție: rândul lor dispare din tabelul 'Warranty Extension'
# (rândul cu [EXT_GAR_BAT] pentru baterie, cel cu [EXT_GAR_PCS] pentru skid); dacă dispar ambele, dispare tot tabelul.
FARA_EXTENSIE_GARANTIE = ("REPT", "TBEA")


# ======================================================================
# LOGICA (fără Tkinter)
# ======================================================================
def numar(valoare, implicit=0.0):
    """Convertește în float; NaN/None/text invalid -> valoarea implicită."""
    try:
        if pd.isna(valoare):
            return implicit
        return float(valoare)
    except (TypeError, ValueError):
        return implicit


def citeste_setari():
    try:
        with open(FISIER_SETARI, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def salveaza_setari(setari):
    os.makedirs(os.path.dirname(FISIER_SETARI), exist_ok=True)
    with open(FISIER_SETARI, "w", encoding="utf-8") as f:
        json.dump(setari, f, ensure_ascii=False, indent=2)


def cauta_pricing_in_onedrive():
    """Caută folderul sincronizat al cărui nume se termină cu NUME_FOLDER_PRICING și care conține Pricing.xlsx,
    pe 2 niveluri sub: C:\\Users\\<utilizator> (Windows, ex. 'TALİVA ...\\Levon Çağla - Purchasing Prices
    Update_BESS') și ~/Library/CloudStorage (Mac, ex. 'OneDrive-SharedLibraries-.../Levon Çağla - ...')."""
    acasa = os.path.expanduser("~")
    radacini = [acasa, os.path.join(acasa, "Library", "CloudStorage")]
    def subfoldere(cale):
        try:
            return [os.path.join(cale, d) for d in os.listdir(cale) if os.path.isdir(os.path.join(cale, d))]
        except OSError:
            return []
    for nivel1 in (n for radacina in radacini for n in subfoldere(radacina)):
        for folder in [nivel1] + subfoldere(nivel1):
            cale = os.path.join(folder, FISIER_PRICING)
            if folder.lower().endswith(NUME_FOLDER_PRICING.lower()) and os.path.exists(cale):
                return cale
    return None


class PricingNegasit(FileNotFoundError):
    pass


def fisiere_lipsa(folder="."):
    """Fișierele care trebuie să stea în folderul 'Ofertare BESS' și lipsesc."""
    return [f for f in (FISIER_PCS, FISIER_SPECS, FISIER_EXCEL_TEMPLATE, FISIER_PPT_TEMPLATE)
            if not os.path.exists(os.path.join(folder, f))]


def alege_folder_lucru():
    """Folderul 'Ofertare BESS': lângă program, altfel cel ales anterior, altfel îl alege utilizatorul (o dată)."""
    for folder in (FOLDER_PROGRAM, citeste_setari().get("folder_ofertare")):
        if folder and os.path.isdir(folder) and not fisiere_lipsa(folder):
            return folder
    messagebox.showinfo("Folder Ofertare BESS",
                        "Nu găsesc template-urile lângă program.\n\n"
                        "Alege folderul 'Ofertare BESS' (cel sincronizat din OneDrive).")
    folder = filedialog.askdirectory(title="Alege folderul Ofertare BESS")
    if folder and not fisiere_lipsa(folder):
        salveaza_setari(dict(citeste_setari(), folder_ofertare=folder))
        return folder
    return FOLDER_PROGRAM


def cale_pricing():
    """Calea aleasă manual (ținută minte), altfel căutare automată în OneDrive, altfel lângă program."""
    for cale in (citeste_setari().get("pricing"), cauta_pricing_in_onedrive(), FISIER_PRICING):
        if cale and os.path.exists(cale):
            return cale
    raise PricingNegasit(
        f"Nu găsesc {FISIER_PRICING}. Verifică dacă OneDrive rulează și dacă folderul "
        f"'{NUME_FOLDER_PRICING}' e sincronizat pe acest PC.")


def cale_specs():
    """Specificatii_baterii_BESS.xlsx din folderul programului, altfel None."""
    return FISIER_SPECS if os.path.exists(FISIER_SPECS) else None


def citeste_copie(sursa, foi):
    """Citim o copie temporară, ca să nu ne blocăm dacă fișierul e deschis în Excel / se sincronizează.
    foi = None -> prima foaie; "*" -> toate foile {nume: DataFrame}; altfel listă de nume -> {nume: DataFrame}."""
    with tempfile.TemporaryDirectory() as tmp:
        copie = os.path.join(tmp, os.path.basename(sursa))
        shutil.copy2(sursa, copie)
        with pd.ExcelFile(copie) as xls:
            if foi is None:
                return pd.read_excel(xls, sheet_name=0)
            if foi == "*":
                return pd.read_excel(xls, sheet_name=None)
            return {foaie: pd.read_excel(xls, sheet_name=foaie) for foaie in foi}


def incarca_date():
    """Citește toate fișierele de pricing și returnează un dicționar de DataFrame-uri."""
    sursa = cale_pricing()
    foi = citeste_copie(sursa, "*")
    df_bat, df_skid, df_aux = foi[FOAIE_BATERIE], foi[FOAIE_SKID], foi["Auxiliary Transformer"]
    nume_poc = next((f for f in foi if FOAIE_SKID_POC_CONTINE in f.lower()), None)
    df_poc = fixeaza_antet(foi[nume_poc]) if nume_poc else None
    df_pcs = citeste_copie(FISIER_PCS, ["Sheet1"])["Sheet1"]
    sursa_specs = cale_specs()
    df_specs = citeste_copie(sursa_specs, None) if sursa_specs else None
    for df in (df_bat, df_skid, df_aux, df_pcs, df_specs, df_poc):
        if df is not None:
            df.columns = df.columns.astype(str).str.strip()
    return {"bat": df_bat, "skid": df_skid, "aux": df_aux, "pcs": df_pcs, "specs": df_specs,
            "skid_poc": df_poc, "foaie_poc": nume_poc,
            "sursa": sursa, "mtime": os.path.getmtime(sursa),
            "mtime_specs": os.path.getmtime(sursa_specs) if sursa_specs else None}


def fixeaza_antet(df):
    """Foaia POC are pe rândul 1 'Euro-USD' și antetul abia pe rândul 2: căutăm rândul cu 'Type'."""
    if any(str(c).strip().lower() == "type" for c in df.columns):
        return df
    for i, rand in df.iterrows():
        valori = [str(v).strip() for v in rand.values]
        if any(v.lower() == "type" for v in valori):
            nou = df.loc[i + 1:].copy()
            nou.columns = [v if v.lower() != "nan" else f"col{j}" for j, v in enumerate(valori)]
            return nou.reset_index(drop=True)
    return df


def tensiune_kv(text):
    """'SES4H-5000-EX-MV 10KV SF6 switchgear' -> 10; '10MW 30KV SF6' -> 30; fără kV în nume -> None."""
    m = re.search(r"(\d+(?:[.,]\d+)?)\s*kv\b", str(text), re.I)
    return float(m.group(1).replace(",", ".")) if m else None


def tensiuni_poc(df_poc):
    """Tensiunile (kV) disponibile în foaia POC, sortate."""
    if df_poc is None:
        return []
    c_tip = df_poc.columns[SKID_COL_TYPE]
    return sorted({kv for kv in (tensiune_kv(t) for t in df_poc[c_tip].dropna()) if kv})


def skiduri_pentru_poc(date, poc_kv):
    """DataFrame-ul de skiduri pentru punctul de conexiune ales."""
    if abs(poc_kv - POC_IMPLICIT_KV) < 1e-9:
        return date["skid"]
    df_poc = date.get("skid_poc")
    if df_poc is None:
        raise ValueError("Nu am găsit în Pricing.xlsx foaia 'MV SKID for different POC'.")
    c_tip = df_poc.columns[SKID_COL_TYPE]
    gasite = df_poc[df_poc[c_tip].map(lambda t: tensiune_kv(t) == poc_kv)]
    if gasite.empty:
        disponibile = ", ".join(f"{kv:g} kV" for kv in tensiuni_poc(df_poc)) or "niciuna"
        raise ValueError(f"Nu avem niciun skid pentru conexiune la {poc_kv:g} kV în foaia "
                         f"'{date.get('foaie_poc')}'.\n\nTensiuni disponibile: {POC_IMPLICIT_KV} kV (România), "
                         f"{disponibile}.")
    return gasite


def pricing_modificat(date):
    """True dacă Pricing.xlsx sau Specificatii_baterii_BESS.xlsx s-au schimbat de la ultima citire."""
    try:
        if date is None or os.path.getmtime(cale_pricing()) != date["mtime"]:
            return True
        sursa_specs = cale_specs()
        return (os.path.getmtime(sursa_specs) if sursa_specs else None) != date["mtime_specs"]
    except OSError:
        return False


def specs_baterie(df_specs, tip):
    """{'model', 'config', 'weight'} pentru bateria din Pricing, sau None dacă nu e în Specificatii_baterii_BESS.xlsx.
    Numele din Pricing (ex. 'REPT 5,015') se caută și în coloana 'Model' (ex. 'JKE-5015K-2H-LAA'),
    și în coloana 'Producător / Produs'. 'model' = codul real din coloana Model (ex. 'Y104R04C12')."""
    if df_specs is None or df_specs.empty:
        return None
    coloane = {c.lower(): c for c in df_specs.columns}
    c_model = coloane.get("model", df_specs.columns[0])
    c_produs = next((c for c in df_specs.columns if "produs" in c.lower() or "produc" in c.lower()), None)
    cheie = lambda x: re.sub(r"\s+", " ", curata_tip(x)).lower()

    for _, rand in df_specs.iterrows():
        nume = [rand[c_model]] + ([rand[c_produs]] if c_produs else [])
        if cheie(tip) not in {cheie(n) for n in nume if not pd.isna(n)}:
            continue

        def val(nume_col):
            v = rand.get(coloane.get(nume_col.lower()))
            if v is None or (not isinstance(v, str) and pd.isna(v)):
                return ""
            return v.strip() if isinstance(v, str) else f"{v:g}"

        greutate = rand.get(coloane.get("weight"))
        if isinstance(greutate, (int, float)) and not pd.isna(greutate):
            greutate = f"~ {greutate:g}".replace(".", ",")          # 44.5 -> '~ 44,5'
        else:
            greutate = val("Weight")
        return {"model": val("Model") or tip, "config": val("Configuration"), "weight": greutate}
    return None


def bani_eu(valoare, zecimale=2):
    """Format european pentru prezentare: 1742.68 -> '1.742,68'; 2186646.5 -> '2.186.646,50'."""
    text = f"{numar(valoare):,.{zecimale}f}"
    return text.replace(",", "#").replace(".", ",").replace("#", ".")


def nr_eu(valoare):
    """Număr fără zecimale inutile, cu virgulă: 5.015 -> '5,015'; 20.06 -> '20,06'; 10.0 -> '10'."""
    return f"{round(numar(valoare), 4):g}".replace(".", ",")


def text_ani(valoare):
    """5 -> '5 years'; text ('5 years', '10 yrs') rămâne cum e; gol -> ''."""
    if valoare is None or (not isinstance(valoare, str) and pd.isna(valoare)):
        return ""
    if isinstance(valoare, str):
        return valoare.strip()
    return f"{float(valoare):g} years"


def intrari_acceptate(nr_skiduri, min_unit, max_unit):
    """'2/4 intrări' pe skid = exact 2 SAU exact 4 (nu 3). Toate skidurile din proiect au aceeași
    configurație, deci totalurile posibile sunt doar n*min și n*max (ex. 2 skiduri: 4 sau 8, nu 6)."""
    return sorted({nr_skiduri * min_unit, nr_skiduri * max_unit})


def text_intrari(lista):
    return " / ".join(f"{x:g}" for x in lista)


def dc_per_baterie(brand_skid):
    for cheie, val in DC_INPUTURI_PER_BRAND.items():
        if cheie.lower() in str(brand_skid).lower():
            return val
    return DC_INPUTURI_PER_BATERIE


def lista_branduri(df_skid):
    col = df_skid.columns[SKID_COL_BRAND]
    return [str(x) for x in df_skid[col].dropna().unique() if str(x).strip()]


def curata_tip(valoare):
    """'JKE-\\n4597K-\\n2H-LAA' -> 'JKE-4597K-2H-LAA'."""
    return re.sub(r"\s*\n\s*", "", str(valoare)).strip()


def parseaza_data(valoare):
    """Extrage data din celula 'Quotation date' (poate conține și text, ex. '19-Oct-26 NEW Agreement')."""
    if valoare is None or pd.isna(valoare):
        return None
    if isinstance(valoare, datetime):
        return pd.Timestamp(valoare)
    text = str(valoare)
    m = re.search(r"\d{1,2}[-/. ][A-Za-z]{3}[-/. ]\d{2,4}", text)
    d = pd.to_datetime(m.group(0) if m else text, errors="coerce", dayfirst=True)
    return None if pd.isna(d) else d


def lista_baterii(df_bat):
    """{model: index rând}. Pentru fiecare model se păstrează DOAR cea mai nouă cotație."""
    c_tip = df_bat.columns[BAT_COL_TYPE]
    c_data = df_bat.columns[BAT_COL_DATA]
    cele_mai_noi = {}
    for idx, rand in df_bat.iterrows():
        tip = curata_tip(rand[c_tip])
        if not tip or tip.lower() == "nan":
            continue
        brut = rand[c_data]
        data = parseaza_data(brut)
        cheie = (
            data is not None,
            data if data is not None else pd.Timestamp.min,
            "new" in str(brut).lower(),     # la dată egală, preferă 'NEW Agreement'
            idx,                            # apoi rândul mai de jos
        )
        if tip not in cele_mai_noi or cheie > cele_mai_noi[tip][0]:
            cele_mai_noi[tip] = (cheie, idx)
    return {tip: v[1] for tip, v in cele_mai_noi.items()}


def coloana_pret_baterie(df_bat, incoterm="DDP"):
    """DDP: coloana I = 'Price / MWh FOB with taxes and DDP from DAP Tax'.
    DAP: coloana H = 'Price / MWh ... with taxes (DAP)'. Căutate după antet, cu poziția fixă ca rezervă."""
    for i, col in enumerate(df_bat.columns):
        c = col.lower()
        if incoterm == "DAP" and "(dap)" in c:
            return i
        if incoterm == "DDP" and "ddp from dap tax" in c:
            return i
    return BAT_COL_PRET_DAP if incoterm == "DAP" else BAT_COL_PRET


def litera_coloana(i):
    return chr(ord("A") + i) if i < 26 else str(i + 1)


def alege_baterie(df_bat, idx, incoterm="DDP"):
    """Citește rândul ales din foaia Battery. Preț = coloana I (DDP) sau H (DAP), transport = coloana J."""
    rand = df_bat.loc[idx]
    tip = curata_tip(rand.iloc[BAT_COL_TYPE])

    cap = numar(rand.iloc[BAT_COL_CAP])               # coloana E (MWh)
    if cap <= 0:
        m = re.search(r"(\d{4})K", tip)
        cap = int(m.group(1)) / 1000 if m else CAPACITATE_BATERIE_FALLBACK

    i_pret = coloana_pret_baterie(df_bat, incoterm)
    pret = numar(rand.iloc[i_pret])
    if pret <= 0:
        raise ValueError(f"Bateria {tip} nu are preț {incoterm} în coloana {litera_coloana(i_pret)} "
                         f"'{df_bat.columns[i_pret]}' (rândul {idx + 2} din Excel).")

    data = parseaza_data(rand.iloc[BAT_COL_DATA])
    return {
        "rand": rand,
        "tip": tip,
        "brand": str(rand.iloc[BAT_COL_BRAND]),
        "cap": cap,
        "pret": pret,
        "incoterm": incoterm,
        "coloana_pret": litera_coloana(i_pret),
        "transport": numar(rand.iloc[BAT_COL_TRANSPORT]),   # coloana J
        "garantie": text_ani(rand.iloc[BAT_COL_GARANTIE]) if len(rand) > BAT_COL_GARANTIE else "",  # coloana U
        "rand_excel": idx + 2,
        "data": data.strftime("%d-%b-%y") if data is not None else "fără dată",
    }


def e_baterie_rack(baterie_text):
    return BATERIE_RACK.lower() in str(baterie_text).lower()


def mod_conectare_rack(brand_skid):
    """Pentru bateriile REPT: 'rack' / 'dc' după brandul skidului, sau None dacă skidul nu e compatibil."""
    for cheie, mod in SKIDURI_PENTRU_RACK.items():
        if cheie.lower() in str(brand_skid).lower():
            return mod
    return None


def variante_intrari(baterie_text, brand_skid, nr_baterii):
    """Listă de (nr_intrari, descriere_conexiune) posibile pentru combinația baterie + skid."""
    if e_baterie_rack(baterie_text) and mod_conectare_rack(brand_skid) == "rack":
        rackuri = nr_baterii * RACKURI_PER_BATERIE
        return [(math.ceil(rackuri / k), f"{rackuri} rack-uri, conectare 1:{k}") for k in RACKURI_PER_INTRARE]
    return [(nr_baterii * dc_per_baterie(brand_skid), "")]


def gaseste_skiduri(date, mw, nr_baterii, brand, nr_fix=None, baterie_text=""):
    """Returnează (optiuni_sortate_dupa_pret, motive_respingere).
    baterie_text = brand + model baterie (pentru regula REPT -> doar skiduri TBEA, conectare pe rack)."""
    df_skid, df_pcs = date["skid"], date["pcs"]
    c_brand = df_skid.columns[SKID_COL_BRAND]
    c_tip = df_skid.columns[SKID_COL_TYPE]
    c_putere = df_skid.columns[SKID_COL_POWER]
    c_pret = df_skid.columns[SKID_COL_PRET]
    c_trans = df_skid.columns[SKID_COL_TRANSPORT]
    c_com = df_skid.columns[SKID_COL_COM]
    c_gar = df_skid.columns[SKID_COL_GARANTIE] if len(df_skid.columns) > SKID_COL_GARANTIE else None

    # Legătura Pricing <-> PCS se face pe nume curățate (spații în plus la final făceau ca skiduri să dispară)
    def cheie(x):
        x = re.sub(r"\s*(without|no) aux\s*$", "", str(x).strip(), flags=re.I)
        return re.sub(r"\s+", " ", x).strip().lower()
    sk = df_skid.dropna(subset=[c_tip]).copy()
    sk["_cheie"] = sk[c_tip].map(cheie)
    pc = df_pcs.dropna(subset=[df_pcs.columns[0]]).copy()
    pc["_cheie"] = pc[df_pcs.columns[0]].map(cheie)
    pc["_putere"] = [numar(p.get("PCS power")) * numar(p.get("no / frame")) for _, p in pc.iterrows()]
    pcs_dupa_nume = {p["_cheie"]: p for _, p in pc.iterrows()}

    def rand_pcs(s):
        """Rândul din PCS după nume; dacă numele nu există acolo (ex. modele noi Kehua 'BCS...'),
        ia un skid din PCS cu același brand și aceeași putere."""
        if s["_cheie"] in pcs_dupa_nume:
            return pcs_dupa_nume[s["_cheie"]], ""
        cod = s["_cheie"].split(" ")[0]          # 'ses4h-5000-ex-mv 10kv sf6 switchgear' -> 'ses4h-5000-ex-mv'
        if cod in pcs_dupa_nume:
            return pcs_dupa_nume[cod], ""
        brand_s = str(s.get(c_brand, "")).strip().lower()
        putere_s = numar(s.get(c_putere))
        for _, p in pc.iterrows():
            if brand_s and brand_s in p["_cheie"] and putere_s > 0 and abs(p["_putere"] - putere_s) < 0.01:
                return p, str(p[df_pcs.columns[0]])
        for brand_cfg, (min_i, max_i) in INTRARI_FARA_PCS.items():
            if brand_cfg.lower() in brand_s and putere_s > 0:
                return pd.Series({"PCS power": putere_s, "no / frame": 1,
                                  "Min no of imputs/ skid": min_i, "Max no of imputs/ skid": max_i}), ""
        return None, ""

    optiuni = []
    motive = {"brand": 0, "putere": 0, "prea_putine": [], "prea_multe": [], "intermediare": [],
              "exacte_respinse": [], "respinse": [], "fara_pcs": [], "incompatibil": 0}

    randuri = []
    for _, s in sk.iterrows():
        p, similar = rand_pcs(s)
        if p is None:
            # nu există nici după nume, nici după brand + putere -> nu știm intrările DC
            motive["fara_pcs"].append({"tip": str(s[c_tip]), "pret_unit": numar(s[c_pret]),
                                       "brand": str(s.get(c_brand, ""))})
            continue
        r = pd.concat([s, p[[c for c in p.index if c not in s.index]]])
        r["_similar"] = similar
        randuri.append(r)
    if not randuri:
        raise ValueError("Niciun skid din Pricing.xlsx nu se potrivește cu 'PCS - MV skid.xlsx'.")

    for skid in randuri:
        if brand != "Oricare" and brand.lower() not in str(skid.get(c_brand, "")).lower():
            motive["brand"] += 1
            continue
        motive["brand_gasit"] = True
        if e_baterie_rack(baterie_text) and mod_conectare_rack(skid.get(c_brand, "")) is None:
            motive["incompatibil"] += 1            # REPT merge (deocamdată) doar pe TBEA
            continue

        # Puterea = putere PCS x nr. PCS pe cadru (din fișierul PCS); altfel coloana Power din Pricing.
        # (în Pricing.xlsx unele skiduri Sinexcel au puterea tastată greșit față de numele modelului)
        calc = numar(skid.get("PCS power")) * numar(skid.get("no / frame"))
        putere = calc if calc > 0 else numar(skid.get(c_putere))
        if putere <= 0:
            motive["putere"] += 1
            continue

        nr_skiduri = nr_fix if nr_fix else math.ceil(mw / putere)
        nume = str(skid[c_tip])
        max_unit = numar(skid.get("Max no of imputs/ skid"))
        min_unit = numar(skid.get("Min no of imputs/ skid"), 1)
        acceptate = intrari_acceptate(nr_skiduri, min_unit, max_unit)
        min_total, max_total = acceptate[0], acceptate[-1]

        pret = numar(skid.get(c_pret))
        variante = variante_intrari(baterie_text, skid.get(c_brand, ""), nr_baterii)
        potrivite = [v for v in variante if any(abs(v[0] - x) < 1e-9 for x in acceptate)]
        nr_intrari, conexiune = potrivite[0] if potrivite else variante[0]
        if not potrivite:
            if len(variante) > 1:
                nr_intrari = " sau ".join(f"{v[0]:g}" for v in variante)
                text = f"{nume} ({nr_skiduri} buc: {variante[0][1].split(',')[0]} -> {nr_intrari} intrări, " \
                       f"acceptă {text_intrari(acceptate)})"
                motive["respinse"].append({
                    "tip": nume, "nr": nr_skiduri, "putere_totala": nr_skiduri * putere,
                    "pret_total": nr_skiduri * pret,
                    "motiv": f"{nr_intrari} intrări (1:1 / 1:2), skidul acceptă {text_intrari(acceptate)}",
                })
                motive["intermediare"].append(text)
                if abs(nr_skiduri * putere - mw) < 0.01:
                    motive["exacte_respinse"].append(text)
                continue
            text = f"{nume} ({nr_skiduri} buc: {nr_intrari:g} intrări, acceptă {text_intrari(acceptate)})"
            motive["respinse"].append({
                "tip": nume, "nr": nr_skiduri, "putere_totala": nr_skiduri * putere,
                "pret_total": nr_skiduri * pret,
                "motiv": f"{nr_intrari:g} intrări DC, skidul acceptă {text_intrari(acceptate)}",
            })
            if nr_intrari < min_total:
                motive["prea_putine"].append(text)
            elif nr_intrari > max_total:
                motive["prea_multe"].append(text)
            else:
                motive["intermediare"].append(text)
            if abs(nr_skiduri * putere - mw) < 0.01:
                motive["exacte_respinse"].append(text)
            continue

        optiuni.append({
            "brand": str(skid.get(c_brand, "")),
            "tip": nume,
            "nr": nr_skiduri,
            "putere_unit": putere,
            "putere_totala": nr_skiduri * putere,
            "intrari": nr_intrari,
            "min_intrari": min_total,
            "max_intrari": max_total,
            "acceptate": text_intrari(acceptate),
            "pret_unit": pret,
            "pret_total": nr_skiduri * pret,
            "transport": numar(skid.get(c_trans)),
            "com": numar(skid.get(c_com)),
            "similar": skid.get("_similar", ""),
            "conexiune": conexiune,
            "putere_pcs": numar(skid.get("PCS power")),                         # din PCS - MV skid.xlsx
            "garantie": text_ani(skid.get(c_gar)) if c_gar else "",            # coloana M
        })

    optiuni.sort(key=lambda o: o["pret_total"])
    return optiuni, motive


def filtreaza_skiduri(optiuni, mw):
    """Primele opțiuni compatibile, de la cel mai ieftin (indiferent dacă puterea e exactă sau mai mare).
    Returnează (lista, exista_exact)."""
    exact = any(abs(o["putere_totala"] - mw) < 0.01 for o in optiuni)
    return optiuni[:NR_OPTIUNI_SKID], exact


def descriere_skid(o, mw=None):
    nume = o["tip"] if o["brand"].lower() in o["tip"].lower() else f"{o['brand']} {o['tip']}"
    marcaj = ""
    if mw is not None:
        marcaj = "  ✓ putere exactă" if abs(o["putere_totala"] - mw) < 0.01 else "  (putere mai mare decât cerută)"
    return (f"{nume}: {o['nr']} buc x {o['putere_unit']:g} MW = {o['putere_totala']:g} MW | "
            f"{o['pret_total']:,.0f} ${marcaj}\n    intrări DC: {o['intrari']:g} "
            f"(acceptă {o['acceptate']})"
            + (f" | {o['conexiune']}" if o.get("conexiune") else "")
            + (f" - limite luate de la '{o['similar']}'" if o.get("similar") else ""))


def raport_skiduri(optiuni, motive):
    """Text cu TOATE skidurile din Pricing: compatibile, respinse (cu motiv) și lipsă din fișierul PCS."""
    linii = ["COMPATIBILE (de la cel mai ieftin):"]
    linii += [f"  {o['pret_total']:>12,.0f} $   {o['nr']} x {o['tip']} = {o['putere_totala']:g} MW"
              for o in optiuni] or ["  (niciunul)"]
    linii += ["", "RESPINSE (număr de intrări DC neacceptat):"]
    linii += [f"  {r['pret_total']:>12,.0f} $   {r['nr']} x {r['tip']} = {r['putere_totala']:g} MW  ->  {r['motiv']}"
              for r in sorted(motive["respinse"], key=lambda r: r["pret_total"])] or ["  (niciunul)"]
    linii += ["", "IGNORATE (nu există în 'PCS - MV skid.xlsx', deci nu știu puterea / intrările):"]
    linii += [f"  {f['pret_unit']:>12,.0f} $/buc   {f['tip']}"
              for f in sorted(motive["fara_pcs"], key=lambda f: f["pret_unit"])] or ["  (niciunul)"]
    if motive["brand"]:
        linii += ["", f"{motive['brand']} skiduri eliminate de filtrul de brand."]
    if motive.get("incompatibil"):
        linii += ["", f"{motive['incompatibil']} skiduri eliminate: bateriile {BATERIE_RACK} merg "
                      f"deocamdată doar pe skiduri {' / '.join(SKIDURI_PENTRU_RACK)}."]
    return "\n".join(linii)


def mesaj_fara_skid(mw, nr_baterii, brand, motive):
    linii = [f"Niciun skid compatibil pentru {mw:g} MW și {nr_baterii} baterii "
             f"(brand: {brand}).", ""]
    if motive["brand"]:
        linii.append(f"- {motive['brand']} skiduri eliminate de filtrul de brand.")
    if brand != "Oricare" and not motive.get("brand_gasit"):
        linii.append(f"- În foaia folosită pentru acest punct de conexiune nu există niciun skid {brand} "
                     f"utilizabil.")
    ignorate = [f["tip"] for f in motive.get("fara_pcs", [])
                if brand == "Oricare" or brand.lower() in f.get("brand", "").lower()]
    if ignorate:
        linii.append(f"- {len(ignorate)} skiduri ignorate, nu sunt în 'PCS - MV skid.xlsx' (nu știu intrările DC): "
                     + "; ".join(ignorate[:5]))
    if motive.get("incompatibil"):
        linii.append(f"- {motive['incompatibil']} skiduri eliminate: bateriile {BATERIE_RACK} merg "
                     f"deocamdată doar pe skiduri {' / '.join(SKIDURI_PENTRU_RACK)}.")
    if motive["putere"]:
        linii.append(f"- {motive['putere']} skiduri fără putere validă în Excel.")
    if motive["prea_putine"]:
        linii.append("- Prea puține intrări DC pentru: " + "; ".join(motive["prea_putine"][:5]))
    if motive["prea_multe"]:
        linii.append("- Prea multe intrări DC pentru: " + "; ".join(motive["prea_multe"][:5]))
    if motive["intermediare"]:
        linii.append("- Număr de intrări neacceptat (skidul acceptă doar valorile listate): "
                     + "; ".join(motive["intermediare"][:5]))
    linii += ["", "Ce poți schimba: brandul (încearcă 'Oricare'), MW sau MWh."]
    return "\n".join(linii)


def kva_aux_baterie(cap_mwh):
    """Consumul auxiliar al unei baterii (kVA), din KVA_AUX_PER_BATERIE (ca LOOKUP din Excel)."""
    kva = KVA_AUX_PER_BATERIE[0][1]
    for prag, valoare in KVA_AUX_PER_BATERIE:
        if cap_mwh >= prag:
            kva = valoare
    return kva


def putere_trafo(detail):
    """'1250 KVA Aux transformer' -> 1250; '450kVA 690V/400V ...' -> 450; fără kVA în nume -> 0."""
    m = re.search(r"(\d+(?:[.,]\d+)?)\s*kva", str(detail), re.I)
    return float(m.group(1).replace(",", ".")) if m else 0


def alege_trafo(df_aux, nr_baterii, cap_baterie):
    """Consum aux total = nr. baterii x kVA / baterie. Pentru fiecare trafo din Pricing:
    nr. bucăți = consum / putere (rotunjit în sus), cost = bucăți x preț. Se alege cel mai ieftin."""
    def coloana(text):
        return next((c for c in df_aux.columns if text in c.lower()), None)
    c_detail, c_dap, c_exw = coloana("detail"), coloana("dap"), coloana("exw")
    if c_detail is None:
        raise ValueError("Foaia 'Auxiliary Transformer' nu are coloana 'Detail'.")

    kva_bat = kva_aux_baterie(cap_baterie)
    kva_total = nr_baterii * kva_bat
    variante = []
    for _, rand in df_aux.iterrows():
        nume = str(rand[c_detail]).strip()
        if TRAFO_FILTRU.lower() not in nume.lower():
            continue
        kva = putere_trafo(nume)
        # preț DAP (total, cu transport) dacă există, altfel EXW
        pret = (numar(rand[c_dap]) if c_dap else 0) or (numar(rand[c_exw]) if c_exw else 0)
        if kva <= 0 or pret <= 0:
            continue
        cant = math.ceil(kva_total / kva)
        variante.append({"nume": nume, "kva": kva, "pret": pret, "cant": cant, "cost": cant * pret})
    if not variante:
        raise ValueError(f"Nu am găsit niciun trafo cu '{TRAFO_FILTRU}', putere (kVA) în nume și preț "
                         "în foaia 'Auxiliary Transformer'.")
    variante.sort(key=lambda v: (v["cost"], v["cant"]))   # la cost egal: mai puține bucăți
    ales = dict(variante[0])
    ales.update(kva_baterie=kva_bat, kva_total=kva_total, variante=variante)
    return ales


def valideaza(intrari):
    """Returnează (erori, avertismente)."""
    erori, avert = [], []
    if not intrari["client"].strip() or not intrari["proiect"].strip():
        erori.append("Clientul și numele proiectului nu pot fi goale.")
    if intrari["mwh"] <= 0 or intrari["mw"] <= 0:
        erori.append("Capacitatea și puterea trebuie să fie pozitive.")
    else:
        ore = intrari["mwh"] / intrari["mw"]
        if not ORE_MIN <= ore <= ORE_MAX:
            avert.append(f"Raport MWh/MW = {ore:.2f} h (normal între {ORE_MIN} și {ORE_MAX}). Ai tastat corect?")
    if intrari["ems"] < 0 or intrari["eng"] < 0:
        erori.append("EMS și ENG nu pot fi negative.")
    if not ADAOS_MIN <= intrari["adaos"] <= ADAOS_MAX:
        erori.append(f"Adaosul trebuie să fie între {ADAOS_MIN} și {ADAOS_MAX}.")
    if not CURS_MIN <= intrari["curs"] <= CURS_MAX:
        erori.append(f"Cursul USD/EUR trebuie să fie între {CURS_MIN} și {CURS_MAX}.")
    return erori, avert


# ======================================================================
# EXCEL
# ======================================================================
def inlocuieste_branduri(sh, branduri, brand_ales):
    """În celulele cu text (nu formule) din foaie, înlocuiește orice brand din listă cu brandul ales.
    Ex.: 'Commissioning Sinexcel' -> 'Commissioning Kehua'."""
    branduri = [b for b in branduri if b and b.strip().lower() != brand_ales.strip().lower()]
    if not branduri or not brand_ales:
        return
    model = re.compile("|".join(re.escape(b.strip()) for b in sorted(branduri, key=len, reverse=True)), re.I)
    zona = sh.used_range
    formule = zona.formula
    if not isinstance(formule, (list, tuple)):
        formule = ((formule,),)
    elif formule and not isinstance(formule[0], (list, tuple)):
        formule = (formule,)
    for i, rand in enumerate(formule):
        for j, val in enumerate(rand):
            if isinstance(val, str) and val and not val.startswith("=") and model.search(val):
                zona[i, j].value = model.sub(brand_ales.strip(), val)


def calculeaza_in_excel(intrari, baterie, nr_baterii, skid, trafo, folder, branduri_bat=(), branduri_skid=()):
    """Completează template-ul, recalculează, salvează calculul într-un fișier nou. Returnează rezultatele."""
    app_xl = None
    wb = None
    try:
        # Pe Mac există o singură instanță Excel (a utilizatorului): nu o ascundem și nu o închidem dacă are alte fișiere
        app_xl = xw.App(visible=(sys.platform == "darwin"), add_book=False)
        wb = app_xl.books.open(os.path.abspath(FISIER_EXCEL_TEMPLATE))
        sh = wb.sheets["Sheet1"]

        # Antet: client, proiect, nr. proiect, data și ora calculului
        acum = datetime.now()
        sh.range("A1:B2").value = [["Client", intrari["client"]], ["Nume Proiect", intrari["proiect"]]]
        sh.range("D1:E3").value = [["Data și ora", acum], ["Nr. proiect", intrari["nr_proiect"]],
                                   ["Versiune", f"v{intrari['versiune']}"]]
        sh.range("E1").number_format = "dd.mm.yyyy hh:mm:ss"
        sh.range("A1:A2").font.bold = True
        sh.range("D1:D3").font.bold = True

        sh.range("A5").value = baterie["tip"]
        sh.range("B5").value = baterie["cap"]
        sh.range("C5").value = nr_baterii
        sh.range("D5").value = baterie["pret"]
        sh.range("F5").value = baterie["transport"]

        sh.range("A6").value = skid["tip"]
        sh.range("C6").value = skid["nr"]
        sh.range("D6").value = skid["pret_unit"]
        sh.range("F6").value = skid["transport"]
        sh.range("G12").formula = f"={skid['com']}*C6"

        # Trafo auxiliar: nr. bucăți (C7) îl calculează Excelul, în blocul 'CALCUL TRAFO AUXILIAR' (A16:B25)
        sh.range("A7").value = trafo["nume"] if trafo else "aux trafo"
        sh.range("B7").value = trafo["kva"] if trafo else None   # puterea trafo-ului ales (kVA)
        sh.range("D7").value = trafo["pret"] if trafo else 0
        sh.range("B17").value = 1 if trafo else 0
        sh.range("A24").value = [[prag, kva] for prag, kva in KVA_AUX_PER_BATERIE]

        sh.range("E9").value = intrari["ems"]
        sh.range("E10").value = intrari["eng"]
        sh.range("H3").value = intrari["adaos"]
        sh.range("I3").value = intrari["curs"]

        # Etichetele din template (ex. 'Commissioning Sinexcel') primesc brandurile alese
        inlocuieste_branduri(sh, branduri_skid, skid["brand"])
        inlocuieste_branduri(sh, branduri_bat, baterie["brand"])

        app_xl.calculate()
        rezultat = {
            "pret_ddp": sh.range("H13").value or 0,
            "pret_mwh": sh.range("H14").value or 0,
            "ems_yearly": sh.range("K5").value or 0,
            "ext_gar_bat": sh.range(CELULA_EXT_GAR_BAT).value or 0,
            "ext_gar_pcs": sh.range(CELULA_EXT_GAR_PCS).value or 0,
            "nr_trafo": sh.range("C7").value or 0,
        }

        # Fiecare calcul = un fișier Excel separat, cu data și ora în nume (nu se suprascrie nimic)
        cale = os.path.abspath(os.path.join(folder, "Calcul_{}_{}_{}_v{}_{}.xlsx".format(
            *(nume_fisier_valid(x) for x in (intrari["nr_proiect"], intrari["client"], intrari["proiect"])),
            intrari["versiune"], acum.strftime("%Y-%m-%d_%H.%M.%S"))))
        wb.save(cale)
        rezultat["cale_excel"] = cale
        return rezultat
    finally:
        try:
            if wb is not None:
                wb.close()
        except Exception:
            pass
        try:
            if app_xl is not None and (sys.platform != "darwin" or not app_xl.books):
                app_xl.quit()
        except Exception:
            pass


# ======================================================================
# POWERPOINT
# ======================================================================
def _paragrafe(shape):
    """Generator de paragrafe din forme, grupuri și tabele."""
    if shape.shape_type == 6 and hasattr(shape, "shapes"):          # grup
        for s in shape.shapes:
            yield from _paragrafe(s)
    if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
        yield from shape.text_frame.paragraphs
    if getattr(shape, "has_table", False) and shape.has_table:
        for rand in shape.table.rows:
            for celula in rand.cells:
                yield from celula.text_frame.paragraphs


def _inlocuieste_in_paragraf(p, inlocuiri):
    if not any(k in p.text for k in inlocuiri):
        return
    # 1) întâi încercăm run cu run (păstrează formatarea)
    for r in p.runs:
        for k, v in inlocuiri.items():
            if k in r.text:
                r.text = r.text.replace(k, v)
    # 2) dacă un placeholder a fost spart pe mai multe run-uri, unim textul în primul run
    if any(k in p.text for k in inlocuiri) and p.runs:
        text = p.text
        for k, v in inlocuiri.items():
            text = text.replace(k, v)
        p.runs[0].text = text
        for r in p.runs[1:]:
            r.text = ""


def _colectii_shapes(prs):
    """Formele din slide-uri + layout-uri + master (data/logo sunt adesea acolo), fără dubluri."""
    vazute, colectii = set(), []
    for slide in prs.slides:
        for obj in (slide, slide.slide_layout, slide.slide_layout.slide_master):
            if id(obj.part) in vazute and obj is not slide:
                continue
            vazute.add(id(obj.part))
            colectii.append(obj.shapes)
    return colectii


def _seteaza_data_fixa(p, data_text):
    """Dacă paragraful conține DOAR o dată (ex. '15/10/26'), o înlocuiește cu data de azi."""
    if re.fullmatch(r"\s*\d{1,2}/\d{1,2}/\d{2,4}\s*", p.text) and p.runs:
        p.runs[0].text = data_text
        for r in p.runs[1:]:
            r.text = ""


def urmatoarea_versiune(nr_proiect, proiect):
    """Caută ofertele deja generate pentru același nr. proiect + proiect (folderele din FOLDER_OFERTE și,
    pentru ofertele mai vechi, cele de lângă program) și returnează următoarea versiune (1 dacă nu e niciuna)."""
    model = re.compile(re.escape(f"{nume_fisier_valid(nr_proiect)}_BESS_Indicative_Proposal_Taliva_"
                                 f"{nume_fisier_valid(proiect)}_") + r".*_v(\d+)(\.pptx)?$", re.I)
    maxim = 0
    foldere = [os.path.join(radacina, f) for radacina in (FOLDER_OFERTE, ".") if os.path.isdir(radacina)
               for f in os.listdir(radacina)]
    for folder in filter(os.path.isdir, foldere):
        try:
            continut = os.listdir(folder)
        except OSError:
            continut = []
        for nume in [os.path.basename(folder)] + continut:
            m = model.match(nume)
            if m:
                maxim = max(maxim, int(m.group(1)))
    return maxim + 1


def nume_fisier_valid(text):
    """Scoate caracterele interzise în numele de fișier pe Windows (\\ / : * ? " < > |)."""
    return re.sub(r'[\\/:*?"<>|]+', "-", str(text)).strip() or "Proiect"


def pptx_in_pdf(cale_pptx, cale_pdf):
    """Salvează prezentarea și ca PDF: prin PowerPoint (Windows: COM; Mac: AppleScript),
    iar dacă nu merge, prin LibreOffice (dacă e instalat)."""
    cale_pptx, cale_pdf = os.path.abspath(cale_pptx), os.path.abspath(cale_pdf)
    try:
        if sys.platform == "win32":
            _pdf_powerpoint_windows(cale_pptx, cale_pdf)
        elif sys.platform == "darwin":
            _pdf_powerpoint_mac(cale_pptx, cale_pdf)
        else:
            raise RuntimeError("PowerPoint nu există pe acest sistem.")
    except Exception as eroare:
        soffice = next((c for c in (shutil.which("soffice"),
                                    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
                                    r"C:\Program Files\LibreOffice\program\soffice.exe") if c and os.path.exists(c)), None)
        if soffice is None:
            raise eroare
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", os.path.dirname(cale_pdf),
                        cale_pptx], check=True, capture_output=True, timeout=300)


def _pdf_powerpoint_windows(cale_pptx, cale_pdf):
    import win32com.client                                   # pywin32 vine odată cu xlwings
    ppt = win32com.client.DispatchEx("PowerPoint.Application")
    prez = None
    try:
        prez = ppt.Presentations.Open(cale_pptx, ReadOnly=True, WithWindow=False)
        prez.SaveAs(cale_pdf, 32)                            # 32 = ppSaveAsPDF
    finally:
        if prez is not None:
            prez.Close()
        ppt.Quit()


def _pdf_powerpoint_mac(cale_pptx, cale_pdf):
    """PowerPoint pentru Mac, prin AppleScript (prima dată macOS cere permisiunea de a controla PowerPoint)."""
    text = lambda c: c.replace("\\", "\\\\").replace('"', '\\"')
    script = f'''
        tell application "Microsoft PowerPoint"
            open POSIX file "{text(cale_pptx)}"
            set prez to active presentation
            save prez in POSIX file "{text(cale_pdf)}" as save as PDF
            close prez saving no
        end tell'''
    rez = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=300)
    if rez.returncode != 0 or not os.path.exists(cale_pdf):
        raise RuntimeError(f"PowerPoint (Mac) nu a putut salva PDF-ul: {rez.stderr.strip()}")


def fara_extensie(text_brand):
    return any(b.lower() in str(text_brand).lower() for b in FARA_EXTENSIE_GARANTIE)


def ajusteaza_tabel_extensie(prs, fara_bat, fara_pcs):
    """Scoate din tabelul 'Warranty Extension' rândurile fără extensie de garanție și renumerotează
    coloana 'No'. Dacă nu rămâne niciun rând, șterge tot tabelul. Se face ÎNAINTE de înlocuirea textului."""
    if not (fara_bat or fara_pcs):
        return
    text_rand = lambda r: " ".join(c.text for c in r.cells)
    for slide in prs.slides:
        for shape in list(slide.shapes):
            if not getattr(shape, "has_table", False) or not shape.has_table:
                continue
            tabel = shape.table
            if not any("[EXT_GAR_" in text_rand(r) for r in tabel.rows):
                continue
            if fara_bat and fara_pcs:
                shape._element.getparent().remove(shape._element)
                continue
            for r in list(tabel.rows):
                t = text_rand(r)
                if (fara_bat and "[EXT_GAR_BAT]" in t) or (fara_pcs and "[EXT_GAR_PCS]" in t):
                    shape.height -= r.height
                    tabel._tbl.remove(r._tr)
            nr = 0
            for r in list(tabel.rows)[1:]:              # coloana 'No': 1, 2, ... fără goluri
                p = r.cells[0].text_frame.paragraphs[0]
                if re.fullmatch(r"\s*\d+\s*", p.text) and p.runs:
                    nr += 1
                    p.runs[0].text = str(nr)
                    for run in p.runs[1:]:
                        run.text = ""


def genereaza_pptx(inlocuiri, cale_final, fara_bat=False, fara_pcs=False):
    """Salvează prezentarea și returnează lista placeholderelor rămase neînlocuite."""
    prs = Presentation(FISIER_PPT_TEMPLATE)
    ajusteaza_tabel_extensie(prs, fara_bat, fara_pcs)
    data_text = inlocuiri.get("[DATA]", datetime.now().strftime(FORMAT_DATA))
    for shapes in _colectii_shapes(prs):
        for shape in shapes:
            for p in _paragrafe(shape):
                _inlocuieste_in_paragraf(p, inlocuiri)
                _seteaza_data_fixa(p, data_text)
    prs.save(cale_final)

    ramase = set()
    for slide in prs.slides:
        for shape in slide.shapes:
            for p in _paragrafe(shape):
                ramase.update(re.findall(r"\[[A-Za-z_]+\]", p.text))
    return sorted(ramase)


# ======================================================================
# INTERFAȚĂ
# ======================================================================
class AplicatieOferte:
    def __init__(self, root):
        self.root = root
        self.root.title("Generator Oferte BESS - Taliva")
        self.root.geometry("470x900")

        ttk.Label(root, text="Sistem Ofertare BESS", font=("Helvetica", 16, "bold")).pack(pady=10)
        frame = ttk.Frame(root)
        frame.pack(padx=20, fill="x")

        self.date = None
        self.baterii_map = {}
        baterii, branduri = [BATERIE_IMPLICITA], ["Oricare"]
        try:
            lipsa = fisiere_lipsa()
            if lipsa:
                raise FileNotFoundError(
                    "Programul trebuie să stea în folderul 'Ofertare BESS', lângă template-uri.\n\n"
                    f"Lipsesc din {os.getcwd()}:\n- " + "\n- ".join(lipsa))
            try:
                self.date = incarca_date()
            except PricingNegasit as e:
                # Pricing.xlsx negăsit automat: îl alege utilizatorul o singură dată, apoi e ținut minte
                messagebox.showinfo("Pricing.xlsx", f"{e}\n\nAlege fișierul Pricing.xlsx din folderul sincronizat.")
                cale = filedialog.askopenfilename(title="Alege Pricing.xlsx", filetypes=[("Excel", "*.xlsx")])
                if not cale:
                    raise
                salveaza_setari(dict(citeste_setari(), pricing=cale))
                self.date = incarca_date()
            self.baterii_map = lista_baterii(self.date["bat"])
            baterii = list(self.baterii_map) or baterii
            branduri = ["Oricare"] + self._branduri_skid()
        except Exception as e:
            messagebox.showwarning("Atenție", f"Nu am putut citi fișierele Excel la pornire:\n{e}")

        self.intrari = {}
        campuri = [
            ("Client:", "GEP"), ("Nume Proiect:", "Proiect X"),
            ("Nr. proiect:", NR_PROIECT_IMPLICIT.format(an_luna=datetime.now().strftime("%Y%m"))),
            ("Locație:", "București, România"),
            ("Capacitate (MWh):", "20"), ("Putere (MW):", "10"), ("Valoare EMS (USD):", "3500"),
            ("Valoare ENG (USD):", "0"), ("Adaos Comercial (Margin):", "1.10"), ("Curs USD/EUR:", "0.89"),
        ]
        for i, (eticheta, implicit) in enumerate(campuri):
            ttk.Label(frame, text=eticheta).grid(row=i, column=0, sticky="w", pady=5)
            e = ttk.Entry(frame, width=25)
            e.insert(0, implicit)
            e.grid(row=i, column=1, sticky="e", pady=5)
            self.intrari[eticheta] = e

        # EMS se completează automat în funcție de capacitate, până îl modifici tu manual
        self.ems_manual = False
        self._setare_auto = False
        self.ems_var = tk.StringVar(value=self.intrari["Valoare EMS (USD):"].get())
        self.cap_var = tk.StringVar(value=self.intrari["Capacitate (MWh):"].get())
        self.intrari["Valoare EMS (USD):"].config(textvariable=self.ems_var)
        self.intrari["Capacitate (MWh):"].config(textvariable=self.cap_var)
        self.ems_var.trace_add("write", self._la_editare_ems)
        self.cap_var.trace_add("write", self._la_schimbare_capacitate)
        self._la_schimbare_capacitate()

        r = len(campuri)
        ttk.Label(frame, text="Model baterie:").grid(row=r, column=0, sticky="w", pady=5)
        self.combo_baterie = ttk.Combobox(frame, values=baterii, state="readonly", width=32)
        self.combo_baterie.set(next((b for b in baterii if b.startswith(BATERIE_IMPLICITA)), baterii[0]))
        self.combo_baterie.grid(row=r, column=1, sticky="e", pady=5)

        ttk.Label(frame, text="Brand Skid preferat:").grid(row=r + 1, column=0, sticky="w", pady=5)
        self.combo_brand = ttk.Combobox(frame, values=branduri, state="readonly", width=22)
        self.combo_brand.current(0)
        self.combo_brand.grid(row=r + 1, column=1, sticky="e", pady=5)

        ttk.Label(frame, text="Nr. skiduri:").grid(row=r + 2, column=0, sticky="w", pady=5)
        self.combo_nr_skid = ttk.Combobox(frame, values=[AUTO_SKID] + [str(i) for i in range(1, 11)],
                                          state="readonly", width=22)
        self.combo_nr_skid.current(0)
        self.combo_nr_skid.grid(row=r + 2, column=1, sticky="e", pady=5)

        ttk.Label(frame, text="Livrare (Incoterms):").grid(row=r + 3, column=0, sticky="w", pady=5)
        self.combo_incoterm = ttk.Combobox(frame, values=list(INCOTERMS), state="readonly", width=22)
        self.combo_incoterm.set(INCOTERM_IMPLICIT)
        self.combo_incoterm.grid(row=r + 3, column=1, sticky="e", pady=5)

        ttk.Label(frame, text="Punct conexiune (POC):").grid(row=r + 4, column=0, sticky="w", pady=5)
        self.combo_poc = ttk.Combobox(frame, values=self._valori_poc(), width=22)
        self.combo_poc.current(0)
        self.combo_poc.grid(row=r + 4, column=1, sticky="e", pady=5)

        self.foloseste_aux = tk.BooleanVar()
        ttk.Checkbutton(frame, text="Include Transformator Auxiliar", variable=self.foloseste_aux)\
            .grid(row=r + 5, column=0, columnspan=2, pady=10, sticky="w")

        self.btn = tk.Button(root, text="Previzualizare și generare", bg="#0052cc", fg="white",
                             font=("Helvetica", 12, "bold"), command=self.previzualizeaza)
        self.btn.pack(pady=20, fill="x", padx=20)

        self.lbl_status = ttk.Label(root, text="Aștept date...", foreground="gray")
        self.lbl_status.pack()

        self.lbl_sursa = ttk.Label(root, text="", foreground="gray", wraplength=430)
        self.lbl_sursa.pack(pady=5)
        self._afiseaza_sursa()

    def _branduri_skid(self):
        """Brandurile din MV SKID + MV SKID for different POC, fără dubluri."""
        branduri = lista_branduri(self.date["skid"])
        if self.date.get("skid_poc") is not None:
            branduri += [b for b in lista_branduri(self.date["skid_poc"]) if b not in branduri]
        return branduri

    def _valori_poc(self):
        valori = [f"{POC_IMPLICIT_KV} kV (România)"]
        if self.date:
            valori += [f"{kv:g} kV" for kv in tensiuni_poc(self.date.get("skid_poc")) if kv != POC_IMPLICIT_KV]
        return valori

    def _afiseaza_sursa(self):
        if self.date:
            ora = datetime.fromtimestamp(self.date["mtime"]).strftime("%d/%m/%y %H:%M")
            self.lbl_sursa.config(text=f"Prețuri din: {self.date['sursa']}\n(modificat ultima dată: {ora})\n"
                                       f"Folder de lucru: {os.getcwd()}")

    def _la_editare_ems(self, *_):
        if not self._setare_auto:
            self.ems_manual = True            # utilizatorul a modificat EMS -> nu-l mai schimbăm automat

    def _la_schimbare_capacitate(self, *_):
        if self.ems_manual:
            return
        try:
            mwh = float(self.cap_var.get().replace(",", "."))
        except ValueError:
            return
        self._setare_auto = True
        try:
            self.ems_var.set(EMS_AUTO if mwh >= EMS_PRAG_MWH else EMS_SUB_PRAG)
        finally:
            self._setare_auto = False

    def status(self, text, culoare="blue"):
        self.lbl_status.config(text=text, foreground=culoare)
        self.root.update()

    def citeste_intrari(self):
        g = lambda k: self.intrari[k].get()
        try:
            return {
                "client": g("Client:"), "proiect": g("Nume Proiect:"), "locatie": g("Locație:"),
                "nr_proiect": g("Nr. proiect:").strip()
                              or NR_PROIECT_IMPLICIT.format(an_luna=datetime.now().strftime("%Y%m")),
                "mwh": float(g("Capacitate (MWh):").replace(",", ".")),
                "mw": float(g("Putere (MW):").replace(",", ".")),
                "ems": float(g("Valoare EMS (USD):").replace(",", ".")),
                "eng": float(g("Valoare ENG (USD):").replace(",", ".")),
                "adaos": float(g("Adaos Comercial (Margin):").replace(",", ".")),
                "curs": float(g("Curs USD/EUR:").replace(",", ".")),
                "poc_kv": float(re.match(r"\s*(\d+(?:[.,]\d+)?)", self.combo_poc.get()).group(1).replace(",", ".")),
            }
        except (ValueError, AttributeError):
            raise ValueError("Unul dintre câmpurile numerice (sau punctul de conexiune) nu conține un număr valid.")

    def previzualizeaza(self):
        try:
            self.status("Verific datele...")
            intrari = self.citeste_intrari()
            erori, avert = valideaza(intrari)
            if erori:
                raise ValueError("\n".join(erori))
            if avert and not messagebox.askyesno("Atenție", "\n".join(avert) + "\n\nContinui oricum?"):
                self.status("Anulat.", "gray")
                return

            # Recitim Pricing.xlsx doar dacă s-a modificat în Drive de la ultima citire
            if pricing_modificat(self.date):
                self.status("Pricing.xlsx s-a modificat, îl recitesc...")
                self.date = incarca_date()
                self.baterii_map = lista_baterii(self.date["bat"])
                self.combo_baterie["values"] = list(self.baterii_map)
                self.combo_brand["values"] = ["Oricare"] + self._branduri_skid()
                self.combo_poc["values"] = self._valori_poc()
                self._afiseaza_sursa()
            date = self.date

            if self.combo_baterie.get() not in self.baterii_map:
                raise ValueError("Selectează o baterie din listă.")
            baterie = alege_baterie(date["bat"], self.baterii_map[self.combo_baterie.get()],
                                    self.combo_incoterm.get())
            nr_baterii = math.ceil(intrari["mwh"] / baterie["cap"])
            nr_cerut = self.combo_nr_skid.get()
            nr_fix = None if nr_cerut == AUTO_SKID else int(nr_cerut)
            date_poc = dict(date, skid=skiduri_pentru_poc(date, intrari["poc_kv"]))
            optiuni, motive = gaseste_skiduri(date_poc, intrari["mw"], nr_baterii, self.combo_brand.get(), nr_fix,
                                              baterie_text=f"{baterie['brand']} {baterie['tip']}")
            if not optiuni:
                raise ValueError(mesaj_fara_skid(intrari["mw"], nr_baterii, self.combo_brand.get(), motive))

            specs = specs_baterie(date.get("specs"), baterie["tip"])
            if specs is None:
                unde = "Specificatii_baterii_BESS.xlsx nu a fost găsit" if date.get("specs") is None else \
                       f"modelul {baterie['tip']} nu e în Specificatii_baterii_BESS.xlsx"
                if not messagebox.askyesno("Atenție", f"{unde}, deci Configuration și Weight vor rămâne goale "
                                                      "în prezentare.\n\nContinui oricum?"):
                    self.status("Anulat.", "gray")
                    return
                specs = {"model": baterie["tip"], "config": "", "weight": ""}
            baterie["specs"] = specs

            trafo = alege_trafo(date["aux"], nr_baterii, baterie["cap"]) if self.foloseste_aux.get() else None

            lista, exact = filtreaza_skiduri(optiuni, intrari["mw"])
            alegere = self.dialog_previzualizare(intrari, baterie, nr_baterii, lista, trafo, exact,
                                                 raport_skiduri(optiuni, motive))
            if alegere is None:
                self.status("Anulat.", "gray")
                return
            self.genereaza(intrari, baterie, nr_baterii, lista[alegere], trafo)

        except Exception as e:
            self.status("Eroare!", "red")
            messagebox.showerror("Eroare", str(e))

    def dialog_previzualizare(self, intrari, baterie, nr_baterii, optiuni, trafo, exact=True, raport=""):
        """Returnează indexul opțiunii de skid alese sau None dacă se anulează."""
        win = tk.Toplevel(self.root)
        win.title("Previzualizare ofertă")
        win.transient(self.root)
        win.grab_set()

        cap_prop = nr_baterii * baterie["cap"]
        cost_bat = cap_prop * baterie["pret"]
        txt = (f"Cerut: {intrari['mwh']:g} MWh / {intrari['mw']:g} MW, conexiune {intrari['poc_kv']:g} kV\n"
               f"Baterii: {nr_baterii} x {baterie['tip']} = {cap_prop:.3f} MWh "
               f"\nLivrare: {INCOTERMS[baterie['incoterm']]}"
               f"\nPreț {baterie['incoterm']} (coloana {baterie['coloana_pret']}): "
               f"{baterie['pret']:,.2f} $/MWh (cotație {baterie['data']}, "
               f"rând Excel {baterie['rand_excel']}) -> baterii: {cost_bat:,.0f} $")
        if trafo:
            txt += (f"\nTrafo auxiliar: {nr_baterii} baterii x {trafo['kva_baterie']:g} kVA = "
                    f"{trafo['kva_total']:g} kVA -> {trafo['cant']} x {trafo['nume']} = {trafo['cost']:,.0f} $")
            alte = [v for v in trafo["variante"] if v["nume"] != trafo["nume"]]
            if alte:
                txt += "\n    alte variante: " + "; ".join(
                    f"{v['cant']} x {v['kva']:g} kVA = {v['cost']:,.0f} $" for v in alte)
        ttk.Label(win, text=txt, justify="left").pack(padx=15, pady=10, anchor="w")

        var = tk.IntVar(value=0)
        bold = ("Helvetica", 10, "bold")
        titlu = "Skiduri compatibile, de la cel mai ieftin (primul e selectat):"
        if not exact:
            titlu += f"\nNiciun skid compatibil nu are exact {intrari['mw']:g} MW."
        ttk.Label(win, text=titlu, font=bold, justify="left").pack(padx=15, anchor="w")
        for i, o in enumerate(optiuni):
            ttk.Radiobutton(win, variable=var, value=i, text=descriere_skid(o, intrari["mw"]))\
                .pack(padx=25, pady=3, anchor="w")

        rezultat = {"v": None}
        def ok():
            rezultat["v"] = var.get()
            win.destroy()

        def arata_toate():
            w = tk.Toplevel(win)
            w.title("Toate skidurile din Pricing.xlsx")
            t = tk.Text(w, width=130, height=30, font=("Consolas", 9))
            t.insert("1.0", raport)
            t.config(state="disabled")
            t.pack(padx=10, pady=10, fill="both", expand=True)

        bf = ttk.Frame(win)
        bf.pack(pady=15)
        tk.Button(bf, text="Generează", bg="#0052cc", fg="white", command=ok).pack(side="left", padx=8)
        tk.Button(bf, text="Vezi toate skidurile", command=arata_toate).pack(side="left", padx=8)
        tk.Button(bf, text="Anulează", command=win.destroy).pack(side="left", padx=8)
        self.root.wait_window(win)
        return rezultat["v"]

    def genereaza(self, intrari, baterie, nr_baterii, skid, trafo):
        # Folderul, prezentarea și PDF-ul au același nume, ex. ST202606109_BESS_..._20MWh_10MW_v2
        intrari["versiune"] = urmatoarea_versiune(intrari["nr_proiect"], intrari["proiect"])
        nume = NUME_PREZENTARE.format(nr_proiect=nume_fisier_valid(intrari["nr_proiect"]),
                                      proiect=nume_fisier_valid(intrari["proiect"]),
                                      mwh=nr_eu(intrari["mwh"]), mw=nr_eu(intrari["mw"])) \
               + f"_v{intrari['versiune']}"
        folder = os.path.abspath(os.path.join(FOLDER_OFERTE, nume))
        os.makedirs(folder, exist_ok=True)

        self.status("Calculez în Excel...")
        c_brand_bat = self.date["bat"].columns[BAT_COL_BRAND]
        branduri_bat = [str(x) for x in self.date["bat"][c_brand_bat].dropna().unique() if str(x).strip()]
        rez = calculeaza_in_excel(intrari, baterie, nr_baterii, skid, trafo, folder,
                                  branduri_bat, self._branduri_skid())

        self.status("Populez prezentarea PowerPoint...")
        inlocuiri = {
            "[DATA]": datetime.now().strftime(FORMAT_DATA),
            "[CLIENT]": intrari["client"],
            "[NUME_PROIECT]": intrari["proiect"],
            "[NR_PROIECT]": intrari["nr_proiect"],
            "[VERSIUNE]": str(intrari["versiune"]),
            "[LOCATIE]": intrari["locatie"],
            "[MWH]": nr_eu(intrari["mwh"]),
            "[MW]": nr_eu(intrari["mw"]),
            "[PROP_MWH]": nr_eu(nr_baterii * baterie["cap"]),
            "[PROP_MW]": nr_eu(skid["putere_totala"]),
            "[BAT_BRAND]": baterie["brand"],
            "[BAT_NAME]": baterie["specs"]["model"],      # codul real (ex. Y104R04C12), nu 'REPT 5,015'
            "[NR_BAT]": str(nr_baterii),
            "[SKID_BRAND]": skid["brand"],
            "[SKID]": skid["tip"],
            "[NR_SKID]": str(skid["nr"]),
            "[PCS_POWER]": nr_eu(skid["putere_pcs"]),                        # ex. 1,25
            "[PRET_TOTAL_USD]": f"{bani_eu(rez['pret_ddp'])} $",
            "[PRET_MWh_USD]": f"{bani_eu(rez['pret_mwh'])} $/MWh",
            "[EMS_YEARLY]": f"{bani_eu(rez['ems_yearly'])} $",
            "[INCOTERMS]": INCOTERMS[baterie["incoterm"]],
            "[POC]": nr_eu(intrari["poc_kv"]),                                # ex. 20 (unitatea kV e în tabel)
            "[BAT_CAP]": nr_eu(baterie["cap"]),                               # ex. 5,015
            "[BAT_CONFIG]": baterie["specs"]["config"],
            "[BAT_WEIGHT]": baterie["specs"]["weight"],
            "[GAR_BAT]": baterie["garantie"],
            "[GAR_PCS]": skid["garantie"],
            "[EXT_GAR_BAT]": bani_eu(rez["ext_gar_bat"]),
            "[EXT_GAR_PCS]": bani_eu(rez["ext_gar_pcs"]),
        }
        cale = os.path.join(folder, nume + ".pptx")
        ramase = genereaza_pptx(inlocuiri, cale,
                                fara_bat=fara_extensie(f"{baterie['brand']} {baterie['tip']}"),
                                fara_pcs=fara_extensie(skid["brand"]))

        self.status("Salvez și varianta PDF...")
        eroare_pdf = None
        try:
            pptx_in_pdf(cale, os.path.join(folder, nume + ".pdf"))
        except Exception as e:
            eroare_pdf = str(e)

        self.status("Finalizat cu succes!", "green")
        mesaj = (f"Fișierele au fost salvate în:\n\n📂 {folder}\n\n"
                 f"{os.path.basename(rez['cale_excel'])}\n{nume}.pptx")
        mesaj += f"\n{nume}.pdf" if eroare_pdf is None else \
                 f"\n\n⚠ PDF-ul nu s-a putut crea (e nevoie de PowerPoint instalat):\n{eroare_pdf}"
        if ramase:
            mesaj += "\n\n⚠ Placeholdere rămase neînlocuite în prezentare:\n" + ", ".join(ramase)
        if eroare_pdf or ramase:
            messagebox.showwarning("Gata, dar...", mesaj)
        else:
            messagebox.showinfo("Gata!", mesaj)


if __name__ == "__main__":
    root = tk.Tk()
    os.chdir(alege_folder_lucru())   # PCS, specificații, template-uri și 'Oferte generate' sunt în 'Ofertare BESS'
    gui = AplicatieOferte(root)
    root.mainloop()
