# -*- coding: utf-8 -*-
"""Logica de ofertare: Pricing, baterii, skiduri, trafo aux, EMS.
Parte din generator_oferte_bess.py: se încarcă în același spațiu de nume (toate funcțiile
și setările sunt comune), deci aici NU se fac importuri proprii și nu se rulează direct."""
# pylint: disable=undefined-variable
# flake8: noqa

class PricingNegasit(FileNotFoundError):
    pass


def fisiere_lipsa(folder="."):
    """Fișierele care trebuie să stea în folderul 'Ofertare BESS' și lipsesc.
    PCS și specificațiile nu mai sunt necesare separat dacă sunt foi în 'Setari generator.xlsx'."""
    foi = foi_setari(os.path.join(folder, FISIER_SETARI_GENERATOR))
    necesare = [FISIER_EXCEL_TEMPLATE, FISIER_PPT_TEMPLATE] + \
        [f for f, foaie in ((FISIER_PCS, FOAIE_PCS_SETARI), (FISIER_SPECS, FOAIE_SPECS_SETARI)) if foaie not in foi]
    return [f for f in necesare if not os.path.exists(os.path.join(folder, f))]


def foi_setari(cale=None):
    """Numele foilor din 'Setari generator.xlsx' (listă goală dacă fișierul nu există)."""
    cale = cale or FISIER_SETARI_GENERATOR
    if not os.path.exists(cale):
        return []
    try:
        from openpyxl import load_workbook
        wb = load_workbook(cale, read_only=True)
        nume = list(wb.sheetnames)
        wb.close()
        return nume
    except Exception:
        return []


def _evalueaza_formula(formula, valoare_celula, adancime=0):
    """Formule simple (ex. '=B2*C2', '=D29/4'): referințe de celule și + - * / ( ). Altfel None.
    Fișierele salvate de program (nu de Excel) nu au valorile formulelor calculate, așa că le calculăm aici."""
    if adancime > 20:
        return None
    expr = formula.lstrip("=").replace("$", "")
    def inlocuieste(m):
        v = valoare_celula(m.group(1), int(m.group(2)), adancime + 1)
        if isinstance(v, bool) or not isinstance(v, (int, float)):
            raise ValueError
        return f"({v!r})"
    try:
        expr = re.sub(r"\b([A-Z]{1,3})(\d+)\b", inlocuieste, expr.upper())
        if not re.fullmatch(r"[\d\s.+\-*/()e]*", expr):
            return None
        return float(eval(expr, {"__builtins__": {}}, {}))
    except Exception:
        return None


def citeste_foaie_cu_formule(cale, foaie):
    """O foaie ca DataFrame (antet pe rândul 1), cu formulele simple calculate dacă lipsesc valorile din Excel."""
    from openpyxl import load_workbook
    with tempfile.TemporaryDirectory() as tmp:
        copie = os.path.join(tmp, os.path.basename(cale))
        shutil.copy2(cale, copie)
        sh_f = load_workbook(copie)[foaie]
        sh_v = load_workbook(copie, data_only=True)[foaie]

    def valoare(col, rand, adancime=0):
        v = sh_v[f"{col}{rand}"].value
        f = sh_f[f"{col}{rand}"].value
        if v is None and isinstance(f, str) and f.startswith("="):
            return _evalueaza_formula(f, valoare, adancime)
        return v

    randuri = []
    for r in range(1, sh_f.max_row + 1):
        randuri.append([valoare(litera_coloana(c - 1), r) for c in range(1, sh_f.max_column + 1)])
    antet = [str(v).strip() if v is not None else f"col{i}" for i, v in enumerate(randuri[0])] if randuri else []
    df = pd.DataFrame(randuri[1:], columns=antet).dropna(how="all")
    return df.loc[:, [c for c in df.columns if not (c.startswith("col") and df[c].isna().all())]]


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
    """Ordinea: folderul Purchasing sincronizat (găsit automat), copia din 'Ofertare BESS' (pusă acolo de fluxul
    Power Automate), iar la final calea aleasă manual (ținută minte). Cu Drive-ul conectat: copia descărcată de acolo."""
    if DRIVE is not None:
        return os.path.abspath(FISIER_PRICING)
    for cale in (cauta_pricing_in_onedrive(), FISIER_PRICING, citeste_setari().get("pricing")):
        if cale and os.path.exists(cale):
            return cale
    raise PricingNegasit(
        f"Nu găsesc {FISIER_PRICING}: nici în folderul '{NUME_FOLDER_PRICING}', nici în 'Ofertare BESS' "
        "(unde îl copiază automat fluxul Power Automate). Verifică dacă OneDrive rulează și dacă "
        "folderul 'Ofertare BESS' e sincronizat pe acest calculator.")


def cale_specs():
    """De unde se citesc specificațiile: 'Setari generator.xlsx' (foaia FOAIE_SPECS_SETARI),
    altfel Specificatii_baterii_BESS.xlsx, altfel None."""
    if FOAIE_SPECS_SETARI in foi_setari():
        return FISIER_SETARI_GENERATOR
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


def _amprenta(*cai):
    """(cale, data modificării, mărime) pentru fiecare fișier: dacă nu s-a schimbat nimic, folosim datele din cache."""
    rezultat = []
    for c in cai:
        try:
            st = os.stat(c)
            rezultat.append((os.path.abspath(c), st.st_mtime, st.st_size))
        except (OSError, TypeError):
            rezultat.append((c, None, None))
    return rezultat


FISIER_CACHE_DATE = os.path.join(os.path.dirname(FISIER_SETARI), "cache_date.pkl")


def incarca_date():
    """Datele din Pricing + setări (PCS, specificații). Dacă fișierele nu s-au schimbat de la ultima citire, se iau
    din cache (pornire mult mai rapidă); altfel se citesc din nou."""
    import pickle
    sursa = cale_pricing()
    amprenta = _amprenta(sursa, FISIER_SETARI_GENERATOR, FISIER_PCS, FISIER_SPECS)
    try:
        with open(FISIER_CACHE_DATE, "rb") as f:
            cache = pickle.load(f)
        if cache.get("amprenta") == amprenta and cache.get("versiune") == 1:
            return cache["date"]
    except Exception:
        pass
    date = _citeste_date(sursa)
    try:
        os.makedirs(os.path.dirname(FISIER_CACHE_DATE), exist_ok=True)
        with open(FISIER_CACHE_DATE + ".tmp", "wb") as f:
            pickle.dump({"amprenta": amprenta, "versiune": 1, "date": date}, f)
        os.replace(FISIER_CACHE_DATE + ".tmp", FISIER_CACHE_DATE)
    except Exception:
        pass
    return date


def _citeste_date(sursa):
    """Citește fișierele de pricing (doar foile folosite) și returnează un dicționar de DataFrame-uri."""
    with tempfile.TemporaryDirectory() as tmp:
        copie = os.path.join(tmp, os.path.basename(sursa))
        shutil.copy2(sursa, copie)
        with pd.ExcelFile(copie) as xls:
            toate = list(xls.sheet_names)
            folosite = [f for f in toate if f in (FOAIE_BATERIE, FOAIE_SKID, "Auxiliary Transformer")
                        or FOAIE_SKID_POC_CONTINE in f.lower()]
            foi = {f: pd.read_excel(xls, sheet_name=f) for f in folosite}
    foi.update({f: None for f in toate if f not in foi})            # numele celorlalte foi (pentru mesaje)
    lipsa = [f for f in (FOAIE_BATERIE, FOAIE_SKID, "Auxiliary Transformer") if f not in foi]
    if lipsa:
        raise ValueError(f"Fișierul de prețuri citit nu are foile: {', '.join(lipsa)}.\n\n"
                         f"Fișier citit:\n{os.path.abspath(sursa)}\n\nFoi găsite în el: {', '.join(foi)}\n\n"
                         "Probabil e o copie veche sau alt fișier: verifică dacă OneDrive a sincronizat "
                         "folderul Purchasing Prices Update_BESS.")
    df_bat, df_skid, df_aux = foi[FOAIE_BATERIE], foi[FOAIE_SKID], foi["Auxiliary Transformer"]
    nume_poc = next((f for f in foi if FOAIE_SKID_POC_CONTINE in f.lower()), None)
    df_poc = fixeaza_antet(foi[nume_poc]) if nume_poc else None
    foi_s = foi_setari()
    if FOAIE_PCS_SETARI in foi_s:
        df_pcs = citeste_foaie_cu_formule(FISIER_SETARI_GENERATOR, FOAIE_PCS_SETARI)
    elif os.path.exists(FISIER_PCS):
        df_pcs = citeste_copie(FISIER_PCS, ["Sheet1"])["Sheet1"]
    else:
        raise ValueError(f"Nu găsesc PCS-urile: nici foaia '{FOAIE_PCS_SETARI}' în {FISIER_SETARI_GENERATOR}, "
                         f"nici fișierul {FISIER_PCS} în Ofertare BESS.")
    sursa_specs = cale_specs()
    if sursa_specs == FISIER_SETARI_GENERATOR:
        df_specs = citeste_foaie_cu_formule(FISIER_SETARI_GENERATOR, FOAIE_SPECS_SETARI)
    else:
        df_specs = citeste_copie(sursa_specs, None) if sursa_specs else None
    for df in (df_bat, df_skid, df_aux, df_pcs, df_specs, df_poc):
        if df is not None:
            df.columns = df.columns.astype(str).str.strip()
    return {"bat": df_bat, "skid": df_skid, "aux": df_aux, "pcs": df_pcs, "specs": df_specs,
            "skid_poc": df_poc, "foaie_poc": nume_poc,
            "sursa": sursa, "mtime": os.path.getmtime(sursa),
            "mtime_specs": os.path.getmtime(sursa_specs) if sursa_specs else None}


# Coloanele din Pricing folosite după poziție: (cheie din date, foaia, [(index, ce conține, tip valori)]).
# Tip "numar" = valorile din coloană trebuie să fie în mare parte numere; "text" = în mare parte text.
COLOANE_PRICING = [
    ("bat", FOAIE_BATERIE, [(BAT_COL_BRAND, "Brand baterie", "text"), (BAT_COL_TYPE, "Model (Type)", "text"),
                            (BAT_COL_CAP, "Capacitate (MWh)", "numar"), (BAT_COL_PRET_DAP, "Preț DAP / MWh", "numar"),
                            (BAT_COL_PRET, "Preț DDP / MWh", "numar"), (BAT_COL_TRANSPORT, "Transport", "numar"),
                            (BAT_COL_COM[0], "Commissioning 1 container", "numar")]),
    ("skid", FOAIE_SKID, [(SKID_COL_BRAND, "Brand skid", "text"), (SKID_COL_TYPE, "Model (Type)", "text"),
                          (SKID_COL_PRET, "Preț", "numar"),
                          (SKID_COL_TRANSPORT, "Transport", "numar"), (SKID_COL_COM, "Commissioning", "numar")]),
    ("skid_poc", "MV SKID for different POC",
     [(SKID_COL_BRAND, "Brand skid", "text"), (SKID_COL_TYPE, "Model (Type)", "text"),
      (SKID_COL_PRET, "Preț", "numar"), (SKID_COL_TRANSPORT, "Transport", "numar")]),
]
# Texte care trebuie să apară în antetul unor coloane (verificare sigură, independentă de istoric)
ANTETE_OBLIGATORII = {("bat", BAT_COL_COM[0]): "commis", ("bat", BAT_COL_TYPE): "type", ("bat", BAT_COL_PRET_DAP): "dap", ("bat", BAT_COL_PRET): "ddp",
                      ("skid", SKID_COL_TYPE): "type", ("skid_poc", SKID_COL_TYPE): "type"}


def _e_numar(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return not pd.isna(v)
    try:
        float(str(v).replace(",", "."))
        return True
    except ValueError:
        return False


def verifica_coloane_pricing(date):
    """Verifică dacă coloanele din Pricing au rămas pe pozițiile folosite de program.
    Returnează (probleme_sigure, antete_schimbate, antete_actuale):
    - probleme_sigure: antet fără textul obligatoriu, coloană lipsă sau valori de alt tip (ex. text în loc de preț);
    - antete_schimbate: antete diferite față de ultima verificare acceptată (ținută minte în setari.json)."""
    probleme, actuale = [], {}
    for cheie, foaie, coloane in COLOANE_PRICING:
        df = date.get(cheie)
        if df is None:
            continue
        nume_foaie = date.get("foaie_poc") if cheie == "skid_poc" else foaie
        c_tip = df.columns[SKID_COL_TYPE if cheie != "bat" else BAT_COL_TYPE] if len(df.columns) > 3 else None
        randuri = df.dropna(subset=[c_tip]) if c_tip is not None else df
        for idx, descriere, tip in coloane:
            litera = litera_coloana(idx)
            if idx >= len(df.columns):
                probleme.append(f"'{nume_foaie}', coloana {litera} ({descriere}): lipsește.")
                continue
            antet = str(df.columns[idx]).strip()
            actuale[f"{cheie}:{litera}"] = antet
            obligatoriu = ANTETE_OBLIGATORII.get((cheie, idx))
            if obligatoriu and obligatoriu not in antet.lower():
                probleme.append(f"'{nume_foaie}', coloana {litera} ({descriere}): antetul e '{antet}', "
                                f"mă așteptam să conțină '{obligatoriu}'.")
                continue
            valori = [v for v in randuri.iloc[:, idx] if not (not isinstance(v, str) and pd.isna(v))
                      and str(v).strip()]
            if len(valori) < 3:
                continue
            numere = sum(_e_numar(v) for v in valori)
            if tip == "numar" and numere < 0.6 * len(valori):
                probleme.append(f"'{nume_foaie}', coloana {litera} ({descriere}, antet '{antet}'): "
                                f"doar {numere} din {len(valori)} valori sunt numere.")
            elif tip == "text" and numere > 0.5 * len(valori):
                probleme.append(f"'{nume_foaie}', coloana {litera} ({descriere}, antet '{antet}'): "
                                f"conține în mare parte numere, nu text.")
    vechi = citeste_setari().get("antete_pricing") or {}
    schimbate = [f"{k.split(':')[0].replace('skid_poc', 'MV SKID POC').replace('bat', FOAIE_BATERIE).replace('skid', FOAIE_SKID)}"
                 f", coloana {k.split(':')[1]}: '{vechi[k]}' → '{v}'"
                 for k, v in actuale.items() if k in vechi and vechi[k] != v]
    return probleme, schimbate, actuale


def accepta_antete_pricing(actuale):
    salveaza_setari(dict(citeste_setari(), antete_pricing=actuale))


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


def doar_poc_implicit(brand):
    return any(b.lower() in str(brand).lower() for b in BRANDURI_DOAR_POC_IMPLICIT)


def skiduri_pentru_poc(date, poc_kv):
    """DataFrame-ul de skiduri pentru punctul de conexiune ales."""
    if abs(poc_kv - POC_IMPLICIT_KV) < 1e-9:
        return date["skid"]
    df_poc = date.get("skid_poc")
    if df_poc is None:
        raise ValueError("Nu am găsit în Pricing.xlsx foaia 'MV SKID for different POC'.")
    c_tip, c_brand = df_poc.columns[SKID_COL_TYPE], df_poc.columns[SKID_COL_BRAND]
    gasite = df_poc[df_poc[c_tip].map(lambda t: tensiune_kv(t) == poc_kv)]
    gasite = gasite[~gasite[c_brand].map(lambda b: doar_poc_implicit(b))]
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
        "com_unit": [numar(rand.iloc[i]) for i in BAT_COL_COM if i < len(rand)],   # K..O: 1, 2, 3, 4, 5+ containere
        "garantie": text_ani(rand.iloc[BAT_COL_GARANTIE]) if len(rand) > BAT_COL_GARANTIE else "",  # coloana U
        "rand_excel": idx + 2,
        "data": data.strftime("%d-%b-%y") if data is not None else "fără dată",
        "data_cotatie": data,
    }


def commissioning_baterie(baterie, nr_baterii):
    """Commissioning / baterie din Pricing, după numărul de containere (1, 2, 3, 4, 5+); None = nu e în Pricing."""
    com = [c for c in (baterie.get("com_unit") or [])]
    if not com or not any(com):
        return None
    return com[min(max(int(nr_baterii), 1), len(com)) - 1]


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


def cheie_skid(x):
    """Numele skidului curățat, pentru legătura Pricing <-> PCS - MV skid.xlsx."""
    x = re.sub(r"\s*(without|no) aux\s*$", "", str(x).strip(), flags=re.I)
    return re.sub(r"\s+", " ", x).strip().lower()


def gaseste_skiduri(date, mw, nr_baterii, brand, nr_fix=None, baterie_text="", kva_baterie=0.0,
                    forteaza_trafo=False):
    """Returnează (optiuni_sortate_dupa_cost_total, motive_respingere).
    baterie_text = brand + model baterie (pentru regula REPT -> doar skiduri TBEA, conectare pe rack).
    Pentru fiecare skid: necesar aux = baterii x kva_baterie + skiduri x consum aux skid; ce nu acoperă trafo-ul
    aux din skid se acoperă cu trafo extern, iar costul lui intră în comparație (cost_total).
    forteaza_trafo: True = trafo extern și când skidul acoperă tot (dimensionat pentru tot necesarul);
    "fara" = niciodată trafo extern (bifa „Fără trafo aux extern”); False = automat."""
    trafo_pricing = variante_trafo(date["aux"]) if date.get("aux") is not None else []
    df_skid, df_pcs = date["skid"], date["pcs"]
    c_brand = df_skid.columns[SKID_COL_BRAND]
    c_tip = df_skid.columns[SKID_COL_TYPE]
    c_putere = df_skid.columns[SKID_COL_POWER]
    c_pret = df_skid.columns[SKID_COL_PRET]
    c_trans = df_skid.columns[SKID_COL_TRANSPORT]
    c_com = df_skid.columns[SKID_COL_COM]
    c_gar = df_skid.columns[SKID_COL_GARANTIE] if len(df_skid.columns) > SKID_COL_GARANTIE else None

    # Legătura Pricing <-> PCS se face pe nume curățate (spații în plus la final făceau ca skiduri să dispară)
    cheie = cheie_skid
    sk = df_skid.dropna(subset=[c_tip]).copy()
    sk["_cheie"] = sk[c_tip].map(cheie)
    pc = df_pcs.dropna(subset=[df_pcs.columns[0]]).copy()
    pc["_cheie"] = pc[df_pcs.columns[0]].map(cheie)
    pc["_putere"] = [numar(p.get("PCS power")) * numar(p.get("no / frame")) for _, p in pc.iterrows()]
    pcs_dupa_nume = {p["_cheie"]: p for _, p in pc.iterrows()}
    chei_pricing = set(sk["_cheie"])

    def rand_pcs(s):
        """Rândul din PCS după nume; dacă numele nu există acolo (ex. modele noi Kehua 'BCS...'),
        ia un skid din PCS cu același brand și aceeași putere."""
        if s["_cheie"] in pcs_dupa_nume:
            return pcs_dupa_nume[s["_cheie"]], ""
        cod = s["_cheie"].split(" ")[0]          # 'ses4h-5000-ex-mv 10kv sf6 switchgear' -> 'ses4h-5000-ex-mv'
        if cod in pcs_dupa_nume:
            return pcs_dupa_nume[cod], ""
        # același skid cu sufix în PCS: 'kehua 2.5mw mv station - sf6free' -> '... +100 aux trafo' (dacă în Pricing
        # nu există și varianta cu sufix, ar fi altă linie) -> legătură exactă, deci se ia și trafo-ul aux din PCS
        lungi = [k for k in pcs_dupa_nume if k.startswith(s["_cheie"] + " ") and k not in chei_pricing]
        if len(lungi) == 1:
            return pcs_dupa_nume[lungi[0]], ""
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
        nominala = numar(skid.get("Skid power"))
        if nominala > 0 and any(b.lower() in str(skid.get(c_brand, "")).lower() for b in PUTERE_NOMINALA_SKID):
            calc = nominala
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
        acceptate_text = text_intrari(acceptate)
        rack = e_baterie_rack(baterie_text) and mod_conectare_rack(skid.get(c_brand, "")) == "rack"
        pcs_pe_skid = numar(skid.get("no / frame"))
        if pcs_pe_skid <= 0 and numar(skid.get("PCS power")) > 0:
            pcs_pe_skid = round(putere / numar(skid.get("PCS power")))
        if not rack and pcs_pe_skid > 0:
            # Regula PCS: 1 baterie -> 1 PCS sau 1 baterie -> 2 PCS
            nr_pcs = nr_skiduri * int(pcs_pe_skid)
            raport = next((k for k in PCS_PE_BATERIE if nr_pcs == k * nr_baterii), None)
            acceptate_text = " sau ".join(f"1 baterie : {k} PCS" for k in PCS_PE_BATERIE)
            if raport is None:
                text = f"{nume} ({nr_skiduri} buc: {nr_baterii} baterii pe {nr_pcs} PCS)"
                motive["respinse"].append({
                    "tip": nume, "nr": nr_skiduri, "putere_totala": nr_skiduri * putere,
                    "pret_total": nr_skiduri * pret,
                    "motiv": f"{nr_baterii} baterii pe {nr_pcs} PCS: nu e {acceptate_text}",
                })
                motive["intermediare"].append(text)
                if abs(nr_skiduri * putere - mw) < 0.01:
                    motive["exacte_respinse"].append(text)
                continue
            nr_intrari = nr_pcs
            conexiune = f"{nr_baterii} baterii -> {nr_pcs} PCS (1 baterie : {raport} PCS)"
        else:
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
            "acceptate": acceptate_text,
            "regula_pcs": not rack and pcs_pe_skid > 0,
            "pret_unit": pret,
            "pret_total": nr_skiduri * pret,
            "transport": numar(skid.get(c_trans)),
            "com": numar(skid.get(c_com)),
            "similar": skid.get("_similar", ""),
            "conexiune": conexiune,
            "putere_pcs": numar(skid.get("PCS power")),                         # din PCS - MV skid.xlsx
            "garantie": text_ani(skid.get(c_gar)) if c_gar else "",            # coloana M
            "data_cotatie": parseaza_data(skid.iloc[SKID_COL_DATA]) if len(skid) > SKID_COL_DATA else None,
        })
        calcul_aux(optiuni[-1], aux_inclus_skid(skid, nume), nr_baterii, kva_baterie, forteaza_trafo, trafo_pricing)

    optiuni.sort(key=lambda o: o["cost_total"])
    return optiuni, motive


def filtreaza_skiduri(optiuni, mw):
    """Primele opțiuni compatibile, de la cel mai ieftin (indiferent dacă puterea e exactă sau mai mare).
    Returnează (lista, exista_exact)."""
    exact = any(abs(o["putere_totala"] - mw) < 0.01 for o in optiuni)
    return optiuni[:NR_OPTIUNI_SKID], exact


def _din_trepte(trepte, putere):
    trepte = sorted(trepte)
    for mw, kva in trepte:
        if putere <= mw + 1e-9:
            return kva
    mw, kva = trepte[-1]
    return math.ceil(kva * putere / mw)                   # peste ultima treaptă: proporțional


def consum_aux_skid(putere_mw):
    """kVA consum propriu al unui skid, după putere (KVA_CONSUM_SKID)."""
    return _din_trepte(KVA_CONSUM_SKID, putere_mw)


def calcul_aux(o, aux_skid, nr_baterii, kva_baterie, forteaza_trafo, trafo_pricing):
    """Completează opțiunea de skid cu necesarul aux, ce acoperă skidul, trafo-ul extern și costul total."""
    o["aux_skid"] = aux_skid
    o["kva_skid"] = consum_aux_skid(o["putere_unit"]) if CONSUM_SKID_IN_NECESAR else 0.0
    o["kva_baterii"] = nr_baterii * kva_baterie
    o["kva_necesar"] = o["kva_baterii"] + o["nr"] * o["kva_skid"]
    o["kva_acoperit"] = o["nr"] * aux_skid
    o["kva_lipsa"] = max(0.0, o["kva_necesar"] - o["kva_acoperit"])
    o["fara_trafo"] = forteaza_trafo == "fara"
    if o["fara_trafo"]:
        o["kva_trafo"] = 0.0                                         # exclus manual
    elif forteaza_trafo and o["kva_lipsa"] == 0:
        o["kva_trafo"] = o["kva_necesar"]
    else:
        o["kva_trafo"] = o["kva_lipsa"]
    o["trafo"] = alege_trafo(trafo_pricing, o["kva_trafo"])
    o["trafo_lipsa_din_pricing"] = o["kva_trafo"] > 0 and o["trafo"] is None
    o["cost_total"] = o["pret_total"] + (o["trafo"]["cost"] if o["trafo"] else 0)


def text_aux(o):
    """'aux: necesar 250 kVA (4 baterii + 1 skid), skid acoperă 200 kVA -> trafo extern: 1 x 400 KVA ... = 30.000 $'."""
    text = (f"aux: necesar {o['kva_necesar']:g} kVA (baterii {o['kva_baterii']:g}"
            + (f" + skid {o['nr']} x {o['kva_skid']:g} kVA), " if o["kva_skid"] else
               "; skidul are trafo propriu pentru consumul lui), ")
            + (f"skidul acoperă {o['kva_acoperit']:g} kVA" if o["kva_acoperit"] else "skid fără trafo aux") + " -> ")
    if o.get("fara_trafo"):
        text += ("fără trafo extern (exclus manual" +
                 (f"; ⚠ rămân neacoperiți {o['kva_lipsa']:g} kVA)" if o["kva_lipsa"] else ")"))
    elif o["trafo"]:
        text += f"trafo extern: {o['trafo']['cant']} x {o['trafo']['nume']} = {o['trafo']['cost']:,.0f} $"
    elif o.get("trafo_lipsa_din_pricing"):
        text += f"⚠ e nevoie de trafo extern ({o['kva_trafo']:g} kVA), dar nu e niciunul în Pricing"
    else:
        text += "nu e nevoie de trafo extern"
    return text


def zile_vechime(data, acum=None):
    """Câte zile are cotația (None dacă nu are dată)."""
    if data is None:
        return None
    try:
        return ((acum or datetime.now()) - data).days
    except TypeError:
        return None


def avertisment_cotatie(data, ce):
    """'⚠ cotația skidului are 75 de zile' dacă e mai veche de ZILE_COTATIE_VECHE."""
    zile = zile_vechime(data)
    return f"  ⚠ cotația {ce} are {zile} de zile" if zile is not None and zile > ZILE_COTATIE_VECHE else ""


def procent_fata_de(cost, minim):
    """'+3,2%' față de cea mai ieftină variantă ('cea mai ieftină' pentru ea)."""
    if not minim or cost - minim < 0.5:
        return "cea mai ieftină"
    return f"+{nr_eu(round((cost / minim - 1) * 100, 1))}%"


def descriere_skid(o, mw=None):
    nume = o["tip"] if o["brand"].lower() in o["tip"].lower() else f"{o['brand']} {o['tip']}"
    marcaj = ""
    if mw is not None:
        marcaj = "  ✓ putere exactă" if abs(o["putere_totala"] - mw) < 0.01 else \
            f"  (putere mai mare decât cerută: limitată prin EMS la {nr_eu(mw)} MW)"
    pret = (f"skid {o['pret_total']:,.0f} $ + trafo {o['trafo']['cost']:,.0f} $ = {o['cost_total']:,.0f} $"
            if o.get("trafo") else f"{o['pret_total']:,.0f} $")
    if o.get("manual"):
        return (f"{nume}: {o['nr']} buc x {o['putere_unit']:g} MW = {o['putere_totala']:g} MW | {pret}{marcaj}"
                f"\n    ✎ skid introdus manual - compatibilitatea cu bateriile NU e verificată de program"
                f"\n    {text_aux(o)}")
    if o.get("regula_pcs"):
        legare = f"\n    conectare: {o['conexiune']}"
    else:
        legare = (f"\n    intrări DC: {o['intrari']:g} (acceptă {o['acceptate']})"
                  + (f" | {o['conexiune']}" if o.get("conexiune") else ""))
    return (f"{nume}: {o['nr']} buc x {o['putere_unit']:g} MW = {o['putere_totala']:g} MW | {pret}{marcaj}"
            + legare
            + (f" - limite luate de la '{o['similar']}'" if o.get("similar") else "")
            + f"\n    {text_aux(o)}")


def raport_skiduri(optiuni, motive):
    """Text cu TOATE skidurile din Pricing: compatibile, respinse (cu motiv) și lipsă din fișierul PCS."""
    linii = ["COMPATIBILE (de la cel mai ieftin, skid + trafo aux extern dacă e nevoie):"]
    linii += [f"  {o['cost_total']:>12,.0f} $   {o['nr']} x {o['tip']} = {o['putere_totala']:g} MW   ({text_aux(o)})"
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
        linii.append("- Conectare baterii - skid neacceptată (vezi motivul la fiecare): "
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


def variante_trafo(df_aux):
    """Trafo-urile aux din Pricing: {'nume', 'kva', 'pret'}. Puterea din nume, prețul DAP (total, cu transport),
    altfel EXW. Listă goală dacă nu există."""
    def coloana(text):
        return next((c for c in df_aux.columns if text in c.lower()), None)
    c_detail, c_dap, c_exw = coloana("detail"), coloana("dap"), coloana("exw")
    if c_detail is None:
        return []
    variante = []
    for _, rand in df_aux.iterrows():
        nume = str(rand[c_detail]).strip()
        if TRAFO_FILTRU.lower() not in nume.lower():
            continue
        kva = putere_trafo(nume)
        pret = (numar(rand[c_dap]) if c_dap else 0) or (numar(rand[c_exw]) if c_exw else 0)
        if kva > 0 and pret > 0:
            variante.append({"nume": nume, "kva": kva, "pret": pret})
    return variante


def alege_trafo(variante, kva_necesar):
    """Trafo-ul extern pentru kva_necesar: bucăți = kVA / putere (rotunjit în sus), cost = bucăți x preț.
    Implicit cât mai puține bucăți, apoi cel mai ieftin (TRAFO_CAT_MAI_PUTINE).
    None dacă nu e nevoie de trafo (sau nu există niciunul în Pricing)."""
    if kva_necesar <= 0 or not variante:
        return None
    optiuni = [dict(v, cant=math.ceil(kva_necesar / v["kva"])) for v in variante]
    for o in optiuni:
        o["cost"] = o["cant"] * o["pret"]
    if TRAFO_CAT_MAI_PUTINE:
        optiuni.sort(key=lambda v: (v["cant"], v["cost"]))   # cât mai puține trafo-uri, apoi cel mai ieftin
    else:
        optiuni.sort(key=lambda v: (v["cost"], v["cant"]))   # cel mai ieftin; la cost egal, mai puține bucăți
    ales = dict(optiuni[0])
    ales.update(kva_total=kva_necesar, variante=optiuni)
    return ales


def aux_inclus_skid(skid, nume):
    """kVA trafo aux inclus într-un skid: coloana COLOANA_AUX_PCS din 'PCS - MV skid.xlsx' (doar dacă skidul e
    găsit acolo după nume, nu prin potrivire 'similară'), altfel din nume ('+ AUX TRAFO 200kVA'), altfel 0."""
    if not skid.get("_similar"):
        for coloana in skid.index:
            if str(coloana).strip().lower() == COLOANA_AUX_PCS.lower() and numar(skid[coloana]) > 0:
                return numar(skid[coloana])
    m = (re.search(r"aux\D{0,12}?(\d+(?:[.,]\d+)?)\s*kva", str(nume), re.I)        # '+ AUX TRAFO 200kVA'
         or re.search(r"(\d+(?:[.,]\d+)?)\s*kva\s*aux", str(nume), re.I))          # '+ 200kVA  AUX'
    return float(m.group(1).replace(",", ".")) if m else 0.0


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
