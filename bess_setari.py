# -*- coding: utf-8 -*-
"""Setări, curs BCE, actualizare automată, Setari generator.xlsx.
Parte din generator_oferte_bess.py: se încarcă în același spațiu de nume (toate funcțiile
și setările sunt comune), deci aici NU se fac importuri proprii și nu se rulează direct."""
# pylint: disable=undefined-variable
# flake8: noqa

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


# ---------------- cursul USD/EUR (Banca Centrală Europeană) ----------------
URL_CURS = "https://www.ecb.europa.eu/stats/eurofxref/eurofxref-daily.xml"


def curs_din_bce(xml):
    """Din XML-ul BCE ('1 EUR = 1,1625 USD', data 2026-11-10) -> (0.8602, '10.11.2026'): EUR pentru 1 USD."""
    usd = re.search(r"currency=['\"]USD['\"]\s+rate=['\"]([\d.]+)", xml)
    data = re.search(r"time=['\"](\d{4})-(\d{2})-(\d{2})", xml)
    if not usd:
        raise ValueError("BCE: cursul USD lipsește din răspuns.")
    return round(1 / float(usd.group(1)), 4), (f"{data.group(3)}.{data.group(2)}.{data.group(1)}" if data else "")


def curs_salvat():
    """Ultimul curs descărcat (setari.json): (curs, data) sau (CURS_IMPLICIT, None)."""
    s = citeste_setari()
    try:
        curs = float(s["curs"])
        if CURS_MIN <= curs <= CURS_MAX:
            return curs, s.get("curs_data")
    except (KeyError, TypeError, ValueError):
        pass
    return CURS_IMPLICIT, None


def actualizeaza_curs(timeout=8):
    """Descarcă cursul BCE și îl ține minte. Returnează (curs, data); fără internet -> excepție."""
    curs, data = curs_din_bce(_descarca_text(URL_CURS, timeout))
    if not CURS_MIN <= curs <= CURS_MAX:
        raise ValueError(f"Curs neobișnuit de la BCE: {curs}")
    salveaza_setari(dict(citeste_setari(), curs=curs, curs_data=data))
    return curs, data


def locatie_cu_tara(locatie, tara, tari=()):
    """Schimbă doar țara din locație: 'Sinești, Romania' + Moldova -> 'Sinești, Moldova'; '' -> 'Moldova'."""
    locatie = str(locatie or "").strip()
    if not locatie or locatie in tari:
        return tara
    if "," in locatie:
        return locatie.rsplit(",", 1)[0].rstrip() + ", " + tara
    return f"{locatie}, {tara}"


# ---------------- actualizare automată de pe GitHub ----------------
def _descarca_text(url, timeout=15):
    cerere = urllib.request.Request(url + f"?t={int(time.time())}", headers={"Cache-Control": "no-cache"})
    with urllib.request.urlopen(cerere, timeout=timeout) as r:
        return r.read().decode("utf-8")


def _normalizat(text):
    return text.replace("\r\n", "\n")


MINUTE_VERIFICARE_ACTUALIZARE = 10       # cu programul deschis, verifică din nou la atâtea minute


def cauta_actualizare():
    """{nume fișier: conținut nou} pentru fișierele programului care diferă de cele de pe GitHub; {} = la zi.
    Fără internet / GitHub indisponibil -> {} (programul merge normal)."""
    if getattr(sys, "frozen", False) or os.path.isdir(os.path.join(FOLDER_PROGRAM, ".git")):
        return {}                                   # .exe sau folderul de dezvoltare: nu se actualizează
    try:
        principal = _descarca_text(URL_ACTUALIZARE + FISIERE_PROGRAM[0])
        m = re.search(r"^FISIERE_PROGRAM = \(([^)]*)\)", principal, re.M)
        fisiere = re.findall(r'"([^"]+)"', m.group(1)) if m else [FISIERE_PROGRAM[0]]
        noi = {}
        for nume in fisiere:
            text = principal if nume == FISIERE_PROGRAM[0] else _descarca_text(URL_ACTUALIZARE + nume)
            cale = os.path.join(FOLDER_PROGRAM, nume)
            try:
                with open(cale, encoding="utf-8") as f:
                    vechi = f.read()
            except OSError:
                vechi = None
            if vechi is None or _normalizat(vechi) != _normalizat(text):
                noi[nume] = text
        return noi
    except Exception:
        return {}


def aplica_actualizare(noi):
    """Scrie fișierele noi (cele vechi rămân ca .bak). Verifică întâi că tot codul nou e corect."""
    for nume, text in noi.items():
        compile(text, nume, "exec")                 # SyntaxError -> nu se schimbă nimic
    for nume, text in noi.items():
        cale = os.path.join(FOLDER_PROGRAM, nume)
        os.makedirs(os.path.dirname(cale), exist_ok=True)
        if os.path.exists(cale):
            shutil.copy2(cale, cale + ".bak")
        with open(cale + ".nou", "w", encoding="utf-8", newline="\n") as f:
            f.write(text)
        os.replace(cale + ".nou", cale)


def reporneste_programul():
    subprocess.Popen([sys.executable, os.path.join(FOLDER_PROGRAM, FISIERE_PROGRAM[0])] + sys.argv[1:],
                     cwd=FOLDER_PROGRAM)


# ---------------- Setari generator.xlsx ----------------
def _nr_text(x):
    """Număr scris în stil românesc: 5.6 -> '5,6', 55.0 -> '55'."""
    return f"{x:g}".replace(".", ",")


def valoare_setare_text(valoare, tip):
    """Valoarea din cod, cum se scrie în Setari generator.xlsx."""
    if tip == "da_nu":
        return "DA" if valoare else "NU"
    if tip in ("numar", "intreg"):
        return float(valoare) if tip == "numar" else int(valoare)
    if tip == "text":
        return str(valoare)
    if tip in ("lista", "lista_intregi"):
        return ", ".join(str(v) for v in valoare)
    if tip in ("praguri", "trepte"):
        return "; ".join(f"{_nr_text(prag)} = {_nr_text(kva)}" for prag, kva in valoare)
    return "; ".join(f"{k} = {_nr_text(v) if isinstance(v, (int, float)) else v}" for k, v in valoare.items())


def _numar_setare(text):
    if isinstance(text, (int, float)) and not isinstance(text, bool):
        return float(text)
    return float(str(text).strip().replace(" ", "").replace(",", "."))


class _FormatSetare(ValueError):
    pass


def citeste_valoare_setare(valoare, tip):
    """Valoarea din Setari generator.xlsx -> valoarea pentru cod. Ridică ValueError cu explicație."""
    try:
        if tip == "numar":
            return _numar_setare(valoare)
        if tip == "intreg":
            n = _numar_setare(valoare)
            if n != int(n):
                raise ValueError
            return int(n)
        text = str(valoare).strip()
        if tip == "text":
            return text
        if tip == "da_nu":
            if text.upper() in ("DA", "YES", "TRUE", "1", "ADEVARAT"):
                return True
            if text.upper() in ("NU", "NO", "FALSE", "0", "FALS"):
                return False
            raise ValueError
        bucati = [b.strip() for b in re.split(r"[;,]" if tip.startswith("lista") else r";", text) if b.strip()]
        if tip == "lista":
            return tuple(bucati)
        if tip == "lista_intregi":
            return tuple(citeste_valoare_setare(b, "intreg") for b in bucati)
        perechi = []
        for b in bucati:
            cheie, egal, val = b.partition("=")
            if not egal or not cheie.strip() or not val.strip():
                raise ValueError
            perechi.append((cheie.strip(), val.strip()))
        if tip in ("praguri", "trepte"):
            praguri = sorted((_numar_setare(k), _numar_setare(v)) for k, v in perechi)
            if not praguri or (tip == "praguri" and praguri[0][0] != 0):
                raise _FormatSetare("primul prag trebuie să fie 0")
            if any(mw <= 0 for mw, _ in praguri) and tip == "trepte":
                raise _FormatSetare("puterile trebuie să fie pozitive")
            return praguri
        if tip == "perechi":
            return {k: _numar_setare(v) for k, v in perechi}
        return dict(perechi)
    except (ValueError, TypeError) as e:
        detaliu = f" ({e})" if isinstance(e, _FormatSetare) else ""
        formate = {"da_nu": "DA sau NU", "numar": "un număr (ex. 1,1)", "intreg": "un număr întreg", "lista": "texte separate prin virgulă",
                   "lista_intregi": "numere întregi separate prin virgulă (ex. 1, 2)",
                   "perechi": "perechi 'nume = număr' separate prin ; (ex. REPT = 4)",
                   "perechi_text": "perechi 'nume = text' separate prin ; (ex. TBEA = rack; Kehua = dc)",
                   "praguri": "praguri 'de la MWh = kVA' separate prin ; (ex. 0 = 55; 5,6 = 65)",
                   "trepte": "trepte 'până la MW = kVA' separate prin ; (ex. 2 = 30; 5 = 40)"}
        raise ValueError(f"trebuie să fie {formate.get(tip, tip)}{detaliu}") from None


def _verifica_setari_noi(noi):
    """Verificări între setări (după citire). Returnează lista de erori."""
    v = {nume: noi.get(nume, globals()[nume]) for nume, _, tip in SETARI_EDITABILE if tip}
    erori = []
    if not v["ADAOS_MIN"] <= v["ADAOS_IMPLICIT"] <= v["ADAOS_MAX"]:
        erori.append("ADAOS_IMPLICIT trebuie să fie între ADAOS_MIN și ADAOS_MAX.")
    if not v["CURS_MIN"] <= v["CURS_IMPLICIT"] <= v["CURS_MAX"]:
        erori.append("CURS_IMPLICIT trebuie să fie între CURS_MIN și CURS_MAX.")
    if v["ORE_MIN"] >= v["ORE_MAX"]:
        erori.append("ORE_MIN trebuie să fie mai mic decât ORE_MAX.")
    if str(v["INCOTERM_IMPLICIT"]).upper() not in INCOTERMS:
        erori.append(f"INCOTERM_IMPLICIT trebuie să fie unul dintre: {', '.join(INCOTERMS)}.")
    if any(m.lower() not in ("rack", "dc") for m in v["SKIDURI_PENTRU_RACK"].values()):
        erori.append("SKIDURI_PENTRU_RACK: fiecare brand trebuie să fie 'rack' sau 'dc'.")
    try:
        v["TEXT_LIMITARE_EMS"].format(propus="6", limitat="5,85")
    except (KeyError, IndexError, ValueError):
        erori.append("TEXT_LIMITARE_EMS poate folosi doar {propus} și {limitat}.")
    if not v["PCS_PE_BATERIE"] or any(n <= 0 for n in v["PCS_PE_BATERIE"]):
        erori.append("PCS_PE_BATERIE trebuie să conțină numere pozitive (ex. 1, 2).")
    return erori


def aplica_setari_generator(cale=None):
    """Citește Setari generator.xlsx (dacă există) și înlocuiește regulile din cod.
    Returnează (nr. setări schimbate, lista de probleme). Rândurile cu greșeli rămân cu valoarea din cod."""
    cale = cale or FISIER_SETARI_GENERATOR
    if not os.path.exists(cale):
        return 0, []
    try:
        df = citeste_copie(cale, [FOAIE_REGULI])[FOAIE_REGULI]
    except Exception as e:
        return 0, [f"Nu am putut citi foaia '{FOAIE_REGULI}': {e}"]
    tipuri = {nume: tip for nume, _, tip in SETARI_EDITABILE if tip}
    noi, probleme = {}, []
    for i, rand in df.iterrows():
        nume = str(rand.iloc[0]).strip()
        if nume not in tipuri:
            continue
        valoare = rand.iloc[2] if len(rand) > 2 else None
        if valoare is None or (not isinstance(valoare, str) and pd.isna(valoare)) or \
                (isinstance(valoare, str) and not valoare.strip() and tipuri[nume] != "text"):
            continue                                            # gol -> rămâne valoarea din cod
        if isinstance(valoare, str) and not valoare.strip():
            valoare = ""
        try:
            noi[nume] = citeste_valoare_setare(valoare, tipuri[nume])
        except ValueError as e:
            probleme.append(f"Rândul {i + 2} ({nume}): '{valoare}' {e}.")
    erori = _verifica_setari_noi(noi)
    if erori:
        return 0, probleme + erori + ["Nu am aplicat nicio setare din fișier până nu sunt corectate."]
    if "SKIDURI_PENTRU_RACK" in noi:
        noi["SKIDURI_PENTRU_RACK"] = {k: v.lower() for k, v in noi["SKIDURI_PENTRU_RACK"].items()}
    if "INCOTERM_IMPLICIT" in noi:
        noi["INCOTERM_IMPLICIT"] = noi["INCOTERM_IMPLICIT"].upper()
    schimbate = sum(1 for nume, val in noi.items() if globals()[nume] != val)
    globals().update(noi)
    return schimbate, probleme


def scrie_setari_generator(cale):
    """Creează Setari generator.xlsx cu valorile folosite acum de program."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    wb = Workbook()
    sh = wb.active
    sh.title = FOAIE_REGULI
    sh.append(["Cheie (nu se modifică)", "Setare", "Valoare", "Format"])
    formate = {"da_nu": "DA / NU", "numar": "număr", "intreg": "număr întreg", "text": "text", "lista": "texte separate prin virgulă",
               "lista_intregi": "numere separate prin virgulă", "perechi": "nume = număr; nume = număr",
               "perechi_text": "nume = rack/dc; ...", "praguri": "de la MWh = kVA; ...",
               "trepte": "până la MW = kVA; ..."}
    for nume, eticheta, tip in SETARI_EDITABILE:
        if tip is None:
            sh.append(["", nume])
            sh.cell(sh.max_row, 2).font = Font(bold=True, size=12)
            continue
        sh.append([nume, eticheta, valoare_setare_text(globals()[nume], tip), formate[tip]])
        sh.cell(sh.max_row, 1).font = Font(color="808080")
        sh.cell(sh.max_row, 3).fill = PatternFill("solid", fgColor="DDEBF7")
        sh.cell(sh.max_row, 3).alignment = Alignment(horizontal="left")
    for celula in sh[1]:
        celula.font = Font(bold=True, color="FFFFFF")
        celula.fill = PatternFill("solid", fgColor="1F4E78")
    for col, latime in zip("ABCD", (28, 58, 28, 34)):
        sh.column_dimensions[col].width = latime
    sh.freeze_panes = "A2"
    sh.append([])
    sh.append(["", "Se modifică doar coloana Valoare (albastru). Gol = valoarea din program. "
                   "Modificările se aplică la următoarea pornire / previzualizare."])
    wb.save(cale)


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
