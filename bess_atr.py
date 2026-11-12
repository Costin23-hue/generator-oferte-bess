# -*- coding: utf-8 -*-
"""Cititorul de ATR / studii de soluție / caiete de sarcini.
Parte din generator_oferte_bess.py: se încarcă în același spațiu de nume (toate funcțiile
și setările sunt comune), deci aici NU se fac importuri proprii și nu se rulează direct."""
# pylint: disable=undefined-variable
# flake8: noqa

# ======================================================================
# CITIRE ATR / STUDIU DE SOLUȚIE (fără Claude: tipare pe modelul ANRE)
# ======================================================================
def text_pdf(cale):
    """Textul unui PDF (pypdf). Gol dacă PDF-ul e scanat (doar imagini)."""
    try:
        import pypdf
    except ImportError:
        raise RuntimeError("Lipsește biblioteca pypdf (pentru citirea PDF-urilor).") from None
    cititor = pypdf.PdfReader(cale)
    return "\n".join((pag.extract_text() or "") for pag in cititor.pages)


def _nr_ro(text):
    """'4100,000' -> 4100.0; '4.100,5' -> 4100.5; '7.700' (mii) -> 7700.0; '7,7' -> 7.7."""
    t = str(text).strip().replace(" ", "")
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"\d{1,3}(\.\d{3})+", t):
        t = t.replace(".", "")
    return float(t)


_NR = r"\d[\d.]*(?:,\d+)?"


def _cauta(model, text, flags=re.I | re.S):
    m = re.search(model, text, flags)
    return m.group(1).strip() if m else None


def _citeste_conditii_pl(plat, r):
    """Condiții de racordare din Polonia (PGE, Enea, Tauron, Energa, Stoen): originalul polonez sau traducerea."""
    mw = lambda model: (lambda v: _nr_ro(v) * 1000 if v else None)(_cauta(model + r"\s*(\d+(?:,\d+)?)\s*MW", plat))
    for nume in ("PGE Dystrybucja", "Enea Operator", "Tauron Dystrybucja", "Energa-Operator", "Stoen Operator"):
        if re.search(nume.replace("-", r"[\s-]+").replace(" ", r"\s+"), plat, re.I):
            r["operator"] = nume
            break
    r["nr"] = _cauta(r"(?:Condi[țţt]ii\s+de\s+racordare|Warunki\s+przy[łl][ąa]czenia)\s+nr\.?\s*([\w/.-]+\w)", plat)
    m = re.search(r"[A-Za-zăâîșțłśźżó],\s*(\d{2})[-.](\d{2})[-.](\d{4})\b", plat[:1500]) or \
        re.search(r"\b(\d{2})[-.](\d{2})[-.](\d{4})\b", plat[:1500])            # 'Varșovia, 13-03-2025'

    if m:
        r["data"] = f"{m[1]}.{m[2]}.{m[3]}"
    firme = re.findall(r"([A-ZĄĆĘŁŃÓŚŹŻĂÂÎȘȚ][\w .&'-]{2,60}?\s+Sp\.\s*z\s*o\.\s*o\.)", plat)
    r["beneficiar"] = next((f.strip() for f in firme if not re.search(r"Dystrybucja|Operator", f)), None)
    r["obiectiv"] = _cauta(r"(?:Denumirea\s+obiectivului[^:]{0,30}|Nazwa\s+obiektu[^:]{0,30}):\s*(.{3,120}?)\.\s", plat)
    r["amplasament"] = _cauta(r"(?:Amplasament|Lokalizacja|Miejsce\s+lokalizacji):\s*(.{3,150}?),\s*(?:ul\.|str\.|parcel|dz)",
                              plat)
    kv = _cauta(r"(?:tensiunea\s+nominal[ăa]\s+de|napi[ęe]ciu\s+znamionowym)\s+(\d+(?:,\d+)?)\s*kV", plat)
    if kv:
        r["poc_kv"] = _nr_ro(kv)
    r["evac_kw"] = (mw(r"Putere\s+aprobat[ăa]\s+evacuat[ăa][^–-]{0,20}[–-]")
                    or mw(r"[Mm]oc\s+przy[łl][ąa]czeniowa[^0-9]{0,60}?wprowadz\w*[^0-9]{0,30}")
                    or mw(r"Putere\s+maxim[ăa]:") or mw(r"Moc\s+maksymalna:"))
    r["abs_kw"] = (mw(r"Putere\s+aprobat[ăa]\s+absorbit[ăa]\s*[–-]")
                   or mw(r"[Mm]oc\s+przy[łl][ąa]czeniowa[^0-9]{0,60}?pob\w*[^0-9]{0,30}"))
    pi = _cauta(r"putere\s+instalat[ăa]\s+de\s+(\d+(?:,\d+)?)\s*MW\s+pentru\s+sistemul\s+de\s+stocare", plat) or \
        _cauta(r"moc\w*\s+zainstalowan\w*\s+(?:magazynu\s+energii\s+)?(\d+(?:,\d+)?)\s*MW", plat)
    if pi:
        r["pi_is_kw"] = _nr_ro(pi) * 1000
    cap = _cauta(r"(?:pojemno[śs][ćc]\w*|capacitat\w*)[^0-9]{0,40}(\d+(?:,\d+)?)\s*MWh", plat)
    if cap:
        r["mwh"] = _nr_ro(cap)
    for k in [k for k, v in r.items() if v is None]:
        del r[k]


def _citeste_caiet_sarcini(plat, r):
    """Caiet de sarcini / specificație tehnică (licitații): obiectiv, beneficiar, amplasament, puterea instalată
    minimă (AC), capacitatea, evacuarea aprobată (din ATR-ul citat), trafo minim, invertoare string."""
    r["obiectiv"] = _cauta(r"Denumirea\s+obiectivului\s+de\s+investi\w+\s+(.{5,200}?)\s+Beneficiar", plat)
    r["beneficiar"] = _cauta(r"Beneficiar\s*/\s*Achizitor.{0,80}?\)\s*(.{3,80}?S\.?R\.?L\.?|.{3,80}?S\.?A\.)", plat)
    r["amplasament"] = _cauta(r"Amplasament\s+(.{3,150}?jude[țţt]ul\s+[\wăâîșțşţĂÂÎȘȚ-]+)", plat)
    kw = _cauta(r"putere\s+total\w*\s+instalat\w*\s+[îi]n\s+curent\s+alternativ\s+de\s+minimum\s+(" + _NR + r")\s*kW",
                plat, re.I)
    if kw:
        r["mw_min"] = _nr_ro(kw) / 1000
    else:
        mw = _cauta(r"putere\s+nominal\w*\s+de\s+(\d+,\d+|\d+)\s*MW", plat, re.I)
        if mw:
            r["mw_min"] = _nr_ro(mw)
    m = re.search(r"puterea\s+maxim\w*\s+aprobat\w*\s+pentru\s+evacuare(\s*/\s*absorb\w*)?.{0,40}?:\s*(\d+(?:,\d+)?)\s*MW",
                  plat, re.I)
    if m:
        r["evac_kw"] = _nr_ro(m.group(2)) * 1000
        if m.group(1):
            r["abs_kw"] = r["evac_kw"]
    t = _cauta(r"putere\s+total\w*\s+instalat\w*\s+a\s+transformatoarelor\s+de\s+minimum\s+(" + _NR + r")\s*kVA", plat)
    if t:
        r["trafo_min_kva"] = _nr_ro(t)
    r["cere_string"] = bool(re.search(r"arhitectur\w*\s+obligatori\w*\s+de\s+tip\s+.{0,3}string|"
                                      r"mandatory\s+.{0,3}string", plat, re.I))


def citeste_atr_text(text):
    """Datele unui ATR (model ANRE) sau ale unui studiu de soluție. Returnează un dicționar cu valorile găsite,
    'avertismente' (ce trebuie verificat de mână) și propunerile pentru ofertă (mw, mw_max, mwh, poc_kv)."""
    r = {"avertismente": [], "trafo_kva": []}
    av = r["avertismente"]
    # antetele / subsolurile repetate pe fiecare pagină (adresa operatorului, 'Pagina 4 / 13') rup tabelele: le scoatem
    original = re.sub(r"\s+", " ", text)
    linii = text.splitlines()
    frecventa = {}
    for l in linii:
        frecventa[l.strip()] = frecventa.get(l.strip(), 0) + 1
    # rânduri identice dar reale (trafo-uri, etichetele tabelului de putere) nu se scot
    tehnic = re.compile(r"kVA|kW|MW|kV\b|trafo|transformator|celul|puter|evacu|absorb|simultan|poate|stocare", re.I)
    lungi = [n for l, n in frecventa.items() if len(l) > 20]
    prag = max(3, 0.6 * max(lungi, default=0))       # antetul apare (aproape) pe fiecare pagină
    text = "\n".join(l for l in linii if not (frecventa[l.strip()] >= prag and len(l.strip()) > 3
                                                and not tehnic.search(l))
                     and not re.fullmatch(r"\s*(Pagina\s*)?\d+\s*/\s*\d+\s*", l))
    if len(text.split()) < 50:
        r["scanat"] = True
        av.append("Documentul pare scanat (fără text): completează valorile manual.")
        return r
    plat = re.sub(r"\s+", " ", text)

    r["tip"] = ("Condiții de racordare (Polonia)" if re.search(
                    r"Condi[țţt]ii\s+de\s+racordare\s+nr|Warunki\s+przy[łl][ąa]cz|PGE\s+Dystrybucja|Enea\s+Operator|"
                    r"Tauron\s+Dystrybucja|Energa[\s-]+Operator|Stoen\s+Operator", plat[:4000], re.I)
                else "Cerințe asigurator / siguranță" if re.search(r"INFORMACJA\s+TECHNICZNA|Magazyny\s+energii\s+zewn",
                                                                    plat[:2000], re.I)
                else "Aviz de amplasament" if re.search(r"AVIZ\s+DE\s+AMPLASAMENT", plat[:3000], re.I)
                else "Aviz tehnic de racordare" if re.search(r"AVIZ\s+TEHNIC\s+DE\s+RACORDARE", text, re.I)
                else "Studiu de soluție" if re.search(r"studiu\w*\s+de\s+solu", plat, re.I)
                else "Caiet de sarcini" if re.search(r"SPECIFICA[ȚŢT]IA\s+TEHNIC|Technical\s+Specification|"
                                                     r"caiet\w*\s+de\s+sarcini", plat, re.I)
                else "Schemă de stație" if re.search(r"Schema\s+normal", plat, re.I) else "document")
    if r["tip"] == "Cerințe asigurator / siguranță":
        av.append("E un document de cerințe de siguranță (asigurator), nu un ATR: nu conține puteri aprobate.")
        for model, nota in ((r"(\d+)\s*MWh\b.{0,80}?|zesp\w+\s+magazyn\w+\s+o\s+mocy\s+nie\s+przekraczaj\w+\s+(\d+)\s*MWh",
                             "grupuri de baterii de maximum {} MWh"),
                            (r"co\s+najmniej\s+(\d+)\s*m\s+od", "distanță minimă {} m față de alte obiecte"),
                            (r"(REI\s*\d+)", "sau perete antifoc {}")):
            m = re.search(model, plat, re.I)
            if m and any(m.groups()):
                av.append("Cerință: " + nota.format(next(g for g in m.groups() if g)) + ".")
        return r
    if r["tip"] == "Condiții de racordare (Polonia)":
        _citeste_conditii_pl(plat, r)
    if r["tip"] == "Aviz de amplasament":
        av.append("E un aviz de amplasament, nu un ATR: nu conține puterea aprobată. Puterea și capacitatea "
                  "se iau din cererea clientului (ATR-ul se cere ulterior).")
        r["obiectiv"] = _cauta(r"pentru\s+obiectivul\s+[\d/ ]*(CONSTRUIRE.{3,150}?)\s+cu\s+destina", plat)
        return r
    if r["tip"] in ("Schemă de stație", "document"):
        av.append("Documentul nu pare a fi un ATR sau un studiu de soluție"
                  + (" (e schema normală a unei stații)" if r["tip"] == "Schemă de stație" else "")
                  + ": valorile de mai jos pot lipsi sau nu se aplică.")
    if r["tip"] == "Caiet de sarcini":
        _citeste_caiet_sarcini(plat, r)
    gasiti = []                                   # operatorul = cel pomenit primul (antetul documentului)
    for nume, cuvant in (("Distribuție Energie Electrică Romania (Electrica)", r"Distribu\w+ Energie Electric"),
                         ("E-Distribuție", r"e-distribu"), ("Delgaz Grid", r"Delgaz"),
                         ("Rețele Electrice", r"Re[țt]ele Electrice\s+Rom|RETELE ELECTRICE|Re[țt]ele Electrice (?:Muntenia|"
                                              r"Dobrogea|Banat|Rom[aâ]nia)"),
                         ("Distribuție Energie Oltenia", r"Distribu\w+\s+(?:Energie\s+)?Oltenia"),
                         ("Transelectrica", r"Transelectrica")):
        m = re.search(cuvant, original, re.I if nume != "Rețele Electrice" else 0)
        if m:
            gasiti.append((m.start(), nume))
    if gasiti:
        r["operator"] = min(gasiti)[1]
    data_re = r"(\d{2}[./]\d{2}[./]\d{4})"
    antet = [m for m in re.finditer(r"\bATR\s+nr\.?\s*(\d{4,})\s*/\s*" + data_re, original, re.I)
             if len(re.findall(r"\bATR\s+nr\.?\s*" + m.group(1), original, re.I)) >= 2]   # antet pe fiecare pagină
    m = (antet[0] if antet else None) or re.search(r"nr\.?\s*(\d{4,})\s*/\s*(?:data\s*)?" + data_re, plat, re.I) or \
        re.search(r"racordare\s+nr\.?\s*(\d{4,})\s+din\s+(?:data\s+)?" + data_re, plat, re.I) or \
        re.search(r"racordare\s+nr\s+din\s+(?:data\s+)?(\d{4,})\s+" + data_re, plat, re.I)  # câmpuri mutate la final
    if r["tip"] == "Caiet de sarcini":
        m = re.search(r"conform\s+ATR\s+nr\.?\s*(\d+)\s*/\s*(\d{2}\.\d{2}\.\d{4})", plat, re.I)
    elif r["tip"] != "Aviz tehnic de racordare":     # studiu: ATR-ul la care se referă
        m = re.search(r"avizul\s+tehnic\s+de\s+racordare\s+nr\.?\s*(\d+)\s+din\s+(\d{2}\.\d{2}\.\d{4})", plat, re.I)
    if m:
        r["nr"], r["data"] = m.group(1), m.group(2).replace("/", ".")
    r["consumator"] = bool(re.search(r"PENTRU\s+CONSUMATOR", plat, re.I))   # rafinat după puterea aprobată
    r["beneficiar"] = r.get("beneficiar") or _cauta(r"adresat[ăa] de\s+(.+?),\s+pentru", plat) or \
        _cauta(r"utilizatorului\s+(.{3,80}?),\s+cu\s+(?:domiciliul|sediul)", plat) or \
        _cauta(r"Beneficiar\s*:?\s*([A-ZĂÂÎȘȚŞŢ0-9 .&-]+?S\.?R\.?L\.?)", plat)
    if not r["beneficiar"]:      # Rețele Electrice (consumator): valorile apar după etichete, la capăt de rând
        zona = _cauta(r"utilizatorului(.{0,250})", plat) or ""
        r["beneficiar"] = _cauta(r"((?:S\.?C\.?\s+)?[A-ZĂÂÎȘȚŞŢ][A-ZĂÂÎȘȚŞŢ0-9 .&-]{2,60}?\s(?:S\.?R\.?L\.?|S\.?A\.?))(?![A-Z])",
                                 zona, re.S)
    if r["beneficiar"]:
        r["beneficiar"] = re.sub(r"\s*/\s*-+\s*$", "", r["beneficiar"]).strip()
    r["obiectiv"] = r.get("obiectiv") or \
        _cauta(r"adresat[ăa] de .{0,120}?,\s+pentru\s+(.{0,200}?)\s+ce\s+a?par[țţt]ine", plat) or \
        _cauta(r"se\s+aprob\w*\s+racordarea.{0,90}?(?:consum\s+[sș]i\s+de\s+producere|producere|consum)\s+(?:permanent/temporar\s+)?"
               r"([^()]{2,80}?)\s*\(denumirea\)", plat) or \
        _cauta(r"\(denumirea[^)]*\)\s*([A-ZĂÂÎȘȚ0-9][A-ZĂÂÎȘȚ0-9 .,+-]{2,60}?)\s*(?:\n|amplasat)", text, re.S)
    r["amplasament"] = r.get("amplasament") or _cauta(r"amplasat\(?[ăa]?\)?\s+[îi]n\s+(jude[țţt]ul\s+[^,]+(?:,\s*-?\s*[^,]+)?)", plat)
    if r["amplasament"]:                              # 'municipiul/ orasul/ comuna/ sat/ sector GURBANESTI' -> 'GURBANESTI'
        r["amplasament"] = re.sub(r"(?:municipiul|ora[sș]ul|comuna|sat|sector)(?:\s*/\s*(?:municipiul|ora[sș]ul|comuna|"
                                  r"sat|sector))+\s*", "", r["amplasament"])
    if not r["amplasament"]:                          # câmpuri mutate la capăt de rând (Rețele Electrice, consumator)
        loc = _cauta(r"amplasat\s+[îi]n\s+jude[țţt]ul.{0,120}?cod\s+postal\s+([A-ZĂÂÎȘȚ][A-ZĂÂÎȘȚ ]{2,60}?)\s*$", text,
                     re.I | re.S | re.M)
        r["amplasament"] = f"județul {loc}" if loc else None

    # Punctul de racordare (tensiunea)
    kv = _cauta(r"punct\w*\s+de\s+racord\w*\s*:?\s*(?:este\s+stabilit|se\s+stabil\w*|se\s+realizeaz\w*)\s+la\s+"
                r"(?:nivelul\s+de\s+tensiune\s+(?:de\s+)?|bara\s+de\s+|tensiunea\s+de\s+)?(\d+(?:[,.]\d+)?)\s*kV", plat)
    kv = kv or _cauta(r"punct\w*\s+de\s+racord\w*\s+este\s+stabilit\s+la\s+nivelul\s+de\s+tensiune\s+kV,\s*la\s*,?\s*"
                      r"(\d+(?:[,.]\d+)?)\s", plat)          # valoarea mutată după 'kV, la ,'
    v = _cauta(r"punct\w*\s+de\s+racord\w*\s+este\s+stabilit\s+la\s+nivelul\s+de\s+tensiune\s*:?\s*(\d[\d/ ]*?)\s*V\b",
               plat, re.I)                            # Oltenia: '110000/0/0 V', '0/20000/0 V'
    if v and not kv:
        volti = max(int(x) for x in re.findall(r"\d+", v))
        if volti >= 1000:
            kv = f"{volti / 1000:g}"
    kv = kv or _cauta(r"punct\w*\s+de\s+(?:conexiun\w*|racord\w*)\s+(?:existent\s+)?de\s+(\d+(?:[,.]\d+)?)\s*kV", plat)
    if kv:
        r["poc_kv"] = _nr_ro(kv)

    # Puterea aprobată (secțiunea 2, etapa finală = ultima valoare de pe rând). Formatele diferă de la un operator la
    # altul: eticheta e înainte sau după valori, kW sau kVA primul, '-' pentru etapele goale -> fiecare rând de valori
    # ia eticheta cea mai apropiată.
    sec = re.search(r"2\.\s*Puterea\s+aprobat\w*(.+?)3\.\s*Descrierea", text, re.I | re.S) or \
        re.search(r"Puterea\s+aprobat\w*\s*:\s*1\.(.+?)2\.\s*Descrierea", text, re.I | re.S) or \
        re.search(r"2\.\s*Puterea\s+aprobat\w*(.{0,3000})", text, re.I | re.S)
    if sec:
        corp = re.sub(r"\)\s*_+(?=\s)", ")", sec.group(1))        # Transelectrica: '(kW) _     139.565,73'
        model_rand = r"(?:\((kW|kV ?A)\)|(?<![\w(])(kW|kVA)(?=\s+[\d-]))((?:\s+(?:" + _NR + r"|-))+)"
        randuri = list(re.finditer(model_rand, corp))
        # etichetele se caută pe textul fără rândurile de valori (uneori un rând taie eticheta în două)
        fara_valori = re.sub(model_rand, lambda m: " " * len(m.group(0)), corp)
        etichete = [(m.start(), m.end(), re.sub(r"\s+", " ", m.group(0)).lower()) for m in re.finditer(
            r"Puterea\s+maxim\w*\s+(?:simultan\w*\s+)?ce\s+poate\s+fi\s+(?:evacuat\w*(?:\s+f[aă]r[aă]\s+realiz\w*|"
            r"\s+[îi]n\s+situa\w*)?|absorbit\w*)", fara_valori, re.I)]
        # Un rând (kVA) ține de eticheta dinaintea lui. Un rând (kW) lipit de eticheta următoare ('4100,000Puterea...',
        # cum iese din unele PDF-uri Electrica) ține de aceea; altfel tot de eticheta dinaintea lui.
        for m in randuri:
            unitate = (m.group(1) or m.group(2)).replace(" ", "")
            valori = re.findall(_NR, m.group(3))
            if not valori or not etichete:
                continue
            urmatoare = [e for e in etichete if e[0] == m.end()]
            inainte = [e for e in etichete if e[0] <= m.start()]
            if unitate == "kW" and urmatoare:
                eticheta = urmatoare[0][2]
            elif inainte:
                eticheta = inainte[-1][2]
            else:
                eticheta = min(etichete, key=lambda e: abs(e[0] - m.end()))[2]
            valoare = _nr_ro(valori[-1])
            if "absorbit" in eticheta:
                tip = "abs"
            elif "fără realiz" in eticheta or "fara realiz" in eticheta:
                tip = "evac_fara_intarire"
            elif "situa" in eticheta:
                tip = "evac_limitare"
            else:
                tip = "evac"
            r.setdefault(f"{tip}_{unitate.lower()}", valoare)
    if "evac_kw" not in r:          # Oltenia: uneori lipsește rândul principal; rămân 'fără întărire' / 'limitare'
        for tip in ("evac_fara_intarire", "evac_limitare"):
            if r.get(f"{tip}_kw"):
                r["evac_kw"], r["evac_kva"] = r[f"{tip}_kw"], r.get(f"{tip}_kva")
                av.append("În tabelul puterii aprobate lipsește rândul „evacuată”: am luat valoarea de la „"
                          + ("fără realizarea lucrărilor de întărire" if tip == "evac_fara_intarire"
                             else "situații de limitare operațională") + "”. Verifică.")
                break
    # Studiu de soluție: tabelul normei tehnice (MW)
    for cheie, eticheta in (("pi_is_kw", "Puterea activă instalată totală"),
                            ("evac_kw", "Puterea activă maximă evacuată"),
                            ("abs_is_kw", "Puterea activă maximă absorbită")):
        v = _cauta(eticheta + r"\s+MW\s+[SDR, ]+\s*(" + _NR + ")", plat, re.I)
        if v and cheie not in r:
            r[cheie] = _nr_ro(v) * 1000
    m = re.search(r"poate\s+evacua(\s+(?:și|si)\s+absorbi)?\s+puterea\s+maxim\w*\s+de\s+(" + _NR + r")\s*(MW|kW)",
                  plat, re.I)
    if m and "evac_kw" not in r:                      # aviz CTE: '... poate evacua și absorbi puterea maximă de 2.5 MW'
        val = float(m.group(2).replace(",", ".")) if m.group(3).upper() == "MW" else _nr_ro(m.group(2))
        r["evac_kw"] = val * 1000 if m.group(3).upper() == "MW" else val
        if m.group(1):
            r.setdefault("abs_kw", r["evac_kw"])
    v = _cauta(r"Capacitatea maximă totală stocată de IS\s+MWh\s+[SDR, ]+\s*(" + _NR + ")", plat, re.I)
    if v:
        r["mwh"] = _nr_ro(v)

    # Instalația de stocare (tabelul 1: Pi IS, Pmax evac IS, Pmax abs IS, capacitate)
    tab = re.search(r"Tip\s+IS\*?\s+Pi\s*IS(.{0,900})", text, re.I | re.S)
    if tab:
        # șiruri de numere consecutive (despărțite doar de spații); fără numerotarea coloanelor (1 2 3 ...) și fără
        # numărul rândului din față: primul șir rămas cu >= 4 valori = Pi IS, Pmax evac, Pmax abs, capacitate (Ah)[, MWh]
        fereastra = tab.group(1)
        numere = list(re.finditer(_NR, fereastra))
        siruri, curent = [], []
        for m in numere:
            if curent and fereastra[curent[-1].end():m.start()].strip():
                siruri.append(curent)
                curent = []
            curent.append(m)
        siruri.append(curent)
        for sir in siruri:
            valori = [_nr_ro(m.group()) for m in sir]
            while len(valori) > 4 and valori[0] == int(valori[0]) and valori[0] < 20:
                valori.pop(0)
            if len(valori) > 4 and 0 < valori[0] < valori[1] / 100:   # rest din numele tipului ('ProeM -372')
                valori.pop(0)
            if len(valori) < 4 or all(v == int(v) and v < 20 for v in valori):
                continue
            r["pi_is_kw"], r["evac_is_kw"], r["abs_is_kw"], r["capacitate_ah"] = valori[:4]
            if 0 < max(valori[:3]) < 50:              # puterile scrise în MW sub antetul (kW): '3,290' = 3290 kW
                for k in ("pi_is_kw", "evac_is_kw", "abs_is_kw"):
                    r[k] = round(r[k] * 1000, 3)
                av.append("Tabelul stocării are puterile scrise în MW (sub antetul kW): le-am înmulțit cu 1000.")
            if len(valori) > 4 and valori[4] < 1000 and re.search(r"MWh", fereastra[:400]):
                r["mwh"] = valori[4]
            break
    pi_text = _cauta(r"stocare\s+(?:nou[ăa]?\s+)?(?:cu\s+)?Pi\s*=\s*(" + _NR + r"\s*(?:MW|kW))", plat)
    if pi_text:
        val = _nr_ro(re.sub(r"\s*(MW|kW)", "", pi_text, flags=re.I))
        r["pi_text_kw"] = val * 1000 if re.search("MW", pi_text, re.I) else val
    if "mwh" not in r and tab:                       # observațiile tabelului: ex. '8000 kWh fără încărcare din RED'
        v = _cauta(r"(" + _NR + r")\s*kWh", tab.group(1)[:700])
        if v:
            r["mwh"] = _nr_ro(v) / 1000
    if "mwh" not in r:                               # ex. '- 1 MW / 4 MWh racordare în TDRI trafo 2' (de 2 ori -> 8 MWh)
        bucati = re.findall(r"(" + _NR + r")\s*MW\s*/\s*(" + _NR + r")\s*MWh\s+racord", plat, re.I)
        if len(bucati) > 1:
            r["mwh"] = sum(_nr_ro(b[1]) for b in bucati)
    if "mwh" not in r:                               # ex. 'Capacitate de stocare de 10.000,00 kWh'
        m = re.search(r"Capacitat\w*\s+(?:de\s+)?stocare\s+de\s+(" + _NR + r")\s*(kWh|MWh)", plat, re.I)
        if m:
            r["mwh"] = _nr_ro(m.group(1)) / (1000 if m.group(2).lower() == "kwh" else 1)
    if "mwh" not in r:                               # nu din coduri de produs ('LUNA2000-4,5MWH')
        v = _cauta(r"(?<![-\w.,])(" + _NR + r")\s*MWh", plat)
        if v:
            r["mwh"] = _nr_ro(v)

    # Cuplare, absorbție, trafo-uri din instalație
    r["cuplare_dc"] = bool(re.search(r"DC\s*/\s*DC|cuplat\w*\s+pe\s+(?:partea\s+de\s+)?(?:curent\s+continuu|DC)|"
                                     r"cuplare\w*\s+(?:pe\s+)?(?:partea\s+de\s+)?(?:curent\s+continuu|DC)", plat, re.I))
    r["fara_absorbtie"] = bool(re.search(r"f[aă]r[aă]\s+absorb|nu\s+va\s+absorbi|nu\s+se\s+va\s+[îi]nc[aă]rca\s+din\s+RED|"
                                         r"[îi]nc[aă]rcare\s+exclusiv\s+din", plat, re.I)) or r.get("abs_is_kw") == 0
    for m in re.finditer(r"(?:trafo|transformator\w*)(?:\s+de\s+putere)?\s+(" + _NR + r")\s*/\s*(" + _NR +
                         r")\s*kV\s*[-–,]\s*(" + _NR + r")\s*(kVA|MVA)", plat, re.I):
        context = plat[max(0, m.start() - 40):m.start()].lower()
        if "servicii" in context or "auxiliar" in plat[m.start():m.start() + 40].lower():
            continue
        kva = _nr_ro(m.group(3)) * (1000 if m.group(4).upper() == "MVA" else 1)
        r["trafo_kva"].append((kva, f"{m.group(1)}/{m.group(2)} kV"))
    # stocare legată pe joasă tensiune (TDRI) prin trafo-urile existente: puterea pe fiecare trafo e limitată de el
    r["cuplare_jt"] = bool(re.search(r"racordat\w*\s+la\s+nivelul\s+de\s+curent\s+alternativ|racordare\s+[îi]n\s+TDRI",
                                     plat, re.I))
    for m in re.finditer(r"racordare\s+[îi]n\s+TDRI\s+trafo\s*\d*\s+(" + _NR + r")\s*kVA", plat, re.I):
        r["trafo_kva"].append((_nr_ro(m.group(1)), "JT, stocare legată în TDRI"))
    for m in re.finditer(r"transformator\w*\s+(" + _NR + r")\s*MVA,?\s*(\d+)\s*/\s*(\d+)\s*kV", plat, re.I):
        r["trafo_kva"].append((_nr_ro(m.group(1)) * 1000, f"{m.group(2)}/{m.group(3)} kV"))

    r["fara_stocare"] = r["tip"] in ("Aviz tehnic de racordare", "Studiu de soluție") and not r.get("pi_is_kw") \
        and not r.get("pi_text_kw") and not re.search(r"instala\w*\s+de\s+stocare\s+(?:nou|cu\s+Pi|de\s+\d)|"
                                                       r"\bIS\s+(?:Pi\s*=|\d)", plat, re.I)
    r["consumator"] = r["consumator"] and not r.get("evac_kw")
    if r["fara_stocare"] and r["consumator"]:
        r["fara_stocare"] = False
        av.append("ATR de consumator (fără evacuare în rețea): BESS doar pentru consum propriu / peak shaving"
                  + (f", absorbție aprobată {nr_eu(r['abs_kw'])} kW" if r.get("abs_kw") else "")
                  + ". Puterea și capacitatea se iau din cererea clientului; pentru evacuare e nevoie de ATR nou.")
    if r["fara_stocare"]:
        av.append("ATR fără instalație de stocare (doar producere/consum): pentru BESS e nevoie de actualizarea "
                  "ATR-ului. Puterea aprobată de mai jos e cea a parcului.")

    # Propunerile pentru ofertă
    evac = r.get("evac_is_kw") or r.get("evac_kw")
    if evac:
        r["mw"] = evac / 1000
    pi = r.get("pi_is_kw")
    if pi and evac and pi > 2.5 * evac:
        el = re.search(r"Tabelul\s*2.{0,600}?1\s+2\s+3\s+4\s+5\s+6\s+7\s+8\s+9\s+1\s+(\d+)\s+(" + _NR + ")", plat, re.I)
        detalii = (f" (în tabelul 2: {el.group(1)} elemente x {el.group(2)} kW)" if el else "")
        av.append(f"Pi stocare din tabel ({nr_eu(pi)} kW) e de peste 2,5 ori evacuarea{detalii}: pare trecută "
                  "capacitatea (kWh) în coloana de kW. Nu am folosit-o ca limită.")
        pi = None
    if pi and r.get("pi_text_kw") and abs(pi - r["pi_text_kw"]) > 1:
        av.append(f"Pi stocare diferă: {nr_eu(pi)} kW în tabel, {nr_eu(r['pi_text_kw'])} kW în text. "
                  "Am pus ca limită valoarea din tabel; verifică.")
    pi = pi or r.get("pi_text_kw")
    if pi:
        r["mw_max"] = pi / 1000
    if r.get("evac_kw") and r.get("evac_is_kw") and abs(r["evac_kw"] - r["evac_is_kw"]) > 1:
        av.append(f"Evacuarea aprobată ({nr_eu(r['evac_kw'])} kW) diferă de Pmax evac IS ({nr_eu(r['evac_is_kw'])} kW).")
    trafo_mt = sum(k for k, rap in r["trafo_kva"] if re.match(r"(20|22|33|35)\D|0[,.]", rap.split("/")[0] + " ")
                   or "/0" in rap)
    mici = [k for k, _ in r["trafo_kva"] if k < 500]
    if mici and any(k >= 1000 for k, _ in r["trafo_kva"]):
        av.append(f"Trafo de {nr_eu(mici[0])} kVA lângă trafo-uri mari: verifică (poate fi o greșeală, ex. 2500).")
    if r["cuplare_dc"]:
        av.append("ATR-ul prevede cuplarea stocării pe DC (DC/DC, prin invertoarele existente). Un MV skid AC "
                  "înseamnă altă soluție decât cea aprobată (actualizare ATR?).")
    if trafo_mt and pi and pi / 1000 > trafo_mt / 1000 * 1.0 + 1e-6 and r["cuplare_dc"]:
        av.append(f"Pi stocare ({nr_eu(pi)} kW) > trafo-urile din parc ({nr_eu(trafo_mt)} kVA), prin care trece "
                  "stocarea cuplată pe DC.")
    if r.get("cuplare_jt"):
        trafo_jt = [k for k, rap in r["trafo_kva"] if "TDRI" in rap]
        av.append("Stocarea se leagă pe joasă tensiune (TDRI), prin trafo-urile existente"
                  + (f" ({' + '.join(nr_eu(k) for k in trafo_jt)} kVA)" if trafo_jt else "")
                  + ": puterea legată pe fiecare trafo nu poate depăși trafo-ul (limitare de transformator).")
    if r["fara_absorbtie"]:
        av.append("Fără absorbție din rețea: bateriile se încarcă doar din parc (de cerut prin EMS).")
    if r.get("poc_kv") and r["poc_kv"] > 35:
        sec_kv = [rap for k, rap in r["trafo_kva"] if rap.startswith(str(int(r["poc_kv"])))]
        av.append(f"POC la {r['poc_kv']:g} kV: skidurile se leagă pe tensiunea medie a stației clientului"
                  + (f" ({sec_kv[0]})" if sec_kv else "") + ". Alege POC-ul skidului în ofertă.")
    # tabelul 2: nr. de elemente de stocare și Pi / element (rândul de după numerotarea coloanelor 1..9)
    m = re.search(r"Tabelul\s*2.{0,600}?1\s+2\s+3\s+4\s+5\s+6\s+7\s+8\s+9\s+1\s+(\d+)\s+(" + _NR + ")", plat, re.I)
    if m and int(m.group(1)) > 0:
        r["nr_elemente"], r["pi_element_kw"] = int(m.group(1)), _nr_ro(m.group(2))
        cod = re.search(r"Tip\s+IS.{0,300}?(\d+(?:[,.]\d+)?)\s*MWH", plat, re.I)    # ex. LUNA2000-4,5MWH-2H1
        if cod and "mwh_baterii" not in r:
            r["mwh_baterii"] = round(r["nr_elemente"] * _nr_ro(cod.group(1)), 3)
            if not r.get("mwh"):
                r["mwh"] = r["mwh_baterii"]
                av.append(f"Capacitatea calculată din tabel: {r['nr_elemente']} elemente x "
                          f"{nr_eu(_nr_ro(cod.group(1)))} MWh = {nr_eu(r['mwh'])} MWh. Verifică.")
    # bateriile din tabel: '3 x "Suntera JKE-5015K-2H" cu doar 10 rack-uri ... --> JKE4180K' -> 3 x 4,18 MWh
    m = re.search(r"(\d+)\s*x\s*\"?\s*(?:Suntera\s+)?JKE-?(\d{4})K", plat, re.I)
    if m:
        dupa = re.search(r"-->\s*JKE-?(\d{4})K", plat[m.end():m.end() + 200], re.I)
        r["mwh_baterii"] = round(int(m.group(1)) * int((dupa or m).group(1 if dupa else 2)) / 1000, 3)
        if r.get("mwh") and abs(r["mwh_baterii"] - r["mwh"]) > 0.1:
            av.append(f"Capacitatea diferă: {nr_eu(r['mwh'])} MWh în text, dar bateriile din tabel dau "
                      f"{m.group(1)} x {(dupa or m).group(1 if dupa else 2)}K = {nr_eu(r['mwh_baterii'])} MWh "
                      "(rack-uri scoase?). Verifică.")
    mw_ref = (r.get("pi_is_kw") or r.get("evac_is_kw") or r.get("evac_kw") or 0) / 1000
    if r.get("mwh_baterii") and mw_ref and r["mwh_baterii"] > mw_ref * ORE_MAX * 1.5:
        if r.get("mwh") == r["mwh_baterii"]:
            del r["mwh"]
        del r["mwh_baterii"]                         # nr. de elemente = module / celule, nu containere
    if r.get("capacitate_ah"):
        tip_is = _cauta(r"Tip\s+IS(.{0,300})", plat) or ""
        tensiune = next((v for k, v in TENSIUNE_BATERIE_V.items() if k.lower() in tip_is.lower()), None)
        if tensiune:
            r["mwh_din_ah"] = round(r["capacitate_ah"] * tensiune / 1e6, 3)
            av.append(f"Capacitate din Ah: {nr_eu(r['capacitate_ah'])} Ah x {nr_eu(tensiune)} V = "
                      f"{nr_eu(r['mwh_din_ah'])} MWh" + (f" (în text: {nr_eu(r['mwh'])} MWh)" if r.get("mwh") else "")
                      + ".")
            r["mwh"] = r["mwh_din_ah"]
    if r.get("mwh") and mw_ref and r["mwh"] / mw_ref < ORE_MIN:
        av.append(f"Capacitatea găsită ({nr_eu(r['mwh'])} MWh) e prea mică pentru {nr_eu(mw_ref)} MW: pare a unui "
                  "singur container. Calculeaz-o din Ah (tensiunea bateriei, jos) sau ia-o din cererea clientului.")
    if "mwh" not in r:
        av.append("Capacitatea (MWh) nu apare în ATR (doar Ah): calculeaz-o din Ah cu tensiunea bateriei (jos) sau "
                  "ia-o din cererea clientului.")
    if r.get("mw_min"):
        av.append(f"Caietul de sarcini cere minimum {nr_eu(r['mw_min'])} MW instalați (AC)"
                  + (f", iar evacuarea aprobată e {nr_eu(r['evac_kw'] / 1000)} MW: skidurile trebuie să aibă "
                     f"cel puțin {nr_eu(r['mw_min'])} MW, limitate prin EMS la {nr_eu(r['evac_kw'] / 1000)} MW. "
                     "Alege Nr. skiduri / skidul astfel încât puterea să ajungă la minimul cerut."
                     if r.get("evac_kw") and r["evac_kw"] / 1000 < r["mw_min"] - 1e-6 else "."))
    if r.get("cere_string"):
        av.append("Caietul cere invertoare de tip STRING (arhitectură obligatorie): un MV skid cu PCS central "
                  "poate fi declarat neconform.")
    if r.get("trafo_min_kva"):
        av.append(f"Post de transformare cerut: minimum {nr_eu(r['trafo_min_kva'])} kVA.")
    if not evac and not r["consumator"]:
        av.append("Nu am găsit puterea aprobată pentru evacuare.")
    return r


# ---------------- citirea cu AI (Claude) ----------------
_NUL = lambda tip: {"anyOf": [{"type": tip}, {"type": "null"}]}
CAMPURI_AI = {
    "tip": ({"type": "string"}, "Aviz tehnic de racordare / Studiu de soluție / Caiet de sarcini / Aviz de amplasament / "
                                 "Cerere de ofertă / alt document"),
    "operator": (_NUL("string"), "operatorul de rețea (Electrica, Rețele Electrice, Delgaz Grid, E-Distribuție, "
                                 "Distribuție Energie Oltenia, Transelectrica)"),
    "nr": (_NUL("string"), "numărul ATR-ului / documentului"),
    "data": (_NUL("string"), "data documentului, zz.ll.aaaa"),
    "beneficiar": (_NUL("string"), "beneficiarul / clientul (firma)"),
    "obiectiv": (_NUL("string"), "denumirea obiectivului / proiectului"),
    "amplasament": (_NUL("string"), "localitatea și județul"),
    "poc_kv": (_NUL("number"), "tensiunea punctului de racordare, kV"),
    "evac_kw": (_NUL("number"), "puterea maximă aprobată pentru evacuare în rețea (etapa finală), kW"),
    "evac_kva": (_NUL("number"), "aceeași, în kVA"),
    "abs_kw": (_NUL("number"), "puterea maximă aprobată absorbită din rețea, kW"),
    "pi_is_kw": (_NUL("number"), "puterea instalată a instalației de stocare (Pi IS), kW"),
    "evac_is_kw": (_NUL("number"), "Pmax evacuată de instalația de stocare, kW"),
    "abs_is_kw": (_NUL("number"), "Pmax absorbită de instalația de stocare, kW"),
    "capacitate_ah": (_NUL("number"), "capacitatea stocării în Ah, dacă e dată în Ah"),
    "mwh": (_NUL("number"), "capacitatea totală a stocării în MWh (calculată doar dacă e sigur; altfel null)"),
    "mw_min": (_NUL("number"), "puterea minimă instalată cerută (caiete de sarcini), MW"),
    "baterii": (_NUL("string"), "tipul / modelul bateriilor și al invertoarelor, cum apar în document"),
    "cuplare_dc": ({"type": "boolean"}, "stocarea e cuplată pe DC (DC/DC, prin invertoarele parcului)"),
    "fara_absorbtie": ({"type": "boolean"}, "bateriile nu au voie să se încarce din rețea"),
    "fara_stocare": ({"type": "boolean"}, "documentul NU prevede nicio instalație de stocare"),
    "consumator": ({"type": "boolean"}, "e ATR de consumator (fără evacuare în rețea)"),
    "cere_string": ({"type": "boolean"}, "se cer obligatoriu invertoare de tip string"),
    "observatii": ({"type": "array", "items": {"type": "string"}},
                   "ce trebuie verificat de om la ofertare: neconcordanțe, limitări, condiții speciale (scurt, în română)"),
}
PROMPT_AI = """Ești inginer la un furnizor de sisteme BESS (baterii + PCS / MV skid) din România. Primești un document
(aviz tehnic de racordare ATR, studiu de soluție, caiet de sarcini, cerere de ofertă). Extrage datele pentru ofertare.
Reguli:
- Puterile în kW (dacă documentul scrie MW, înmulțește cu 1000). Atenție la virgula zecimală românească.
- Din tabelul puterii aprobate ia valoarea de la ETAPA FINALĂ.
- Nu inventa: ce nu apare clar în document = null.
- Dacă stocarea apare doar în Ah, completează capacitate_ah și pune mwh doar dacă documentul dă și tensiunea/energia.
- În observatii pune scurt (română) ce e important pentru ofertă: evacuare diferită de Pi stocare, lipsa absorbției,
  cuplare DC, legare pe joasă tensiune prin trafo-urile existente, puteri minime cerute, invertoare string etc.
Descrierea câmpurilor:
""" + "\n".join(f"- {k}: {d}" for k, (_, d) in CAMPURI_AI.items())


def citeste_atr_ai(cale_pdf, cheie_api, model=None):
    """Citește PDF-ul cu Claude (API Anthropic) și întoarce același dicționar ca citeste_atr_text."""
    import anthropic
    with open(cale_pdf, "rb") as f:
        pdf = base64.standard_b64encode(f.read()).decode("ascii")
    schema = {"type": "object", "properties": {k: s for k, (s, _) in CAMPURI_AI.items()},
              "required": list(CAMPURI_AI), "additionalProperties": False}
    client = anthropic.Anthropic(api_key=cheie_api)
    model = model or MODEL_AI
    parametri = dict(
        model=model, max_tokens=16000,
        output_config={"effort": "medium", "format": {"type": "json_schema", "schema": schema}},
        messages=[{"role": "user", "content": [
            {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": pdf}},
            {"type": "text", "text": PROMPT_AI}]}])
    if model.startswith(("claude-sonnet-5-5", "claude-opus-5", "claude-fable")):
        # dacă modelul refuză (rar), API-ul reîncearcă singur pe alt model
        raspuns = client.beta.messages.create(betas=["server-side-fallback-2026-07-01"], fallbacks="default",
                                              **parametri)
    else:
        raspuns = client.messages.create(**parametri)
    if raspuns.stop_reason == "refusal":
        raise RuntimeError("Claude a refuzat documentul. Folosește citirea normală.")
    if raspuns.stop_reason == "max_tokens":
        raise RuntimeError("Răspunsul AI a fost tăiat (document prea lung). Folosește citirea normală.")
    text = next(b.text for b in raspuns.content if b.type == "text")
    return rezultat_din_ai(json.loads(text), model)


def rezultat_din_ai(d, model=""):
    """Datele de la AI -> dicționarul cititorului (cu propunerile pentru ofertă: mw, mw_max)."""
    r = {k: v for k, v in d.items() if v not in (None, "")}
    r.setdefault("trafo_kva", [])
    av = r["avertismente"] = [f"(AI) {o}" for o in d.get("observatii") or []]
    r.pop("observatii", None)
    r["citit_cu"] = model
    evac = r.get("evac_is_kw") or r.get("evac_kw")
    if evac:
        r["mw"] = evac / 1000
    if r.get("pi_is_kw") and (not evac or r["pi_is_kw"] <= 2.5 * evac):
        r["mw_max"] = r["pi_is_kw"] / 1000
    if r.get("capacitate_ah") and not r.get("mwh"):
        tensiune = next((v for k, v in TENSIUNE_BATERIE_V.items() if k.lower() in (r.get("baterii") or "").lower()), None)
        if tensiune:
            r["mwh"] = round(r["capacitate_ah"] * tensiune / 1e6, 3)
            av.append(f"Capacitate din Ah: {nr_eu(r['capacitate_ah'])} Ah x {nr_eu(tensiune)} V = {nr_eu(r['mwh'])} MWh.")
    if r.get("mw_min") and evac and evac / 1000 < r["mw_min"] - 1e-6:
        av.append(f"Minimum {nr_eu(r['mw_min'])} MW instalați, limitați prin EMS la {nr_eu(evac / 1000)} MW.")
    if not r.get("mwh"):
        av.append("Capacitatea (MWh) nu e sigură: din Ah cu tensiunea bateriei sau din cererea clientului.")
    return r
