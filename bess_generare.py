# -*- coding: utf-8 -*-
"""Pregătirea și generarea ofertei (fereastră și motor).
Parte din generator_oferte_bess.py: se încarcă în același spațiu de nume (toate funcțiile
și setările sunt comune), deci aici NU se fac importuri proprii și nu se rulează direct."""
# pylint: disable=undefined-variable
# flake8: noqa

# ======================================================================
# PREGĂTIRE ȘI GENERARE (comune pentru fereastră și pentru modul motor)
# ======================================================================
def branduri_skid(date):
    """Brandurile din MV SKID + MV SKID for different POC, fără dubluri."""
    branduri = lista_branduri(date["skid"])
    if date.get("skid_poc") is not None:
        branduri += [b for b in lista_branduri(date["skid_poc"]) if b not in branduri]
    return branduri


def _valoare_manuala(m, cheie, eticheta, implicit=None, bani=False):
    """Număr dintr-un câmp manual; gol -> implicit (sau eroare dacă implicit e None).
    bani=True: punctul e separator de mii ('75.000' = 75000, '1.250,50' = 1250.5);
    altfel punctul și virgula sunt zecimale ('5.015' = '5,015' = 5.015 MWh)."""
    text = str(m.get(cheie, "")).strip()
    if not text:
        if implicit is None:
            raise ValueError(f"Completează câmpul '{eticheta}' în tab-ul Componente manuale.")
        return implicit
    text = text.replace(" ", "").replace("$", "")
    if ("," in text and "." in text) or (bani and re.fullmatch(r"\d{1,3}(\.\d{3})+", text)):
        text = text.replace(".", "").replace(",", ".")      # 1.234,56 / 75.000 (bani) -> 1234.56 / 75000
    try:
        return float(text.replace(",", "."))
    except ValueError:
        raise ValueError(f"Câmpul '{eticheta}' din Componente manuale nu e un număr: {m.get(cheie)}")


def _procent_manual(m):
    """Extra warranty manual, în %: '3', '3%', '2,5' -> 3 / 2.5; gol -> 0 (fără extra warranty)."""
    pct = _valoare_manuala(dict(m, ext_gar=str(m.get("ext_gar", "")).replace("%", "")),
                           "ext_gar", "Extra warranty (%)", 0.0)
    if not 0 <= pct <= 100:
        raise ValueError("Extra warranty manual trebuie să fie un procent între 0 și 100.")
    return pct


def _text_garantie(text):
    text = str(text or "").strip()
    try:
        return text_ani(float(text.replace(",", ".")))           # 5 -> '5 years'
    except ValueError:
        return text


def baterie_manuala(m, incoterm):
    """Bateria din tab-ul Componente manuale, cu aceleași chei ca alege_baterie()."""
    if not str(m.get("brand", "")).strip() or not str(m.get("model", "")).strip():
        raise ValueError("Completează brandul și modelul bateriei manuale.")
    cap = _valoare_manuala(m, "cap", "Capacitate / baterie (MWh)")
    if cap <= 0:
        raise ValueError("Capacitatea bateriei manuale trebuie să fie pozitivă.")
    pret = _valoare_manuala(m, "pret_ddp" if incoterm == "DDP" else "pret_dap", f"Preț {incoterm} ($/MWh)", bani=True)
    greutate = str(m.get("weight", "")).strip()
    try:
        greutate = f"~ {float(greutate.replace(',', '.')):g}".replace(".", ",") if greutate else ""
    except ValueError:
        pass
    return {
        "tip": m["model"].strip(), "brand": m["brand"].strip(), "cap": cap, "pret": pret, "incoterm": incoterm,
        "coloana_pret": "manual", "data": "introdus manual", "rand_excel": "-",
        "transport": 0.0 if m.get("transport_inclus") else _valoare_manuala(m, "transport", "Transport", 0.0, True),
        "garantie": _text_garantie(m.get("garantie")),
        "manual": True, "ext_gar_pct": _procent_manual(m),
        "kva_aux": _valoare_manuala(m, "kva_aux", "Consum aux / baterie", kva_aux_baterie(cap)),
        "specs": {"model": str(m.get("model_cod", "")).strip() or m["model"].strip(),
                  "config": str(m.get("config", "")).strip(), "weight": greutate},
    }


def optiune_skid_manuala(date, mw, nr_baterii, kva_baterie, forteaza_trafo, nr_fix, m):
    """Skidul din tab-ul Componente manuale, ca opțiune de skid (fără verificarea intrărilor DC)."""
    if not str(m.get("brand", "")).strip() or not str(m.get("model", "")).strip():
        raise ValueError("Completează brandul și modelul skidului manual.")
    putere = _valoare_manuala(m, "putere", "Putere / skid (MW)")
    pret = _valoare_manuala(m, "pret", "Preț / skid ($)", bani=True)
    if putere <= 0 or pret <= 0:
        raise ValueError("Puterea și prețul skidului manual trebuie să fie pozitive.")
    nr = nr_fix or math.ceil(mw / putere - 1e-9)
    o = {"brand": m["brand"].strip(), "tip": m["model"].strip(), "nr": nr, "putere_unit": putere,
         "putere_totala": nr * putere, "intrari": 0, "min_intrari": 0, "max_intrari": 0, "acceptate": "neverificat",
         "pret_unit": pret, "pret_total": nr * pret,
         "transport": 0.0 if m.get("transport_inclus") else _valoare_manuala(m, "transport", "Transport", 0.0, True),
         "com": _valoare_manuala(m, "com", "Commissioning / skid", 0.0, True), "similar": "", "conexiune": "",
         "putere_pcs": _valoare_manuala(m, "putere_pcs", "Putere PCS (MW)", putere),
         "garantie": _text_garantie(m.get("garantie")),
         "manual": True, "ext_gar_pct": _procent_manual(m)}
    trafo_pricing = variante_trafo(date["aux"]) if date.get("aux") is not None else []
    calcul_aux(o, _valoare_manuala(m, "aux_kva", "Trafo aux inclus", 0.0), nr_baterii, kva_baterie,
               forteaza_trafo, trafo_pricing)
    return o


def _ani_camp(valoare):
    """Garanția pentru câmpul manual: 15 -> 15, '10 years' -> '10 years', gol -> ''."""
    return text_ani(valoare).replace(" years", "") if not isinstance(valoare, str) else text_ani(valoare)


def baterie_din_pricing(date, baterii_map, model):
    """Câmpurile bateriei manuale completate din Pricing (ultima cotație a modelului) și Specificații."""
    df = date["bat"]
    rand = df.loc[baterii_map[model]]
    tip = curata_tip(rand.iloc[BAT_COL_TYPE])
    specs = specs_baterie(date.get("specs"), tip) or {}
    return {
        "brand": str(rand.iloc[BAT_COL_BRAND]).strip(), "model": tip, "cap": numar(rand.iloc[BAT_COL_CAP]),
        "pret_ddp": numar(rand.iloc[coloana_pret_baterie(df, "DDP")]),
        "pret_dap": numar(rand.iloc[coloana_pret_baterie(df, "DAP")]),
        "transport": numar(rand.iloc[BAT_COL_TRANSPORT]),
        "garantie": _ani_camp(rand.iloc[BAT_COL_GARANTIE]) if len(rand) > BAT_COL_GARANTIE else "",
        "ext_gar": "", "kva_aux": "",
        "model_cod": specs.get("model", ""), "config": specs.get("config", ""), "weight": specs.get("weight", ""),
    }


def skiduri_din_pricing(date):
    """{eticheta: (foaie, index)} pentru toate skidurile din 'MV SKID' și 'MV SKID for different POC'."""
    rezultat = {}
    for foaie, df in (("skid", date.get("skid")), ("skid_poc", date.get("skid_poc"))):
        if df is None:
            continue
        c_brand, c_tip = df.columns[SKID_COL_BRAND], df.columns[SKID_COL_TYPE]
        for idx, rand in df.dropna(subset=[c_tip]).iterrows():
            eticheta = f"{str(rand[c_brand]).strip()} | {str(rand[c_tip]).strip()}" + \
                       (" [alt POC]" if foaie == "skid_poc" else "")
            rezultat.setdefault(eticheta, (foaie, idx))
    return dict(sorted(rezultat.items()))


def skid_din_pricing(date, foaie, idx):
    """Câmpurile skidului manual completate din Pricing și din 'PCS - MV skid.xlsx'."""
    df = date[foaie]
    rand = df.loc[idx]
    c = df.columns
    nume, brand = str(rand[c[SKID_COL_TYPE]]).strip(), str(rand[c[SKID_COL_BRAND]]).strip()
    pcs = date["pcs"]
    dupa_nume = {cheie_skid(r[pcs.columns[0]]): r for _, r in pcs.dropna(subset=[pcs.columns[0]]).iterrows()}
    p = dupa_nume.get(cheie_skid(nume))
    if p is None:
        p = dupa_nume.get(cheie_skid(nume).split(" ")[0])
    putere_pcs = numar(p.get("PCS power")) if p is not None else 0.0
    putere = putere_pcs * (numar(p.get("no / frame")) if p is not None else 0)
    if p is not None and numar(p.get("Skid power")) > 0 and \
            any(b.lower() in brand.lower() for b in PUTERE_NOMINALA_SKID):
        putere = numar(p.get("Skid power"))
    putere = putere or numar(rand[c[SKID_COL_POWER]])
    aux = aux_inclus_skid(pd.Series({"_similar": "", COLOANA_AUX_PCS: p.get(COLOANA_AUX_PCS) if p is not None else None}),
                          nume)
    return {
        "brand": brand, "model": nume, "putere": putere, "putere_pcs": putere_pcs or putere,
        "pret": numar(rand[c[SKID_COL_PRET]]), "transport": numar(rand[c[SKID_COL_TRANSPORT]]),
        "com": numar(rand[c[SKID_COL_COM]]), "aux_kva": aux or "",
        "garantie": _ani_camp(rand[c[SKID_COL_GARANTIE]]) if len(c) > SKID_COL_GARANTIE else "", "ext_gar": "",
    }


def baterii_din_lista(baterii_map):
    """Modelele din Pricing afișate în listă (BATERII_IN_LISTA)."""
    return [b for b in baterii_map if any(t.lower() in b.lower() for t in BATERII_IN_LISTA)]


def ore_baterie(model):
    """Durata bateriei din numele modelului: 'JKE-5015K-4H-LAA' -> 4; necunoscută -> None."""
    m = re.search(r"-(\d+(?:\.\d+)?)H-", str(model), re.I)
    if m:
        return float(m.group(1))
    return next((ore for nume, ore in ORE_BATERIE_FARA_COD.items() if nume.lower() in str(model).lower()), None)


def cost_estimativ(p):
    """Costul de achiziție al variantei celei mai ieftine: baterii (+ transport) + skid (+ transport) + trafo extern."""
    b, o = p["baterie"], p["lista"][0]
    return (p["nr_baterii"] * (b["cap"] * b["pret"] + b["transport"]
                               + (commissioning_baterie(b, p["nr_baterii"]) or 0))
            + o["cost_total"] + o["nr"] * (o["transport"] + o["com"]))


def pregateste_oferta(date, baterii_map, intrari, model_baterie, incoterm, brand, nr_fix, cu_trafo,
                      manual_bat=None, manual_skid=None):
    """Alege bateria, skidurile compatibile (de la cel mai ieftin) și trafo-ul. Nu generează fișiere.
    model_baterie / brand = OPTIUNE_MANUALA_... -> componenta din manual_bat / manual_skid;
    model_baterie = OPTIUNE_AUTO_BATERIE -> bateria din listă cu oferta cea mai ieftină."""
    if model_baterie == OPTIUNE_AUTO_BATERIE:
        ore_proiect = intrari["mwh"] / intrari["mw"]
        variante, motive_auto = [], []
        for model in baterii_din_lista(baterii_map):
            ore = ore_baterie(model)
            if ore and ore > ore_proiect + 0.05:          # ex. baterie 4H nu poate livra un proiect de 2 ore
                motive_auto.append(f"{model}: baterie {ore:g}H, proiectul e de {ore_proiect:.2g} ore")
                continue
            try:
                p = pregateste_oferta(date, baterii_map, intrari, model, incoterm, brand, nr_fix, cu_trafo,
                                      manual_bat, manual_skid)
            except ValueError as e:
                motive_auto.append(f"{model}: {str(e).splitlines()[0]}")
                continue
            variante.append((cost_estimativ(p), model, p))
        if not variante:
            raise ValueError("Auto: nicio baterie din listă nu se potrivește proiectului.\n- " + "\n- ".join(motive_auto))
        variante.sort(key=lambda v: v[0])
        p = variante[0][2]
        p["auto_comparatie"] = [f"{model}: {bani_eu(cost, 0)} $ ({procent_fata_de(cost, variante[0][0])})"
                                for cost, model, _ in variante]
        p["auto_excluse"] = motive_auto
        return p
    if model_baterie == OPTIUNE_MANUALA_BATERIE:
        baterie = baterie_manuala(manual_bat or {}, incoterm)
    elif model_baterie in baterii_map:
        baterie = alege_baterie(date["bat"], baterii_map[model_baterie], incoterm)
    else:
        raise ValueError(f"Bateria '{model_baterie}' nu există în Pricing.xlsx. Selectează o baterie din listă.")
    rezerva_bat = float(intrari.get("rezerva_bat") or 0)
    if rezerva_bat:                       # rezervă de preț (ex. +10% în caz că se schimbă prețurile)
        baterie = dict(baterie, pret_baza=baterie["pret"], rezerva_pct=rezerva_bat,
                       pret=baterie["pret"] * (1 + rezerva_bat / 100))
    nr_baterii = math.ceil(intrari["mwh"] / baterie["cap"] - 1e-9)
    kva_baterie = baterie.get("kva_aux") or kva_aux_baterie(baterie["cap"])
    if brand == OPTIUNE_MANUALA_SKID:
        optiuni = [optiune_skid_manuala(date, intrari["mw"], nr_baterii, kva_baterie, cu_trafo, nr_fix,
                                        manual_skid or {})]
        motive = {"brand": 0, "putere": 0, "prea_putine": [], "prea_multe": [], "intermediare": [],
                  "exacte_respinse": [], "respinse": [], "fara_pcs": [], "incompatibil": 0}
    else:
        if abs(intrari["poc_kv"] - POC_IMPLICIT_KV) > 1e-9:
            if brand != "Oricare" and doar_poc_implicit(brand):
                raise ValueError(f"{brand} are skiduri doar pentru POC {POC_IMPLICIT_KV} kV (România), "
                                 f"nu pentru {intrari['poc_kv']:g} kV.")
        date_poc = dict(date, skid=skiduri_pentru_poc(date, intrari["poc_kv"]))
        optiuni, motive = gaseste_skiduri(date_poc, intrari["mw"], nr_baterii, brand, nr_fix,
                                          baterie_text=f"{baterie['brand']} {baterie['tip']}",
                                          kva_baterie=kva_baterie, forteaza_trafo=cu_trafo)
    if not optiuni:
        raise ValueError(mesaj_fara_skid(intrari["mw"], nr_baterii, brand, motive))
    aplica_rezerva_skid(optiuni, float(intrari.get("rezerva_skid") or 0))
    mw_max = float(intrari.get("mw_max") or 0)
    if mw_max:                            # ex. Pi-ul instalației de stocare din ATR
        peste = [o for o in optiuni if o["putere_totala"] > mw_max + 1e-6]
        optiuni = [o for o in optiuni if o["putere_totala"] <= mw_max + 1e-6]
        if not optiuni:
            raise ValueError(f"Toate skidurile compatibile depășesc puterea maximă admisă de {nr_eu(mw_max)} MW "
                             f"(ex. {peste[0]['tip']}: {nr_eu(peste[0]['putere_totala'])} MW).\n\nVerifică "
                             "„Putere max. admisă” (din ATR) sau alege alt brand / nr. de skiduri.")

    specs = baterie["specs"] if baterie.get("manual") else specs_baterie(date.get("specs"), baterie["tip"])
    specs_lipsa = None
    if specs is None:
        specs_lipsa = ("Specificatii_baterii_BESS.xlsx nu a fost găsit" if date.get("specs") is None else
                       f"modelul {baterie['tip']} nu e în Specificatii_baterii_BESS.xlsx")
        specs = {"model": baterie["tip"], "config": "", "weight": ""}
    baterie["specs"] = specs

    lista, exact = filtreaza_skiduri(optiuni, intrari["mw"])
    # trafo-ul aux extern depinde de skidul ales: e în fiecare opțiune (o["trafo"]); aici cel al celei mai ieftine
    return {"baterie": baterie, "nr_baterii": nr_baterii, "optiuni": optiuni, "motive": motive,
            "lista": lista, "exact": exact, "trafo": lista[0]["trafo"], "specs_lipsa": specs_lipsa}


def aplica_rezerva_skid(optiuni, pct):
    """Rezervă de preț pe skid (% peste prețul din Pricing / manual); reordonează după costul nou."""
    if not pct:
        return
    for o in optiuni:
        o["pret_unit_baza"], o["rezerva_pct"] = o["pret_unit"], pct
        o["pret_unit"] = o["pret_unit"] * (1 + pct / 100)
        o["pret_total"] = o["nr"] * o["pret_unit"]
        o["cost_total"] = o["pret_total"] + (o["trafo"]["cost"] if o.get("trafo") else 0)
    optiuni.sort(key=lambda o: o["cost_total"])


def genereaza_oferta(date, intrari, baterie, nr_baterii, skid, trafo, status=lambda text: None, nume_existente=()):
    """Calculul în Excel + prezentarea PPTX + PDF, dintr-o dată (modul motor). Returnează căile fișierelor."""
    calcul = calculeaza_oferta(date, intrari, baterie, nr_baterii, skid, trafo, status, nume_existente,
                               online=xw is None and DRIVE is not None)
    return finalizeaza_oferta(calcul, intrari, baterie, nr_baterii, skid, trafo, status, reciteste=False)


# Celulele din Excelul de calcul citite după verificare (dacă ai modificat ceva în Excel, prezentarea ia valorile noi)
CELULE_REZULTAT = {"pret_ddp": "H13", "pret_mwh": "H14", "ems_yearly": "K5", "cost": "G13",
                   "ext_gar_bat": CELULA_EXT_GAR_BAT, "ext_gar_pcs": CELULA_EXT_GAR_PCS, "nr_trafo": "C7"}


def citeste_rezultate_excel(cale):
    """Rezultatele din Excelul de calcul salvat (valorile calculate de Excel la ultima salvare)."""
    from openpyxl import load_workbook
    wb = load_workbook(cale, data_only=True, read_only=True)
    sh = wb["Sheet1"]
    rez = {cheie: sh[celula].value for cheie, celula in CELULE_REZULTAT.items()}
    tabel = []
    for rand in sh.iter_rows(min_row=4, max_row=14, max_col=9, values_only=True):
        tabel.append(list(rand))
    wb.close()
    if rez["pret_ddp"] is None:
        raise ValueError("Excelul de calcul nu are valorile calculate (H13 e gol). Deschide-l, salvează-l (Ctrl+S) "
                         "și încearcă din nou.")
    rez = {k: (v if v is not None else 0) for k, v in rez.items()}
    rez["tabel"] = tabel
    return rez


def citeste_rezultate_online(fisier):
    """Ca citeste_rezultate_excel, dar din Excelul din Drive, prin Excel Online (vede și modificările făcute în browser)."""
    foaie = FoaieGraph(DRIVE, fisier)
    rez = {cheie: foaie.citeste(celula) for cheie, celula in CELULE_REZULTAT.items()}
    if rez["pret_ddp"] is None:
        raise ValueError("Excelul de calcul din Drive nu are prețul total (H13 e gol).")
    rez = {k: (v if v is not None else 0) for k, v in rez.items()}
    rez["tabel"] = [[None if v == "" else v for v in r] for r in foaie.citeste("A4:I14")]
    return rez


def citeste_rezultate(calcul):
    """Rezultatele calculului: din Excel Online (mod Microsoft 365) sau din fișierul local salvat."""
    if calcul["rez"].get("excel_drive"):
        return citeste_rezultate_online(calcul["rez"]["excel_drive"])
    return citeste_rezultate_excel(calcul["cale_excel"])


def e_calcul_test(intrari):
    """Calcul de probă (estimare rapidă): Client = 'test' și Nume Proiect = 'test'."""
    return all(str(intrari.get(c, "")).strip().lower() == TEXT_TEST for c in ("client", "proiect"))


def pregateste_folder_oferta(intrari, nume_existente=(), info_drive=None):
    """Stabilește intrari['versiune'] și creează folderul local al calculului (calculul de probă: în '_teste')."""
    if e_calcul_test(intrari):
        intrari["versiune"] = urmatoarea_versiune(intrari["nr_proiect"], intrari["proiect"], nume_existente)
        zi, nume_folder = FOLDER_TESTE, nume_oferta(intrari)
    else:
        intrari["versiune"], existent = versiune_de_lucru(intrari["nr_proiect"], intrari["proiect"],
                                                          nume_existente, info_drive)
        zi, nume_folder = existent or (datetime.now().strftime(FORMAT_FOLDER_ZI), nume_oferta(intrari))
    folder = os.path.abspath(os.path.join(FOLDER_OFERTE, zi, nume_folder))
    os.makedirs(folder, exist_ok=True)
    return folder


def aliniaza_folder(calcul):
    """Folderul unei versiuni refolosite poate avea alt MWh / MW în nume decât oferta finală: îl redenumim
    (local și în Drive) ca să aibă același nume ca prezentarea."""
    vechi = calcul["folder"]
    if os.path.basename(vechi) == calcul["nume"]:
        return
    nou = os.path.join(os.path.dirname(vechi), calcul["nume"])
    if not os.path.exists(nou):
        os.rename(vechi, nou)
    else:                                                 # există deja: mutăm fișierele în el
        for f in os.listdir(vechi):
            shutil.move(os.path.join(vechi, f), os.path.join(nou, f))
        shutil.rmtree(vechi, ignore_errors=True)
    if DRIVE is not None:
        try:
            DRIVE.redenumeste_oferta(os.path.basename(os.path.dirname(vechi)), os.path.basename(vechi), calcul["nume"])
        except Exception:
            pass
    calcul["folder"] = nou
    calcul["cale_excel"] = os.path.join(nou, os.path.basename(calcul["cale_excel"]))
    calcul["rez"]["cale_excel"] = calcul["cale_excel"]


def calculeaza_oferta(date, intrari, baterie, nr_baterii, skid, trafo, status=lambda text: None, nume_existente=(),
                      online=False, info_drive=None):
    """Pasul 1: doar Excelul de calcul, în 'Oferte generate/<nume>_vN'. Prezentarea se face după verificare."""
    # Folderul, prezentarea și PDF-ul au același nume, ex. ST202606109_BESS_..._20MWh_10MW_v2. Versiunea crește
    # doar după o prezentare: calculele refăcute înainte de OK stau toate în folderul aceleiași versiuni.
    folder = pregateste_folder_oferta(intrari, nume_existente, info_drive)
    nume = nume_oferta(intrari)

    status("Calculez în Excel...")
    c_brand_bat = date["bat"].columns[BAT_COL_BRAND]
    branduri_bat = [str(x) for x in date["bat"][c_brand_bat].dropna().unique() if str(x).strip()]
    calcul_excel = calculeaza_in_excel_online if online else calculeaza_in_excel
    rez = calcul_excel(intrari, baterie, nr_baterii, skid, trafo, folder, branduri_bat, branduri_skid(date))
    calcul = {"nume": nume, "folder": folder, "rez": rez, "cale_excel": rez["cale_excel"], "online": online}
    try:
        rez["tabel"] = citeste_rezultate(calcul)["tabel"]
    except Exception:
        rez["tabel"] = []
    return calcul


def antet_calcul(intrari):
    """Celulele cu datele ofertei din Excelul de calcul (aceleași ca în completeaza_calcul)."""
    return [("A1:B2", [["Client", intrari["client"]], ["Nume Proiect", intrari["proiect"]]]),
            ("E2:E3", [[intrari["nr_proiect"]], [f"v{intrari['versiune']}"]])]


def transforma_calcul_test(calcul, intrari, nume_existente=(), info_drive=None):
    """Calculul de probă devine ofertă: Excelul de probă (cu modificările tale) se copiază în folderul ofertei, cu
    clientul / proiectul / nr. proiect / versiunea reale. intrari = intrările ofertei (fără 'test')."""
    folder = pregateste_folder_oferta(intrari, nume_existente, info_drive)
    acum = datetime.now()
    cale = os.path.join(folder, nume_calcul(intrari, acum))
    rez = dict(calcul["rez"])
    if calcul.get("online"):
        id_folder = DRIVE.folder_oferta(folder)
        DRIVE._descarca(*rez["excel_drive"], cale)          # ultima variantă (poate modificată în browser)
        fisier = DRIVE.urca_fisier(id_folder, os.path.basename(cale), cale)
        foaie = FoaieGraph(DRIVE, fisier)
        for adresa, valori in antet_calcul(intrari):
            foaie.scrie(adresa, valori)
        foaie.calculeaza()
        try:
            DRIVE._descarca(*fisier, cale)
        except Exception:
            pass
        rez.update(excel_drive=fisier, folder_drive=(fisier[0], id_folder), link_excel=DRIVE.link(fisier))
    else:
        if xw is None:
            raise RuntimeError("Excel nu e instalat (sau lipsește biblioteca xlwings).")
        shutil.copy2(calcul["cale_excel"], cale)
        app_xl = xw.App(visible=(sys.platform == "darwin"), add_book=False)
        wb = None
        try:
            wb = app_xl.books.open(cale)
            foaie = FoaieXlwings(wb.sheets["Sheet1"])
            for adresa, valori in antet_calcul(intrari):
                foaie.scrie(adresa, valori)
            foaie.calculeaza()
            wb.save()
        finally:
            try:
                if wb is not None:
                    wb.close()
                if sys.platform != "darwin" or not app_xl.books:
                    app_xl.quit()
            except Exception:
                pass
    rez["cale_excel"] = cale
    return {"nume": nume_oferta(intrari), "folder": folder, "rez": rez, "cale_excel": cale,
            "online": calcul.get("online")}


FISIER_DATE_OFERTA = "date_oferta.json"          # datele din formular, salvate în folderul fiecărei oferte
MODEL_NUME_OFERTA = re.compile(r"^(?P<nr>[^_]+)_BESS_Indicative_Proposal_Taliva_(?P<proiect>.+)_(?P<mwh>[\d.,]+)MWh_"
                               r"(?P<mw>[\d.,]+)MW_v(?P<v>\d+)$")


def date_din_nume_oferta(nume):
    """Pentru ofertele vechi (fără date_oferta.json): nr., proiect, MWh, MW, versiunea din numele folderului."""
    m = MODEL_NUME_OFERTA.match(str(nume).strip())
    if not m:
        return None
    return {"Nr. proiect:": m["nr"], "Nume Proiect:": m["proiect"], "Capacitate (MWh):": m["mwh"],
            "Putere (MW):": m["mw"], "_versiune": int(m["v"])}


def nume_skid_prezentare(skid):
    """Numele skidului din prezentare: dacă are trafo aux inclus (coloana din setări) și numele nu-l scrie deja,
    se adaugă ' + AUX TRAFO 150kVA'."""
    nume = str(skid["tip"]).strip()
    aux = skid.get("aux_skid") or 0
    if aux > 0 and not re.search(r"aux", nume, re.I):
        nume += f" + AUX TRAFO {aux:g}kVA"
    return nume


def finalizeaza_oferta(calcul, intrari, baterie, nr_baterii, skid, trafo, status=lambda text: None, reciteste=True):
    """Pasul 2 (după OK): prezentarea PPTX + PDF, cu valorile din Excelul de calcul (recitit, dacă l-ai modificat)."""
    aliniaza_folder(calcul)
    nume, folder, rez = calcul["nume"], calcul["folder"], dict(calcul["rez"])
    if reciteste:
        status("Recitesc Excelul de calcul...")
        rez.update(citeste_rezultate(calcul))
    status("Populez prezentarea PowerPoint...")
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
        "[PROP_MW]": text_putere_propusa(skid["putere_totala"], intrari["mw"]),   # cu limitare EMS
        "[BAT_BRAND]": baterie["brand"],
        "[BAT_NAME]": baterie["specs"]["model"],      # codul real (ex. Y104R04C12), nu 'REPT 5,015'
        "[NR_BAT]": str(nr_baterii),
        "[SKID_BRAND]": skid["brand"],
        "[SKID]": nume_skid_prezentare(skid),           # + 'AUX TRAFO 150kVA' din setări, dacă nu e în nume
        # trafo aux extern lângă numele skidului (gol dacă nu e nevoie): ' + Aux transformer 400 kVA'
        "[SKID_AUX]": (f" + {'' if trafo['cant'] == 1 else str(trafo['cant']) + ' x '}Aux transformer "
                       f"{nr_eu(trafo['kva'])} kVA") if trafo else "",
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
        # la componentele manuale K9 / K11 vin din procentul introdus (fără procent, rândul dispare din tabel)
        "[EXT_GAR_BAT]": bani_eu(rez["ext_gar_bat"]),
        "[EXT_GAR_PCS]": bani_eu(rez["ext_gar_pcs"]),
    }
    cale_pptx = os.path.join(folder, nume + ".pptx")
    ramase = genereaza_pptx(inlocuiri, cale_pptx,
                            fara_bat=fara_extensie(f"{baterie['brand']} {baterie['tip']}")
                                     or bool(baterie.get("manual") and not baterie.get("ext_gar_pct")),
                            fara_pcs=fara_extensie(skid["brand"])
                                     or bool(skid.get("manual") and not skid.get("ext_gar_pct")))

    status("Salvez și varianta PDF...")
    cale_pdf, eroare_pdf = os.path.join(folder, nume + ".pdf"), None
    try:
        if calcul.get("online"):
            pptx_in_pdf_online(cale_pptx, cale_pdf, folder)          # PDF făcut de Microsoft 365
        else:
            pptx_in_pdf(cale_pptx, cale_pdf)
    except Exception as e:
        cale_pdf, eroare_pdf = None, str(e)
    return {"nume": nume, "folder": folder, "cale_pptx": cale_pptx, "cale_pdf": cale_pdf,
            "cale_excel": rez["cale_excel"], "eroare_pdf": eroare_pdf, "ramase": ramase, "rez": rez}
