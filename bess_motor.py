# -*- coding: utf-8 -*-
"""Modul motor (cereri din Power Apps).
Parte din generator_oferte_bess.py: se încarcă în același spațiu de nume (toate funcțiile
și setările sunt comune), deci aici NU se fac importuri proprii și nu se rulează direct."""
# pylint: disable=undefined-variable
# flake8: noqa

# ======================================================================
# MOD MOTOR: cererile vin din Power Apps / Power Automate prin OneDrive
# ======================================================================
# Power Automate scrie cererile ca fișiere .json în 'Ofertare BESS/Cereri'. Motorul (acest program pornit cu
# argumentul --motor, pe un PC Windows cu Excel și PowerPoint) le generează și scrie rezultatul în
# 'Cereri/Rezultate/<id>.json'; Power Automate trimite apoi emailul cu fișierele.
FOLDER_CERERI = "Cereri"
FOLDER_REZULTATE = os.path.join(FOLDER_CERERI, "Rezultate")
FOLDER_PROCESATE = os.path.join(FOLDER_CERERI, "Procesate")
FISIER_LISTE = os.path.join(FOLDER_CERERI, "Liste.xlsx")     # listele pentru meniurile din Power Apps
SECUNDE_VERIFICARE = 10
SECUNDE_ASTEPTARE_SINCRONIZARE = 5    # o cerere se procesează abia când fișierul nu s-a mai modificat de atâtea secunde


def numar_cerere(valoare, implicit=None):
    """'1,10' / '1.10' / 1.1 -> 1.1; gol -> implicit (sau eroare dacă implicit e None)."""
    if valoare is None or str(valoare).strip() == "":
        if implicit is None:
            raise ValueError("lipsește o valoare numerică obligatorie")
        return implicit
    return float(str(valoare).strip().replace(" ", "").replace(",", "."))


def intrari_din_cerere(c):
    """Transformă cererea JSON (din Power Apps) în aceleași intrări ca formularul din fereastră."""
    mwh = numar_cerere(c.get("mwh"))
    return {
        "client": str(c.get("client", "")).strip(),
        "proiect": str(c.get("proiect", "")).strip(),
        "locatie": str(c.get("locatie", "")).strip(),
        "nr_proiect": str(c.get("nr_proiect", "")).strip()
                      or NR_PROIECT_IMPLICIT.format(an_luna=datetime.now().strftime("%Y%m")),
        "mwh": mwh,
        "mw": numar_cerere(c.get("mw")),
        # EMS: dacă nu e completat, aceeași regulă ca în fereastră (sub prag / de la prag)
        "ems": numar_cerere(c.get("ems"), ems_pentru(mwh)),
        "ems_auto": not str(c.get("ems", "") or "").strip(),
        "eng": numar_cerere(c.get("eng"), 0.0),
        "adaos": numar_cerere(c.get("adaos"), 1.10),
        "curs": numar_cerere(c.get("curs"), curs_salvat()[0]),
        "poc_kv": numar_cerere(c.get("poc_kv"), float(POC_IMPLICIT_KV)),
    }


def cale_onedrive(cale_locala):
    """Calea fișierului în OneDrive, pentru Power Automate: '/Ofertare BESS/Oferte generate/.../x.pdf'."""
    if not cale_locala:
        return None
    radacina = os.path.dirname(os.getcwd())
    return "/" + os.path.relpath(cale_locala, radacina).replace(os.sep, "/")


def limitat_prin_ems(putere_propusa, mw_cerut):
    """True dacă skidurile au mai multă putere decât s-a cerut -> puterea evacuată se limitează prin EMS."""
    return mw_cerut > 0 and putere_propusa - mw_cerut > 0.005


def text_putere_propusa(putere_propusa, mw_cerut):
    """[PROP_MW]: '6' sau, cu limitare prin EMS, '6 (limited by EMS to 5,85)'."""
    if limitat_prin_ems(putere_propusa, mw_cerut):
        return TEXT_LIMITARE_EMS.format(propus=nr_eu(putere_propusa), limitat=nr_eu(mw_cerut))
    return nr_eu(putere_propusa)


def ems_pentru(mwh):
    """Valoarea EMS (USD) după capacitate: de la EMS_PRAG_MWH în sus EMS_AUTO, sub prag EMS_SUB_PRAG."""
    return float(EMS_AUTO if mwh >= EMS_PRAG_MWH - 1e-9 else EMS_SUB_PRAG)


def ajusteaza_ems(intrari, p):
    """EMS automat după capacitatea OFERTATĂ (nr. baterii x capacitate), nu după cea cerută:
    ex. cerut 13,5 MWh -> 3 x 5,015 = 15,045 MWh -> EMS de la prag. Doar dacă EMS nu a fost scris de mână.
    Returnează valoarea nouă sau None dacă nu s-a schimbat nimic."""
    if not intrari.get("ems_auto"):
        return None
    ems = ems_pentru(p["nr_baterii"] * p["baterie"]["cap"])
    if ems == intrari["ems"]:
        return None
    intrari["ems"] = ems
    return ems


def pregateste_din_cerere(date, baterii_map, c):
    """Intrările, avertismentele și opțiunile (baterie, skiduri, trafo) pentru o cerere din Power Apps."""
    intrari = intrari_din_cerere(c)
    erori, avert = valideaza(intrari)
    if erori:
        raise ValueError("\n".join(erori))
    nr = str(c.get("nr_skiduri", "")).strip()
    p = pregateste_oferta(date, baterii_map, intrari, str(c.get("baterie", BATERIE_IMPLICITA)).strip(),
                          str(c.get("incoterm", INCOTERM_IMPLICIT)).strip().upper(),
                          str(c.get("brand", "") or "Oricare").strip(),
                          int(nr) if nr.isdigit() else None,
                          str(c.get("trafo", "")).strip().lower() in ("true", "1", "da", "yes"))
    if p["specs_lipsa"]:
        avert.append(f"{p['specs_lipsa']}: Configuration și Weight vor fi goale în prezentare.")
    ajusteaza_ems(intrari, p)
    return intrari, avert, p


def previzualizeaza_cerere(date, baterii_map, c):
    """Răspunsul la o cerere de previzualizare: bateria, skidurile compatibile (de ales) și trafo-ul.
    Nu deschide Excel / PowerPoint, deci e rapid."""
    intrari, avert, p = pregateste_din_cerere(date, baterii_map, c)
    b, nr = p["baterie"], p["nr_baterii"]
    trafo = p["lista"][0]["trafo"]
    return {
        "status": "previzualizare",
        "baterie": (f"{nr} x {b['tip']} = {nr_eu(round(nr * b['cap'], 3))} MWh | {bani_eu(b['pret'])} $/MWh "
                    f"({b['incoterm']}, cotație {b['data']})"),
        "trafo": (f"{trafo['cant']} x {trafo['nume']} ({nr_eu(trafo['kva_total'])} kVA de acoperit) = "
                  f"{bani_eu(trafo['cost'], 0)} $") if trafo else "",
        "exista_putere_exacta": p["exact"],
        "optiuni": [{"skid": o["tip"],
                     "text": f"{o['nr']} x {o['tip']} = {nr_eu(o['putere_totala'])} MW | {bani_eu(o['cost_total'], 0)} $",
                     "detalii": (("✓ putere exactă" if abs(o["putere_totala"] - intrari["mw"]) < 0.01
                                  else f"limitată prin EMS la {nr_eu(intrari['mw'])} MW")
                                 + f" | intrări DC: {o['intrari']:g} (acceptă {o['acceptate']})"
                                 + (f" | {o['conexiune']}" if o.get("conexiune") else "")
                                 + f" | {text_aux(o)}")}
                    for o in p["lista"]],
        "toate_skidurile": raport_skiduri(p["optiuni"], p["motive"]),
        "avertismente": avert,
    }


def proceseaza_cerere(date, baterii_map, c):
    """Generează oferta pentru o cerere și returnează rezultatul (dicționar scris apoi în Rezultate/)."""
    intrari, avert, p = pregateste_din_cerere(date, baterii_map, c)
    # Skid ales explicit (numele exact din previzualizare), altfel cel mai ieftin
    skid_cerut = str(c.get("skid", "")).strip().lower()
    skid = next((o for o in p["optiuni"] if o["tip"].strip().lower() == skid_cerut), p["lista"][0])

    r = genereaza_oferta(date, intrari, p["baterie"], p["nr_baterii"], skid, skid["trafo"])
    if r["eroare_pdf"]:
        avert.append(f"PDF-ul nu s-a putut crea: {r['eroare_pdf']}")
    if r["ramase"]:
        avert.append("Placeholdere rămase neînlocuite în prezentare: " + ", ".join(r["ramase"]))

    alternative = [f"{o['nr']} x {o['tip']} = {nr_eu(o['putere_totala'])} MW | {bani_eu(o['pret_total'], 0)} $"
                   for o in p["lista"] if o is not skid]
    return {
        "status": "ok",
        "nume": r["nume"],
        "pret_total_usd": bani_eu(r["rez"]["pret_ddp"]),
        "pret_mwh_usd": bani_eu(r["rez"]["pret_mwh"]),
        "baterie": f"{p['nr_baterii']} x {p['baterie']['tip']}",
        "skid": f"{skid['nr']} x {skid['tip']} ({nr_eu(skid['putere_totala'])} MW, {bani_eu(skid['pret_total'], 0)} $)",
        "alternative": alternative,
        "avertismente": avert,
        "fisier_pptx": cale_onedrive(r["cale_pptx"]),
        "fisier_pdf": cale_onedrive(r["cale_pdf"]),
        "fisier_excel": cale_onedrive(r["cale_excel"]),
    }


def scrie_json(cale, continut):
    """Scrie întâi într-un fișier temporar și abia apoi îl redenumește, ca OneDrive să nu urce un fișier pe jumătate."""
    temp = cale + ".tmp"
    with open(temp, "w", encoding="utf-8") as f:
        json.dump(continut, f, ensure_ascii=False, indent=2)
    os.replace(temp, cale)


def exporta_liste(date, baterii_map):
    """Cereri/Liste.xlsx cu tabele Excel (Baterii, Branduri, POC) pentru meniurile din Power Apps.
    Se rescrie doar dacă s-a schimbat conținutul."""
    from openpyxl import Workbook
    from openpyxl.worksheet.table import Table
    liste = {
        "Baterii": ("Model", list(baterii_map)),
        "Branduri": ("Brand", ["Oricare"] + branduri_skid(date)),
        "POC": ("kV", [nr_eu(POC_IMPLICIT_KV)] + [nr_eu(kv) for kv in tensiuni_poc(date.get("skid_poc"))
                                                    if kv != POC_IMPLICIT_KV]),
    }
    semnatura = json.dumps(liste, ensure_ascii=False, sort_keys=True)
    cale_semnatura = FISIER_LISTE + ".continut"
    try:
        with open(cale_semnatura, encoding="utf-8") as f:
            if f.read() == semnatura and os.path.exists(FISIER_LISTE):
                return
    except OSError:
        pass
    wb = Workbook()
    wb.remove(wb.active)
    for nume, (coloana, valori) in liste.items():
        ws = wb.create_sheet(nume)
        ws.append([coloana])
        for v in valori:
            ws.append([v])
        ws.add_table(Table(displayName=f"Tabel{nume}", ref=f"A1:A{len(valori) + 1}"))
    temp = FISIER_LISTE + ".tmp.xlsx"
    wb.save(temp)
    os.replace(temp, FISIER_LISTE)
    with open(cale_semnatura, "w", encoding="utf-8") as f:
        f.write(semnatura)


class Motor:
    """Fereastră mică: verifică periodic folderul Cereri și generează ofertele cerute."""

    def __init__(self, root):
        self.root = root
        root.title("Generator Oferte BESS - Motor")
        root.geometry("560x260")
        ttk.Label(root, text="Motor oferte (Power Apps)", font=("Helvetica", 14, "bold")).pack(pady=10)
        self.lbl = ttk.Label(root, text="Pornesc...", justify="left", wraplength=520)
        self.lbl.pack(padx=15, anchor="w")
        self.jurnal = tk.Text(root, height=8, width=70, font=("Consolas", 9))
        self.jurnal.pack(padx=10, pady=10, fill="both", expand=True)
        for folder in (FOLDER_CERERI, FOLDER_REZULTATE, FOLDER_PROCESATE):
            os.makedirs(folder, exist_ok=True)
        self.date, self.baterii_map = None, {}
        self.root.after(500, self.verifica)

    def scrie(self, text):
        linie = f"{datetime.now():%d.%m %H:%M:%S}  {text}"
        self.jurnal.insert("end", linie + "\n")
        self.jurnal.see("end")
        try:
            with open(os.path.join(FOLDER_CERERI, "motor_log.txt"), "a", encoding="utf-8") as f:
                f.write(linie + "\n")
        except OSError:
            pass

    def verifica(self):
        try:
            if pricing_modificat(self.date):
                self.date = incarca_date()
                self.baterii_map = lista_baterii(self.date["bat"])
                exporta_liste(self.date, self.baterii_map)
                self.scrie(f"Prețuri încărcate din {self.date['sursa']}")
            acum = datetime.now().timestamp()
            cereri = sorted(f for f in os.listdir(FOLDER_CERERI) if f.lower().endswith(".json"))
            for fisier in cereri:
                cale = os.path.join(FOLDER_CERERI, fisier)
                if acum - os.path.getmtime(cale) >= SECUNDE_ASTEPTARE_SINCRONIZARE:
                    self.proceseaza(cale)
            self.lbl.config(text=f"Aștept cereri în: {os.path.abspath(FOLDER_CERERI)}\n"
                                 f"Ultima verificare: {datetime.now():%H:%M:%S}")
        except Exception as e:
            self.lbl.config(text=f"Eroare: {e}")
            self.scrie(f"EROARE: {e}")
        self.root.after(SECUNDE_VERIFICARE * 1000, self.verifica)

    def proceseaza(self, cale):
        fisier = os.path.basename(cale)
        id_cerere = os.path.splitext(fisier)[0]
        cerere = {}
        try:
            with open(cale, encoding="utf-8-sig") as f:
                cerere = json.load(f)
            id_cerere = str(cerere.get("id") or id_cerere)
            self.scrie(f"Cerere {id_cerere}: {cerere.get('client', '')} / {cerere.get('proiect', '')}")
            self.root.update()
            if str(cerere.get("tip", "")).strip().lower() == "previzualizare":
                rezultat = previzualizeaza_cerere(self.date, self.baterii_map, cerere)
                self.scrie(f"  previzualizare: {len(rezultat['optiuni'])} skiduri compatibile")
            else:
                rezultat = proceseaza_cerere(self.date, self.baterii_map, cerere)
                self.scrie(f"  gata: {rezultat['nume']}")
        except Exception as e:
            rezultat = {"status": "eroare", "mesaj": str(e)}
            self.scrie(f"  EROARE: {e}")
        rezultat.update(id=id_cerere, email=cerere.get("email", ""), client=cerere.get("client", ""),
                        proiect=cerere.get("proiect", ""), data=datetime.now().strftime("%d.%m.%Y %H:%M"))
        scrie_json(os.path.join(FOLDER_REZULTATE, f"{nume_fisier_valid(id_cerere)}.json"), rezultat)
        os.replace(cale, os.path.join(FOLDER_PROCESATE, fisier))


# Butonul albastru: pe Mac culoarea de fundal e ignorată (textul alb ar deveni invizibil) -> buton standard
STIL_BUTON = {} if sys.platform == "darwin" else {"bg": "#0052cc", "fg": "white"}
