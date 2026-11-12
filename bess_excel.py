# -*- coding: utf-8 -*-
"""Excelul de calcul (Excel instalat sau Excel Online).
Parte din generator_oferte_bess.py: se încarcă în același spațiu de nume (toate funcțiile
și setările sunt comune), deci aici NU se fac importuri proprii și nu se rulează direct."""
# pylint: disable=undefined-variable
# flake8: noqa

# ======================================================================
# EXCEL
# ======================================================================
class FoaieXlwings:
    """Foaia de calcul deschisă în Excel instalat (xlwings)."""
    def __init__(self, sh):
        self.sh = sh

    def scrie(self, adresa, valoare):
        self.sh.range(adresa).value = valoare

    def format_numar(self, adresa, fmt):
        self.sh.range(adresa).number_format = fmt

    def ingroasa(self, adresa):
        self.sh.range(adresa).font.bold = True

    def formule(self):
        """(rândul și coloana de start 1-based, formulele zonei folosite ca listă de rânduri)."""
        zona = self.sh.used_range
        formule = zona.formula
        if not isinstance(formule, (list, tuple)):
            formule = ((formule,),)
        elif formule and not isinstance(formule[0], (list, tuple)):
            formule = (formule,)
        return zona.row, zona.column, [list(r) for r in formule]

    def citeste(self, adresa):
        return self.sh.range(adresa).value

    def calculeaza(self):
        self.sh.book.app.calculate()


def inlocuieste_branduri_foaie(foaie, branduri, brand_ales):
    """În celulele cu text (nu formule) din foaie, înlocuiește orice brand din listă cu brandul ales.
    Ex.: 'Commissioning Sinexcel' -> 'Commissioning Kehua'."""
    branduri = [b for b in branduri if b and b.strip().lower() != brand_ales.strip().lower()]
    if not branduri or not brand_ales:
        return
    model = re.compile("|".join(re.escape(b.strip()) for b in sorted(branduri, key=len, reverse=True)), re.I)
    rand0, col0, formule = foaie.formule()
    for i, rand in enumerate(formule):
        for j, val in enumerate(rand):
            if isinstance(val, str) and val and not val.startswith("=") and model.search(val):
                foaie.scrie(f"{litera_coloana(col0 - 1 + j)}{rand0 + i}", model.sub(brand_ales.strip(), val))


def completeaza_calcul(foaie, intrari, baterie, nr_baterii, skid, trafo, acum, branduri_bat=(), branduri_skid=()):
    """Completează template-ul de calcul (aceleași celule, indiferent dacă e Excel instalat sau Excel Online)."""
    # Antet: client, proiect, nr. proiect, data și ora calculului
    foaie.scrie("A1:B2", [["Client", intrari["client"]], ["Nume Proiect", intrari["proiect"]]])
    foaie.scrie("D1:E3", [["Data și ora", acum], ["Nr. proiect", intrari["nr_proiect"]],
                          ["Versiune", f"v{intrari['versiune']}"]])
    foaie.format_numar("E1", "dd.mm.yyyy hh:mm:ss")
    foaie.ingroasa("A1:A2")
    foaie.ingroasa("D1:D3")

    # rezerva de preț apare ca formulă (preț de bază x (1 + %)), ca să se vadă în Excel
    pret_bat = (f"={baterie['pret_baza']}*(1+{baterie['rezerva_pct']:g}%)" if baterie.get("rezerva_pct")
                else baterie["pret"])
    pret_skid = (f"={skid['pret_unit_baza']}*(1+{skid['rezerva_pct']:g}%)" if skid.get("rezerva_pct")
                 else skid["pret_unit"])
    foaie.scrie("A5:D5", [[baterie["tip"], baterie["cap"], nr_baterii, pret_bat]])
    foaie.scrie("F5", baterie["transport"])

    foaie.scrie("A6:D6", [[skid["tip"], skid["putere_unit"], skid["nr"], pret_skid]])   # B8 = B6 x C6
    foaie.scrie("F6", skid["transport"])
    foaie.scrie("G12", f"={skid['com']}*C6")

    # Trafo auxiliar extern: nr. bucăți (C7 = B21) îl calculează Excelul, în blocul 'CALCUL TRAFO AUXILIAR'.
    # B19 = kVA de acoperit cu trafo extern (necesar baterii + skiduri minus trafo-ul aux din skid).
    foaie.scrie("A7", trafo["nume"] if trafo else "aux trafo")
    # B20 = puterea unui trafo; B7 = puterea totală = putere unit x nr. trafo (C7 = B21, calculat de Excel)
    foaie.scrie("B20", trafo["kva"] if trafo else None)
    foaie.scrie("B7", "=B20*C7" if trafo else None)
    foaie.scrie("D7", trafo["pret"] if trafo else 0)
    foaie.scrie("B17:B18", [[1 if trafo else 0], [baterie.get("kva_aux") or kva_aux_baterie(baterie["cap"])]])
    foaie.scrie("A19:B19", [["kVA de acoperit cu trafo extern", skid.get("kva_trafo", 0)]])
    foaie.scrie("D17:E20", [["Necesar aux total (kVA)", skid.get("kva_necesar", 0)],
                            ["  din care consum skiduri (kVA)" if CONSUM_SKID_IN_NECESAR else
                             "  consum skiduri: trafo propriu separat", skid["nr"] * skid.get("kva_skid", 0)],
                            ["Acoperit de trafo aux din skid (kVA)", skid.get("kva_acoperit", 0)],
                            ["Trafo aux inclus / skid (kVA)", skid.get("aux_skid", 0)]])
    foaie.scrie(f"A24:B{23 + len(KVA_AUX_PER_BATERIE)}", [[prag, kva] for prag, kva in KVA_AUX_PER_BATERIE])

    foaie.scrie("E9:E10", [[intrari["ems"]], [intrari["eng"]]])
    foaie.scrie("H3:I3", [[intrari["adaos"], intrari["curs"]]])
    rezerve = [f"{n} +{c['rezerva_pct']:g}%" for n, c in (("baterie", baterie), ("skid", skid)) if c.get("rezerva_pct")]
    if rezerve:
        foaie.scrie("B3", "Rezervă preț: " + ", ".join(rezerve))

    # Etichetele din template (ex. 'Commissioning Sinexcel') primesc brandurile alese
    inlocuieste_branduri_foaie(foaie, branduri_skid, skid["brand"])
    inlocuieste_branduri_foaie(foaie, branduri_bat, baterie["brand"])

    # Componente manuale: extra warranty = % din totalul cu transport + commissioning (K9 / K11 = $/MWh)
    for comp, celula_pct, eticheta, total in ((baterie, "K6", f"Battery {baterie['brand']}", "(G5+G11)"),
                                              (skid, "K10", f"PCS {skid['brand']}", "(G6+G12)")):
        if comp.get("manual"):
            pct = comp.get("ext_gar_pct") or 0
            foaie.scrie("J" + celula_pct[1:], f"{eticheta} {pct:g}%" if pct else "fără extra warranty")
            foaie.scrie(celula_pct, f"={total}*{pct / 100:g}" if pct else "=0")

    foaie.calculeaza()
    return {
        "pret_ddp": foaie.citeste("H13") or 0,
        "pret_mwh": foaie.citeste("H14") or 0,
        "ems_yearly": foaie.citeste("K5") or 0,
        "cost": foaie.citeste("G13") or 0,                      # valoarea de achiziție (pentru tabelul Moldova)
        "ext_gar_bat": foaie.citeste(CELULA_EXT_GAR_BAT) or 0,
        "ext_gar_pcs": foaie.citeste(CELULA_EXT_GAR_PCS) or 0,
        "nr_trafo": foaie.citeste("C7") or 0,
    }


def nume_calcul(intrari, acum):
    return "Calcul_{}_{}_{}_v{}_{}.xlsx".format(
        *(nume_fisier_valid(x) for x in (intrari["nr_proiect"], intrari["client"], intrari["proiect"])),
        intrari["versiune"], acum.strftime("%Y-%m-%d_%H.%M.%S"))


def calculeaza_in_excel(intrari, baterie, nr_baterii, skid, trafo, folder, branduri_bat=(), branduri_skid=()):
    """Completează template-ul, recalculează, salvează calculul într-un fișier nou. Returnează rezultatele."""
    if xw is None:
        raise RuntimeError("Excel nu e instalat (sau lipsește biblioteca xlwings): bifează „Calcul și PDF prin "
                           "Microsoft 365”.")
    app_xl = None
    wb = None
    try:
        # Pe Mac există o singură instanță Excel (a utilizatorului): nu o ascundem și nu o închidem dacă are alte fișiere
        app_xl = xw.App(visible=(sys.platform == "darwin"), add_book=False)
        wb = app_xl.books.open(os.path.abspath(FISIER_EXCEL_TEMPLATE))
        sh = wb.sheets["Sheet1"]

        acum = datetime.now()
        rezultat = completeaza_calcul(FoaieXlwings(sh), intrari, baterie, nr_baterii, skid, trafo, acum,
                                      branduri_bat, branduri_skid)

        # Fiecare calcul = un fișier Excel separat, cu data și ora în nume (nu se suprascrie nimic)
        cale = os.path.abspath(os.path.join(folder, nume_calcul(intrari, acum)))
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


# ---------------- același calcul, prin Excel Online (Microsoft 365) ----------------
def _valoare_excel(v):
    """Valoare pentru Excel Online: datetime -> număr de serie Excel, None -> '' (golește celula)."""
    if isinstance(v, datetime):
        return (v - datetime(1899, 12, 30)).total_seconds() / 86400
    return "" if v is None else v


class FoaieGraph:
    """Foaia de calcul dintr-un fișier din Drive, completată și calculată de Excel Online (fără Excel instalat).
    Scrierile se adună și se trimit grupat (Graph $batch, câte 20, una după alta)."""
    def __init__(self, drive, fisier, nume_foaie="Sheet1"):
        self.drive, self.fisier, self.nume_foaie = drive, fisier, nume_foaie
        self.url = drive._url_foaie(fisier, nume_foaie)
        self.de_trimis = []

    def _range(self, adresa):
        return f"{self.url}/range(address='{adresa}')"

    def scrie(self, adresa, valoare):
        randuri = valoare if isinstance(valoare, list) else [[valoare]]
        randuri = [[_valoare_excel(v) for v in r] for r in randuri]
        cheie = "formulas" if any(isinstance(v, str) and v.startswith("=") for r in randuri for v in r) else "values"
        self.de_trimis.append(("PATCH", self._range(adresa), {cheie: randuri}))

    def format_numar(self, adresa, fmt):
        self.de_trimis.append(("PATCH", self._range(adresa), {"numberFormat": [[fmt]]}))

    def ingroasa(self, adresa):
        self.de_trimis.append(("PATCH", self._range(adresa) + "/format/font", {"bold": True}))

    def trimite(self):
        while self.de_trimis:
            grup, self.de_trimis = self.de_trimis[:20], self.de_trimis[20:]
            self.drive.batch(grup)

    def formule(self):
        self.trimite()
        j = self.drive.cerere_excel("GET", f"{self.url}/usedRange?$select=address,formulas")
        m = re.search(r"!?\$?([A-Z]+)\$?(\d+)", j["address"].split("!")[-1])
        return int(m.group(2)), _index_coloana(m.group(1)), j["formulas"]

    def calculeaza(self):
        self.trimite()
        self.drive.cerere_excel("POST", self.url.rsplit("/worksheets/", 1)[0] + "/application/calculate",
                                {"calculationType": "Full"})

    def citeste(self, adresa):
        self.trimite()
        v = self.drive.cerere_excel("GET", self._range(adresa) + "?$select=values")["values"]
        valoare = v[0][0] if len(v) == 1 and len(v[0]) == 1 else v
        return None if valoare == "" else valoare


def calculeaza_in_excel_online(intrari, baterie, nr_baterii, skid, trafo, folder, branduri_bat=(), branduri_skid=()):
    """Ca calculeaza_in_excel, dar fără Excel instalat: template-ul se copiază în Drive, în folderul ofertei
    ('Oferte generate/<zi>/<nume>/Calcul_....xlsx'), Excel Online îl completează și îl calculează, iar o copie se
    descarcă local. Excelul din Drive e cel 'oficial' (îl poți deschide în browser și modifica)."""
    if DRIVE is None:
        raise RuntimeError("Calculul prin Microsoft 365 merge doar cu Drive-ul conectat.")
    acum = datetime.now()
    nume = nume_calcul(intrari, acum)
    id_folder = DRIVE.folder_oferta(folder)
    fisier = DRIVE.urca_fisier(id_folder, nume, FISIER_EXCEL_TEMPLATE)
    rezultat = completeaza_calcul(FoaieGraph(DRIVE, fisier), intrari, baterie, nr_baterii, skid, trafo, acum,
                                  branduri_bat, branduri_skid)
    cale = os.path.abspath(os.path.join(folder, nume))
    try:
        DRIVE._descarca(fisier[0], fisier[1], cale)
    except Exception:
        pass                                              # copia locală e doar pentru arhivă; contează cea din Drive
    rezultat.update(cale_excel=cale, excel_drive=fisier, folder_drive=(fisier[0], id_folder),
                    link_excel=DRIVE.link(fisier))
    return rezultat


def pptx_in_pdf_online(cale_pptx, cale_pdf, folder):
    """PDF-ul făcut de Microsoft 365: prezentarea se urcă în folderul ofertei din Drive și se descarcă ca PDF."""
    id_folder = DRIVE.folder_oferta(folder)
    fisier = DRIVE.urca_fisier(id_folder, os.path.basename(cale_pptx), cale_pptx)
    continut = DRIVE.cerere_excel("GET", f"/drives/{fisier[0]}/items/{fisier[1]}/content?format=pdf", brut=True)
    with open(cale_pdf, "wb") as f:
        f.write(continut)
