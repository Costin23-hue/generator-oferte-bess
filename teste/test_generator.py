"""Teste automate pentru logica programului (alegerea bateriilor, skidurilor, trafo-ului, setări, Pricing).
Folosesc date INVENTATE (teste/date_test.py), nu fișierele reale. Nu pornesc Excel / PowerPoint / Drive.

Rulare (din folderul programului):   py -3.14 -m unittest discover -s teste -v
sau dublu-click pe 'Ruleaza teste.bat'.
"""
import os
import sys
import tempfile
import types
import unittest

AICI = os.path.dirname(os.path.abspath(__file__))
sys.path[:0] = [os.path.dirname(AICI), AICI]

try:
    import tkinter  # noqa: F401
except ImportError:                       # pe serverele fără interfață grafică
    for nume in ("tkinter", "tkinter.ttk", "tkinter.filedialog", "tkinter.messagebox"):
        sys.modules[nume] = types.ModuleType(nume)
    for nume in ("ttk", "filedialog", "messagebox"):
        setattr(sys.modules["tkinter"], nume, sys.modules["tkinter." + nume])

import generator_oferte_bess as g          # noqa: E402
from date_test import date_test, intrari_test   # noqa: E402

JINKO_2H = "JKE-5015K-2H-LAA"
REPT = "REPT 5,015"


class Oferte(unittest.TestCase):
    """Proiecte-tip: ce baterie, câte bucăți, ce skid și ce trafo aux trebuie să iasă."""

    def setUp(self):
        self.date = date_test()
        self.baterii = g.lista_baterii(self.date["bat"])

    def oferta(self, baterie, mwh, mw, brand="Oricare", poc_kv=20, nr_fix=None, trafo=None):
        return g.pregateste_oferta(self.date, self.baterii, intrari_test(mwh, mw, poc_kv), baterie, "DDP",
                                   brand, nr_fix, trafo)

    def test_cotatia_cea_mai_noua(self):
        b = g.alege_baterie(self.date["bat"], self.baterii[JINKO_2H], "DDP")
        self.assertEqual(b["pret"], 85000)                 # cotația din 09.2026, nu cea din 01.2026
        self.assertEqual(g.alege_baterie(self.date["bat"], self.baterii[JINKO_2H], "DAP")["pret"], 80000)

    def test_20mwh_10mw_jinko(self):
        p = self.oferta(JINKO_2H, 20, 10)
        o = p["lista"][0]
        self.assertEqual(p["nr_baterii"], 4)
        self.assertEqual((o["tip"], o["nr"]), ("BCS5000K-B-HUD MV", 2))
        # aux: 4 baterii x 55 + 2 skiduri x 40 kVA (5 MW, neconfirmat) = 300 kVA -> 1 trafo de 450 kVA
        self.assertEqual(o["trafo"]["kva_total"], 300)
        self.assertEqual((o["trafo"]["kva"], o["trafo"]["cant"]), (450, 1))

    def test_kehua_7500_3_baterii(self):
        # 6 PCS x 1250 kW: 3 baterii (1 baterie : 2 PCS)
        p = self.oferta(JINKO_2H, 15, 7.5, brand="Kehua")
        o = p["lista"][0]
        self.assertEqual(p["nr_baterii"], 3)
        self.assertEqual((o["tip"], o["nr"]), ("BCS7500K-B-HUD MV", 1))

    def test_tbea_are_aux_inclus(self):
        p = self.oferta(REPT, 20, 10, brand="TBEA")     # 4 baterii x 12 rack-uri = 48 -> 24 intrări (1:2)
        o = p["lista"][0]
        self.assertEqual((o["brand"], o["nr"]), ("TBEA", 1))
        self.assertEqual(o["putere_unit"], 10)               # puterea nominală din 'Skid power'
        self.assertEqual(o["aux_skid"], 200)

    def test_tbea_doar_la_20kv(self):
        with self.assertRaisesRegex(ValueError, "doar pentru POC 20 kV"):
            self.oferta(JINKO_2H, 20, 10, brand="TBEA", poc_kv=33)

    def test_alt_poc_foloseste_foaia_poc(self):
        o = self.oferta(JINKO_2H, 20, 10, poc_kv=33)["lista"][0]
        self.assertIn("33kV", o["tip"])
        self.assertNotEqual(o["brand"], "TBEA")

    def test_rept_merge_si_pe_kehua_la_alt_poc(self):
        o = self.oferta(REPT, 40, 10, poc_kv=33)["lista"][0]
        self.assertEqual(o["brand"], "Kehua")

    def test_auto_exclude_bateriile_4h_la_proiect_2h(self):
        p = self.oferta(g.OPTIUNE_AUTO_BATERIE, 20, 10)
        self.assertEqual(p["baterie"]["tip"], JINKO_2H)
        excluse = " ".join(p["auto_excluse"])
        self.assertIn(REPT, excluse)                         # REPT = 4 ore

    def test_lista_baterii_2h_si_rack_uri_scoase(self):
        lista = g.baterii_din_lista(self.baterii)
        self.assertIn("JKE-4597K-2H-LAA", lista)            # varianta cu rack-uri scoase
        self.assertIn(REPT, lista)
        self.assertNotIn("JKE-5015K-4H-LAA", lista)          # 4H nu mai apare (cele de 2H merg și la 4 ore)

    def test_baterie_cu_rack_uri_scoase(self):
        p = self.oferta("JKE-4597K-2H-LAA", 18, 10)
        self.assertEqual(p["nr_baterii"], 4)                 # 4 x 4,597 = 18,388 MWh

    def test_auto_alege_cea_mai_ieftina(self):
        p = self.oferta(g.OPTIUNE_AUTO_BATERIE, 40, 10)
        self.assertEqual(p["baterie"]["tip"], REPT)

    def test_ems_dupa_capacitatea_ofertata(self):
        # cerut 13,5 MWh -> 3 x 5,015 = 15,045 MWh >= 15 -> EMS de la prag (3.500), nu cel sub prag (6.500)
        intrari = dict(intrari_test(13.5, 5), ems=g.ems_pentru(13.5), ems_auto=True)
        p = g.pregateste_oferta(self.date, self.baterii, intrari, JINKO_2H, "DDP", "Oricare", None, None)
        self.assertEqual(intrari["ems"], 6500)
        self.assertEqual(g.ajusteaza_ems(intrari, p), 3500)
        self.assertEqual(intrari["ems"], 3500)

    def test_ems_scris_de_mana_nu_se_schimba(self):
        intrari = dict(intrari_test(13.5, 5), ems=5000, ems_auto=False)
        p = g.pregateste_oferta(self.date, self.baterii, intrari, JINKO_2H, "DDP", "Oricare", None, None)
        self.assertIsNone(g.ajusteaza_ems(intrari, p))
        self.assertEqual(intrari["ems"], 5000)

    def test_fara_trafo_aux(self):
        o = self.oferta(JINKO_2H, 20, 10, trafo="fara")["lista"][0]
        self.assertFalse(o.get("trafo"))


class TrafoAux(unittest.TestCase):
    """Consumul aux: baterii + skiduri (după putere); trafo extern doar unde skidul nu acoperă."""

    def setUp(self):
        self.date = date_test()
        self.baterii = g.lista_baterii(self.date["bat"])

    def skid(self, brand, mwh=20, mw=10, trafo=None, baterie=JINKO_2H):
        p = g.pregateste_oferta(self.date, self.baterii, intrari_test(mwh, mw), baterie, "DDP", brand, None, trafo)
        return p["lista"][0]

    def test_trepte_consum_skid(self):
        self.assertEqual([g.consum_aux_skid(mw) for mw in (1, 2, 3, 5, 7.5, 10, 12.5)],
                         [30, 30, 40, 40, 50, 60, 75])

    def test_fara_trafo_in_skid(self):
        o = self.skid("Kehua")                               # 4 x 55 + 2 x 40 = 300 kVA, skidul nu are aux
        self.assertEqual((o["kva_necesar"], o["kva_acoperit"], o["kva_trafo"]), (300, 0, 300))
        self.assertEqual(o["trafo"]["kva"] * o["trafo"]["cant"], 450)

    def test_trafo_in_skid_acopera_tot_sistemul(self):
        o = self.skid("Sinexcel")                            # 2 skiduri x 400 kVA >= 300 kVA
        self.assertEqual((o["kva_necesar"], o["kva_acoperit"], o["kva_trafo"]), (300, 800, 0))
        self.assertIsNone(o["trafo"])
        self.assertEqual(o["cost_total"], o["pret_total"])  # fără cost de trafo

    def test_trafo_in_skid_nu_ajunge(self):
        # TBEA: 1 skid cu 200 kVA; 4 baterii REPT x 55 + 60 kVA (10 MW) = 280 -> lipsesc 80 kVA
        o = self.skid("TBEA", baterie=REPT)
        self.assertEqual((o["kva_necesar"], o["kva_acoperit"], o["kva_lipsa"]), (280, 200, 80))
        self.assertEqual(o["kva_trafo"], 80)
        self.assertEqual(o["trafo"]["kva"], 450)

    def test_un_trafo_mare_in_loc_de_mai_multe_mici(self):
        variante = g.variante_trafo(self.date["aux"])           # 450 kVA = 12.000 $, 1250 kVA = 25.000 $
        ales = g.alege_trafo(variante, 1100)
        self.assertEqual((ales["kva"], ales["cant"]), (1250, 1))  # nu 3 x 450 (36.000 $... sau mai ieftin)
        ales = g.alege_trafo(variante, 1800)
        self.assertEqual((ales["kva"], ales["cant"]), (1250, 2))  # nu 4 x 450
        ales = g.alege_trafo(variante, 300)
        self.assertEqual((ales["kva"], ales["cant"]), (450, 1))   # la o bucată: cel mai ieftin care acoperă

    def test_trafo_cel_mai_ieftin_daca_setarea_e_nu(self):
        vechi = g.TRAFO_CAT_MAI_PUTINE
        g.TRAFO_CAT_MAI_PUTINE = False
        try:
            self.date["aux"].loc[:, "Price USD DAP"] = [5000, 25000]
            ales = g.alege_trafo(g.variante_trafo(self.date["aux"]), 1100)
            self.assertEqual((ales["kva"], ales["cant"]), (450, 3))   # 15.000 $ < 25.000 $
        finally:
            g.TRAFO_CAT_MAI_PUTINE = vechi

    def test_trafo_fortat_chiar_daca_skidul_acopera(self):
        o = self.skid("Sinexcel", trafo=True)
        self.assertEqual(o["kva_trafo"], 300)
        self.assertIsNotNone(o["trafo"])

    def test_fara_trafo_extern_bifat(self):
        o = self.skid("Kehua", trafo="fara")
        self.assertIsNone(o["trafo"])
        self.assertEqual(o["kva_lipsa"], 300)               # rămâne avertismentul cu kVA neacoperiți
        self.assertIn("neacoperiți 300", g.text_aux(o))

    def test_auto_alege_skidul_cu_trafo_inclus_cand_iese_mai_ieftin(self):
        # Kehua 200.000 x 2 + trafo 12.000 = 412.000; Sinexcel 210.000 x 2 fără trafo = 420.000
        self.assertEqual(self.skid("Oricare")["brand"], "Kehua")
        self.date["aux"].loc[:, "Price USD DAP"] = [30000, 50000]       # trafo extern mai scump
        self.assertEqual(self.skid("Oricare")["brand"], "Sinexcel")

    def test_skid_manual_7_5_mw(self):
        m = {"brand": "X", "model": "Y", "putere": "7,5", "pret": "100000", "aux_kva": ""}
        o = g.optiune_skid_manuala(self.date, 15, 3, 55, None, None, m)
        self.assertEqual(o["kva_necesar"], 3 * 55 + 2 * 50)


class ComponenteManuale(unittest.TestCase):
    def test_valori_manuale(self):
        m = {"a": "75.000", "b": "1.250,50", "c": "5.015", "d": "5,015", "e": ""}
        self.assertEqual(g._valoare_manuala(m, "a", "", bani=True), 75000)
        self.assertEqual(g._valoare_manuala(m, "b", "", bani=True), 1250.5)
        self.assertEqual(g._valoare_manuala(m, "c", ""), 5.015)
        self.assertEqual(g._valoare_manuala(m, "d", ""), 5.015)
        self.assertEqual(g._valoare_manuala(m, "e", "", 7), 7)

    def test_procent_extra_warranty(self):
        self.assertEqual([g._procent_manual({"ext_gar": v}) for v in ("", "3", "3%", "2,5")], [0, 3, 3, 2.5])

    def test_preia_din_pricing(self):
        d = date_test()
        b = g.baterie_din_pricing(d, g.lista_baterii(d["bat"]), JINKO_2H)
        self.assertEqual((b["pret_ddp"], b["pret_dap"], b["garantie"]), (85000, 80000, "15"))
        s = g.skid_din_pricing(d, *g.skiduri_din_pricing(d)["TBEA | TBEA 10MW + AUX TRAFO 200kVA"])
        self.assertEqual((s["putere"], s["aux_kva"]), (10, 200))


class Formatare(unittest.TestCase):
    def test_bani_si_numere(self):
        self.assertEqual(g.bani_eu(1234567.891), "1.234.567,89")
        self.assertEqual(g.nr_eu(5.015), "5,015")

    def test_putere_limitata_prin_ems(self):
        self.assertEqual(g.text_putere_propusa(6, 5.85), "6 (limited by EMS to 5,85)")
        self.assertEqual(g.text_putere_propusa(10, 10), "10")
        self.assertEqual(g.text_putere_propusa(7.5, 7.5), "7,5")

    def test_ore_baterie(self):
        self.assertEqual(g.ore_baterie(JINKO_2H), 2)
        self.assertEqual(g.ore_baterie(REPT), 4)


class VerificarePricing(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self._setari = g.FISIER_SETARI
        g.FISIER_SETARI = os.path.join(self.tmp.name, "setari.json")

    def tearDown(self):
        g.FISIER_SETARI = self._setari
        self.tmp.cleanup()

    def test_pricing_corect(self):
        probleme, schimbate, actuale = g.verifica_coloane_pricing(date_test())
        self.assertEqual((probleme, schimbate), ([], []))

    def test_coloana_inserata_e_oprita(self):
        d = date_test()
        d["bat"].insert(5, "Coloană nouă", ["text"] * len(d["bat"]))
        self.assertTrue(g.verifica_coloane_pricing(d)[0])

    def test_titlu_redenumit_cere_confirmare(self):
        g.accepta_antete_pricing(g.verifica_coloane_pricing(date_test())[2])
        d = date_test()
        d["skid"] = d["skid"].rename(columns={"Price": "Price USD"})
        probleme, schimbate, _ = g.verifica_coloane_pricing(d)
        self.assertEqual(probleme, [])
        self.assertEqual(len(schimbate), 1)


class SetariGenerator(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.cale = os.path.join(self.tmp.name, "Setari generator.xlsx")
        self.valori = {n: getattr(g, n) for n, _, t in g.SETARI_EDITABILE if t}

    def tearDown(self):
        for n, v in self.valori.items():
            setattr(g, n, v)
        self.tmp.cleanup()

    def modifica(self, schimbari):
        from openpyxl import load_workbook
        wb = load_workbook(self.cale)
        for rand in wb[g.FOAIE_REGULI].iter_rows(min_row=2):
            if rand[0].value in schimbari:
                rand[2].value = schimbari[rand[0].value]
        wb.save(self.cale)

    def test_fisier_nou_nu_schimba_nimic(self):
        g.scrie_setari_generator(self.cale)
        self.assertEqual(g.aplica_setari_generator(self.cale), (0, []))

    def test_setari_aplicate(self):
        g.scrie_setari_generator(self.cale)
        self.modifica({"ADAOS_IMPLICIT": 1.15, "KVA_AUX_PER_BATERIE": "0 = 55; 5,6 = 70",
                       "BATERII_IN_LISTA": "5015, 6250"})
        self.assertEqual(g.aplica_setari_generator(self.cale), (3, []))
        self.assertEqual(g.ADAOS_IMPLICIT, 1.15)
        self.assertEqual(g.KVA_AUX_PER_BATERIE, [(0, 55), (5.6, 70)])
        self.assertEqual(g.BATERII_IN_LISTA, ("5015", "6250"))

    def test_greseala_nu_aplica_nimic(self):
        g.scrie_setari_generator(self.cale)
        self.modifica({"ADAOS_IMPLICIT": 5})
        schimbate, probleme = g.aplica_setari_generator(self.cale)
        self.assertEqual(schimbate, 0)
        self.assertTrue(probleme)
        self.assertEqual(g.ADAOS_IMPLICIT, self.valori["ADAOS_IMPLICIT"])


class FoldereZile(unittest.TestCase):
    def test_versiunea_tine_cont_si_de_folderele_pe_zile(self):
        with tempfile.TemporaryDirectory() as tmp:
            vechi = os.getcwd()
            os.chdir(tmp)
            try:
                os.makedirs(os.path.join(g.FOLDER_OFERTE, "2026-11-05",
                                         "ST1_BESS_Indicative_Proposal_Taliva_P_20MWh_10MW_v1"))
                os.makedirs(os.path.join(g.FOLDER_OFERTE, "ST1_BESS_Indicative_Proposal_Taliva_P_20MWh_10MW_v2"))
                self.assertEqual(g.urmatoarea_versiune("ST1", "P"), 3)
                self.assertEqual(g.urmatoarea_versiune("ST1", "P", ["ST1_BESS_Indicative_Proposal_Taliva_P_5MWh_v7"]),
                                 8)
            finally:
                os.chdir(vechi)
        self.assertTrue(g.MODEL_FOLDER_ZI.match("2026-11-06"))


class FoiInSetari(unittest.TestCase):
    """PCS și specificațiile ca foi în 'Setari generator.xlsx' (formulele calculate de program)."""

    def test_pcs_cu_formule(self):
        from openpyxl import Workbook
        with tempfile.TemporaryDirectory() as tmp:
            cale = os.path.join(tmp, "Setari generator.xlsx")
            wb = Workbook()
            sh = wb.active
            sh.title = g.FOAIE_PCS_SETARI
            sh.append(["SKID", "PCS power", "no / frame", "Skid power"])
            sh.append(["SES4H-5000-EX-MV", 1.25, 4, "=B2*C2"])
            sh.append(["Frame 4 : MV Skid 5150kVA", "=D3/4", "4 frame", 5.15])
            sh.append(["Gresit", "=SUM(B2:B3)", 1, "=B4*2"])
            wb.save(cale)
            self.assertIn(g.FOAIE_PCS_SETARI, g.foi_setari(cale))
            df = g.citeste_foaie_cu_formule(cale, g.FOAIE_PCS_SETARI)
        self.assertEqual(df["Skid power"].tolist()[:2], [5.0, 5.15])
        self.assertAlmostEqual(df["PCS power"].tolist()[1], 1.2875)
        self.assertTrue(df.isna().iloc[2]["PCS power"])      # formulă necunoscută -> gol, fără eroare


class Microsoft365(unittest.TestCase):
    """Calculul prin Excel Online și PDF-ul prin Drive (fără Excel / PowerPoint instalate), cu Drive simulat."""

    def setUp(self):
        import json
        self.json = json
        self.cereri = []
        self.batch_esuat = 0
        d = g.DriveCloud.__new__(g.DriveCloud)
        d.ofertare = ("D", "ROOT")
        d._cerere = self.cerere
        self.drive = d

    def cerere(self, metoda, cale, date=None, tip=None, brut=False, **k):
        corp = self.json.loads(date) if date else None
        self.cereri.append((metoda, cale, corp))
        if cale == "/$batch":
            stare = 503 if self.batch_esuat else 200
            self.batch_esuat = max(0, self.batch_esuat - 1)
            return {"responses": [{"id": r["id"], "status": stare} for r in corp["requests"]]}
        if "usedRange" in cale:
            return {"address": "Sheet1!A10:B11", "formulas": [["Commisioning Sinexcel", "=A1"], ["x", ""]]}
        if "range(address='A4:I14')" in cale:
            return {"values": [["Product", "", "", "", "", "", "", "", ""]] * 11}
        if "range(address=" in cale:
            return {"values": [[100.0]]}
        if cale.endswith("format=pdf"):
            return b"%PDF-1.7 test"
        return {}

    def scrieri(self):
        return [r for c in self.cereri if c[1] == "/$batch" for r in c[2]["requests"]]

    def oferta(self):
        date = date_test()
        p = g.pregateste_oferta(date, g.lista_baterii(date["bat"]), intrari_test(), JINKO_2H, "DDP", "Oricare",
                                None, None)
        return p, p["lista"][0]

    def test_completare_prin_excel_online(self):
        from datetime import datetime
        p, o = self.oferta()
        rez = g.completeaza_calcul(g.FoaieGraph(self.drive, ("D", "F")), dict(intrari_test(), versiune=1),
                                   p["baterie"], p["nr_baterii"], o, o["trafo"], datetime(2026, 11, 9, 14, 0),
                                   ["Jinko"], ["Kehua", "Sinexcel"])
        self.assertEqual(rez["pret_ddp"], 100.0)
        scrieri = {r["url"].split("address='")[1].split("'")[0]: r["body"] for r in self.scrieri()}
        self.assertEqual(scrieri["A5:D5"]["values"], [[JINKO_2H, 5.015, 4, 85000.0]])
        self.assertEqual(scrieri["G12"], {"formulas": [["=8000.0*C6"]]})
        self.assertAlmostEqual(scrieri["D1:E3"]["values"][0][1], 46335.5833, places=3)   # data ca număr Excel
        self.assertEqual(scrieri["A10"]["values"], [["Commisioning Kehua"]])           # brandul skidului ales
        self.assertTrue(all(len(c[2]["requests"]) <= 20 for c in self.cereri if c[1] == "/$batch"))
        self.assertIn(("POST", "/drives/D/items/F/workbook/application/calculate"),
                      [c[:2] for c in self.cereri])

    def test_scrierile_merg_in_ordine(self):
        foaie = g.FoaieGraph(self.drive, ("D", "F"))
        for i in range(25):
            foaie.scrie(f"A{i + 1}", i)
        foaie.trimite()
        loturi = [c[2]["requests"] for c in self.cereri if c[1] == "/$batch"]
        self.assertEqual([len(x) for x in loturi], [20, 5])
        self.assertEqual(loturi[0][5]["dependsOn"], ["5"])

    def test_fisier_ocupat_se_reincearca(self):
        g_sleep = g.time.sleep
        g.time.sleep = lambda s: None
        try:
            self.batch_esuat = 2                       # primele două încercări: 503 (fișier ocupat)
            foaie = g.FoaieGraph(self.drive, ("D", "F"))
            foaie.scrie("A1", 1)
            foaie.trimite()
        finally:
            g.time.sleep = g_sleep
        self.assertEqual(sum(c[1] == "/$batch" for c in self.cereri), 3)

    def test_golire_celula(self):
        foaie = g.FoaieGraph(self.drive, ("D", "F"))
        foaie.scrie("B7", None)
        foaie.trimite()
        self.assertEqual(self.scrieri()[0]["body"], {"values": [[""]]})

    def test_pdf_prin_drive(self):
        vechi = g.DRIVE
        g.DRIVE = self.drive
        self.drive.folder_oferta = lambda folder: "FOLDER"
        self.drive.urca_fisier = lambda id_folder, nume, cale: ("D", "PPTX")
        try:
            with tempfile.TemporaryDirectory() as tmp:
                pdf = os.path.join(tmp, "o.pdf")
                g.pptx_in_pdf_online(os.path.join(tmp, "o.pptx"), pdf, tmp)
                with open(pdf, "rb") as f:
                    self.assertEqual(f.read(), b"%PDF-1.7 test")
        finally:
            g.DRIVE = vechi
        self.assertEqual(self.cereri[-1][1], "/drives/D/items/PPTX/content?format=pdf")


class CalculDeProba(unittest.TestCase):
    def test_client_si_proiect_test(self):
        self.assertTrue(g.e_calcul_test({"client": "Test", "proiect": " test "}))
        self.assertFalse(g.e_calcul_test({"client": "test", "proiect": "Proiect X"}))
        self.assertFalse(g.e_calcul_test({"client": "GEP", "proiect": "test"}))

    def test_folderul_de_teste_nu_se_muta_pe_zile(self):
        import json
        cereri = []
        d = g.DriveCloud.__new__(g.DriveCloud)
        d.ofertare = ("D", "ROOT")
        d._cerere = lambda m, c, date=None, **k: cereri.append((m, c)) or {"id": "OFERTE"}
        d._copii = lambda f, sub="": [{"name": "_teste", "folder": {}, "id": "T",
                                       "createdDateTime": "2026-11-09T10:00:00Z"}]
        self.assertEqual(d.organizeaza_oferte_pe_zile(), 0)
        self.assertFalse(any(m == "PATCH" for m, _ in cereri))


ANTET = "Distribuție Energie Electrică Romania\nStr. Exemplu, Nr. 1, Cluj-Napoca\nTel: +40 000 000 000\n"
ATR_TEST = (ANTET + """AVIZ TEHNIC DE RACORDARE
nr. 1234567890/data 01.04.2026
Ca urmare a cererii înregistrate cu nr. 1234567890 din data 26.03.2026, având ca scop Modificarea instalatiei
adresată de FIRMA TEST SRL, pentru CEM 4,1 MW TEST ce aparţine utilizatorului FIRMA TEST SRL
amplasat(ă) în judeţul Bihor, - LUNCA, sat X, in conditiile mentionate in continuare.
d) Instalatie de stocare:
Tabelul 1
Nr. crt. Tip IS* Pi IS (kW) Pmax evac IS (kW) Pmax abs IS (kW) Capacitate (Ah) Observaţii
1 2 3 4 5 6 7
1 li-ion 4800,000 4100,000 0,000 13440,00
   * Instalaţie de stocare de tip electric (baterie Li-Ion), termic, cinetic.
3 / 14
""" + ANTET + """2. Puterea aprobata:
Etapa finală
(kW) 4100,000 4100,000 4100,000Puterea maximă simultană ce poate fi
""" + ANTET + """evacuată (kVA) 4550,000 4550,000 4550,000
(kW)
Puterea maximă simultană ce poate fi evacuată fără realizarea lucrărilor de întărire
(kVA)
(kW) 12,000 12,000 12,000Puterea maximă simultană ce poate fi absorbită din reţea (kVA) 13,333 13,333 13,333
3.Descrierea succintă a soluţiei de racordare
a) Punctul de racordare este stabilit la nivelul de tensiune 20 kV, la LEA 20 kV TEST
* trafo 20/0,69 kV - 1600 kVA;
* trafo servicii auxiliare 0,69/0,4 kV - 16 kVA;
* trafo 20/0,69 kV - 1600 kVA;
* trafo 20/0,69 kV - 1600 kVA;
c') In instalatia de utilizare se va monta o instalatie de stocare cu Pi = 4,1 MW.
Cuplarea baterie de stocare se face pe partea de curent continuu, cuplare DC/DC.
IS nou instalata nu va absorbi putere activa din RED.
""" + ANTET) * 1


class CitireATR(unittest.TestCase):
    def test_atr_model_anre(self):
        r = g.citeste_atr_text(ATR_TEST + " cuvânt" * 60)
        self.assertEqual((r["nr"], r["data"], r["beneficiar"]), ("1234567890", "01.04.2026", "FIRMA TEST SRL"))
        self.assertTrue(r["operator"].startswith("Distribuție Energie Electrică"))
        self.assertEqual(r["poc_kv"], 20)
        self.assertEqual((r["evac_kw"], r["evac_kva"], r["abs_kw"]), (4100, 4550, 12))   # cu antet între rânduri
        self.assertEqual((r["pi_is_kw"], r["evac_is_kw"], r["abs_is_kw"]), (4800, 4100, 0))
        self.assertEqual([k for k, _ in r["trafo_kva"]], [1600, 1600, 1600])      # fără trafo-ul de servicii
        self.assertEqual((r["mw"], r["mw_max"]), (4.1, 4.8))
        self.assertTrue(r["cuplare_dc"] and r["fara_absorbtie"])
        av = " ".join(r["avertismente"])
        self.assertIn("4100 kW în text", av)                 # Pi diferit în tabel față de text
        self.assertIn("cuplarea stocării pe DC", av)
        self.assertIn("MWh", av)

    def test_kwh_in_coloana_de_kw(self):
        text = ATR_TEST.replace("4800,000 4100,000", "4179,900 996,000").replace("4100,000", "996,000")
        r = g.citeste_atr_text(text + " cuvânt" * 60)
        self.assertIsNone(r.get("mw_max") if r.get("pi_text_kw") is None else None)
        self.assertIn("kWh", " ".join(r["avertismente"]))

    def test_studiu_de_solutie(self):
        text = ("Actualizare Studiu de solutie privind Racordarea la SEN avand avizul tehnic de racordare nr. 18189 "
                "din 12.12.2024 Rețele Electrice Dobrogea. punctul de racord se stabiliteste la bara de 20 kV din statie. "
                "Puterea activă instalată totală MW S, D, R 7,7 Puterea activă maximă evacuată MW S, D, R 7 "
                "Puterea activă maximă absorbită MW S, D, R 0 Capacitatea maximă totală stocată de IS MWh S, D, R 15,4 "
                "cu încărcare exclusiv din CEF " + "text " * 60)
        r = g.citeste_atr_text(text)
        self.assertEqual((r["tip"], r["nr"], r["poc_kv"]), ("Studiu de soluție", "18189", 20))
        self.assertEqual((r["mw"], r["mw_max"], r["mwh"]), (7.0, 7.7, 15.4))
        self.assertTrue(r["fara_absorbtie"])

    def test_document_scanat(self):
        r = g.citeste_atr_text("   ")
        self.assertTrue(r["scanat"])

    def test_viteza(self):
        import time
        t0 = time.time()
        g.citeste_atr_text((ATR_TEST + " (kW) " + "1,0 " * 50 + "\n") * 40)
        self.assertLess(time.time() - t0, 5)                  # fără blocaje pe documente mari


class RezervaSiPutereMaxima(unittest.TestCase):
    def setUp(self):
        self.date = date_test()
        self.baterii = g.lista_baterii(self.date["bat"])

    def oferta(self, **extra):
        intrari = dict(intrari_test(20, 10), **extra)
        return g.pregateste_oferta(self.date, self.baterii, intrari, JINKO_2H, "DDP", "Oricare", None, None)

    def test_rezerva_10_la_suta(self):
        fara, cu = self.oferta(), self.oferta(rezerva_bat=10, rezerva_skid=10)
        self.assertAlmostEqual(cu["baterie"]["pret"], fara["baterie"]["pret"] * 1.1)
        self.assertEqual(cu["baterie"]["pret_baza"], fara["baterie"]["pret"])
        o_fara, o_cu = fara["lista"][0], cu["lista"][0]
        self.assertAlmostEqual(o_cu["pret_unit"], o_fara["pret_unit"] * 1.1)
        self.assertAlmostEqual(o_cu["cost_total"], o_cu["pret_total"] + o_cu["trafo"]["cost"])

    def test_rezerva_in_excel_ca_formula(self):
        p = self.oferta(rezerva_bat=10, rezerva_skid=5)
        o = p["lista"][0]

        class Foaie:
            def __init__(s): s.scrieri = {}
            def scrie(s, a, v): s.scrieri[a] = v
            def format_numar(s, *a): pass
            def ingroasa(s, *a): pass
            def formule(s): return 1, 1, []
            def calculeaza(s): pass
            def citeste(s, a): return 1
        f = Foaie()
        from datetime import datetime
        g.completeaza_calcul(f, dict(intrari_test(), versiune=1), p["baterie"], p["nr_baterii"], o, o["trafo"],
                             datetime(2026, 11, 10))
        self.assertEqual(f.scrieri["A5:D5"][0][3], "=85000.0*(1+10%)")
        self.assertTrue(f.scrieri["A6:D6"][0][3].endswith("*(1+5%)"))
        self.assertIn("Rezervă preț", f.scrieri["B3"])

    def test_putere_maxima_admisa(self):
        # 10 MW cerut: fără limită iese 2 x 5 MW; cu max. 9 MW nu mai e niciun skid valid
        self.assertEqual(self.oferta()["lista"][0]["putere_totala"], 10)
        with self.assertRaisesRegex(ValueError, "puterea maximă admisă"):
            self.oferta(mw_max=9)
        self.assertEqual(self.oferta(mw_max=10)["lista"][0]["putere_totala"], 10)


class Actualizare(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self._folder, self._descarca = g.FOLDER_PROGRAM, g._descarca_text
        g.FOLDER_PROGRAM = self.tmp.name
        self.github = {"generator_oferte_bess.py": 'FISIERE_PROGRAM = ("generator_oferte_bess.py", "modul.py")\nX = 2\n',
                       "modul.py": "Y = 1\n"}
        g._descarca_text = lambda url, timeout=15: self.github[url.rsplit("/", 1)[1]]
        with open(os.path.join(self.tmp.name, "generator_oferte_bess.py"), "w", encoding="utf-8", newline="\r\n") as f:
            f.write('FISIERE_PROGRAM = ("generator_oferte_bess.py",)\nX = 1\n')

    def tearDown(self):
        g.FOLDER_PROGRAM, g._descarca_text = self._folder, self._descarca
        self.tmp.cleanup()

    def test_actualizare_si_la_zi(self):
        noi = g.cauta_actualizare()
        self.assertEqual(set(noi), {"generator_oferte_bess.py", "modul.py"})
        g.aplica_actualizare(noi)
        self.assertTrue(os.path.exists(os.path.join(self.tmp.name, "generator_oferte_bess.py.bak")))
        self.assertEqual(g.cauta_actualizare(), {})

    def test_cod_gresit_nu_se_instaleaza(self):
        with self.assertRaises(SyntaxError):
            g.aplica_actualizare({"generator_oferte_bess.py": "def ("})
        with open(os.path.join(self.tmp.name, "generator_oferte_bess.py"), encoding="utf-8") as f:
            self.assertIn("X = 1", f.read())

    def test_fara_internet(self):
        def eroare(url, timeout=15):
            raise OSError("fără internet")
        g._descarca_text = eroare
        self.assertEqual(g.cauta_actualizare(), {})


if __name__ == "__main__":
    unittest.main()
