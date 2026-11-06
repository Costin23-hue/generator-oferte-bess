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
        self.assertIn("JKE-5015K-4H-LAA", excluse)

    def test_auto_alege_cea_mai_ieftina(self):
        p = self.oferta(g.OPTIUNE_AUTO_BATERIE, 40, 10)
        self.assertEqual(p["baterie"]["tip"], REPT)

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
