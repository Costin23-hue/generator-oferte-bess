import base64
import hashlib
import json
import math
import os
import re
import secrets
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import tkinter as tk
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
from datetime import datetime
from http.server import BaseHTTPRequestHandler, HTTPServer
from tkinter import ttk, messagebox, filedialog

import pandas as pd
try:                                    # Excel instalat (Windows / Mac); fără el se lucrează prin Microsoft 365
    import xlwings as xw
except ImportError:
    xw = None
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
# Ofertele se pun pe zile: 'Oferte generate/2026-11-06/<oferta>' (formatul an-lună-zi se sortează corect după nume)
FORMAT_FOLDER_ZI = "%Y-%m-%d"
MODEL_FOLDER_ZI = re.compile(r"^\d{4}-\d{2}-\d{2}$")
# Calcul de probă: Client = "test" și Nume Proiect = "test" -> nr. "TEST", nimic în BESS_Proposals, nimic păstrat în
# Drive; fișierele stau temporar în 'Oferte generate/_teste' și se șterg la Renunță / alt calcul / închidere.
TEXT_TEST = "test"
FOLDER_TESTE = "_teste"
NR_PROIECT_TEST = "TEST"

# Actualizare automată (doar când rulează fișierul .py, nu .exe): la pornire se compară fișierele programului cu
# cele de pe GitHub (repository public); dacă diferă, apare butonul "Versiune nouă – Actualizează".
URL_ACTUALIZARE = "https://raw.githubusercontent.com/Costin23-hue/generator-oferte-bess/main/"
FISIERE_PROGRAM = ("generator_oferte_bess.py",)

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
# PCS-urile și specificațiile pot sta și ca foi în 'Setari generator.xlsx' (au prioritate față de fișierele separate)
FOAIE_PCS_SETARI = "PCS - MV skid"
FOAIE_SPECS_SETARI = "Specificatii baterii"

# Lucrul direct cu Drive-ul (Microsoft Graph), fără foldere sincronizate: fiecare utilizator se loghează
# o dată cu contul de firmă; fișierele se descarcă din cele două foldere, iar ofertele se urcă în
# 'Ofertare BESS/Oferte generate'. Aplicația 'Generator Oferte BESS' e înregistrată în Entra ID
# (permisiuni delegate: vede doar ce vede utilizatorul logat). DRIVE_CLIENT_ID gol = doar foldere sincronizate.
DRIVE_CLIENT_ID = "3adacc30-a1fa-4556-bda2-2c0a0bd415e7"
DRIVE_TENANT_ID = "0ccbdc57-de92-4033-bd5d-a02a0032aee0"
LINK_FOLDER_PRICING = ("https://talivaenerji-my.sharepoint.com/:f:/r/personal/levon_solartoday_com_tr/Documents/"
                       "Purchasing%20Prices%20Update_BESS?d=wfb3649798b8c4fe9a0b73c1983d9a28a&csf=1&web=1&e=r1gLmf")
LINK_FOLDER_OFERTARE = ("https://talivaenerji-my.sharepoint.com/:f:/g/personal/costin_ispas_solartoday_ro/"
                        "IgDmNADThuauS7z4mdYKXlWHAQ_0slDB91z6qOyWZudqkCg?e=CL9KcT")
DRIVE_SCOPURI = "offline_access User.Read Files.ReadWrite.All Sites.Read.All"
# Aplicația a fost aprobată de administratorul Microsoft 365 (29.10.2026). Cu DRIVE_ACTIV = False programul
# lucrează doar cu folderele sincronizate (Pricing.xlsx e copiat oricum în 'Ofertare BESS' de fluxul Power Automate).
DRIVE_ACTIV = True

# Tabelul cu ofertele (BESS_Proposals.xlsx, în Drive): nr. de contract se ia de acolo (primul rând cu nr. ST și fără
# client, în foaia țării alese), iar după verificarea ofertei, butonul "Trece în BESS_Proposals" completează rândul.
# Generarea singură NU scrie nimic în tabel. Coloanele: litera -> câmpul din ofertă.
LINK_BESS_PROPOSALS = ("https://talivaenerji-my.sharepoint.com/:x:/g/personal/levon_solartoday_com_tr/"
                       "IQCrDwTTcwp6QpyYu2f5O3EkAd-x7ecf-4rXK6cg7lbUOzI?e=a2a6WF")
TABELE_OFERTE = {
    "Romania": {"foaie": "2026-Romania", "rand_antet": 8, "col_nr": "X", "col_client": "D", "col_libere": "CDE",
                "coloane": {"C": "data_cerere", "D": "client", "E": "proiect", "F": "locatie", "G": "tara",
                            "H": "tip", "I": "responsabil", "J": "brand_skid", "K": "mw", "L": "mwh",
                            "M": "brand_baterie", "N": "pret_total", "O": "pret_mwh", "P": "po", "Y": "data_oferta"}},
    "Moldova": {"foaie": "Moldova", "rand_antet": 1, "col_nr": "P", "col_client": "C", "col_libere": "ACD",
                "coloane": {"A": "data_cerere", "B": "data_oferta", "C": "client", "D": "proiect", "E": "locatie",
                            "F": "tip", "G": "responsabil", "H": "mw", "I": "mwh", "J": "c_rate",
                            "K": "mwh_container", "L": "pret_total", "M": "cost", "O": "po"}},
    "Poland": {"foaie": "Poland", "rand_antet": 1, "col_nr": "N", "col_client": "B", "col_libere": "ABC",
               "coloane": {"A": "data_cerere", "B": "client", "C": "proiect", "D": "locatie", "E": "tip",
                           "F": "responsabil", "G": "mw", "H": "mwh", "I": "mwh_container", "J": "pret_total",
                           "K": "pret_mwh", "L": "po"}},
}
TIPURI_PROIECT = ["Utility", "C&I"]
MAX_RANDURI_TABEL = 3000
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
# Bateriile afișate în listă: doar modelele care conțin unul dintre aceste texte. "2H-" = toate bateriile Jinko de 2 ore
# (5,015, 6,25 și variantele cu rack-uri scoase, ex. JKE-4597K-2H-LAA); cele de 2 ore merg și la proiecte de 4 ore,
# așa că variantele 4H nu mai apar. Plus REPT.
BATERII_IN_LISTA = ("2H-", "REPT")
# "Auto": se încearcă toate bateriile din listă (cu durata potrivită proiectului) și se alege oferta cea mai ieftină
OPTIUNE_AUTO_BATERIE = "Auto (cea mai ieftină)"
# Durata bateriilor care nu au '-2H-' / '-4H-' în nume
ORE_BATERIE_FARA_COD = {"REPT": 4}
CAPACITATE_BATERIE_FALLBACK = 5.015      # MWh, folosită doar dacă nu se găsește în Excel

# Câte intrări DC ocupă o baterie pe skid (limitele Min/Max din PCS - MV skid.xlsx sunt în intrări DC)
DC_INPUTURI_PER_BATERIE = 2
DC_INPUTURI_PER_BRAND = {"Sinexcel": 1}   # excepții pe brand de skid (Sinexcel: 1 intrare DC / baterie)
# Compatibilitate baterii <-> skid, după numărul de PCS (nr. skiduri x 'no / frame' din 'PCS - MV skid.xlsx',
# sau puterea skidului / 'PCS power' dacă 'no / frame' lipsește): o baterie alimentează 1 PCS sau 2 PCS
# (ex. 4 baterii pe 8 PCS = 1:2; 8 baterii pe 8 PCS = 1:1). Coloanele Min/Max intrări se folosesc doar
# când numărul de PCS nu se cunoaște și la REPT pe TBEA (conectare pe rack).
PCS_PE_BATERIE = (1, 2)

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
# Branduri la care puterea skidului e cea nominală din coloana 'Skid power' din 'PCS - MV skid.xlsx'
# (TBEA: 48 x 215 kW = 10,32 MW, dar în ofertă apare 10 MW), nu putere PCS x nr. PCS.
PUTERE_NOMINALA_SKID = ("TBEA",)

# Consum auxiliar / baterie, după capacitatea bateriei: (de la MWh, kVA), crescător.
# 5 MWh (Jinko, REPT) -> 55 kVA (technical proposal Jinko: "55 kVA per battery container"); 6,25 MWh -> 65 kVA.
# Tabelul se scrie și în Excelul de calcul (A24:B25).
KVA_AUX_PER_BATERIE = [(0, 55), (5.6, 65)]
# Consumul auxiliar propriu al unui skid (PCS + MV skid), kVA, după puterea skidului: (până la MW, kVA), crescător.
# Valori acoperitoare (cele mai mari), pentru orice skid. Între trepte: treapta următoare; peste ultima: proporțional.
KVA_CONSUM_SKID = [(2, 30), (5, 40), (7.5, 50), (10, 60)]
# Trafo aux inclus în skid (kVA): coloana cu acest nume din 'PCS - MV skid.xlsx'; dacă lipsește, din numele
# skidului ('... + AUX TRAFO 200kVA'); altfel skidul e considerat fără aux.
# Necesar = baterii x kVA/baterie + skiduri x consum aux skid; ce nu acoperă skidurile se acoperă automat cu
# trafo extern (cel mai ieftin din 'Auxiliary Transformer'), inclus în compararea prețurilor skidurilor.
COLOANA_AUX_PCS = "Aux trafo kVA"
# Trafo auxiliar: din foaia 'Auxiliary Transformer' se iau toate rândurile cu putere (kVA) în 'Detail' și preț
# (ex. '1250 KVA Aux transformer', '450kVA 690V/400V ... TRAFO'); prețul: DAP dacă există, altfel EXW.
# TRAFO_FILTRU = text obligatoriu în 'Detail' ("" = toate rândurile).
TRAFO_FILTRU = ""
# Alegerea trafo-ului extern: True = cât mai puține bucăți (ex. 1 x 1250 kVA, nu 4 x 450 kVA), iar la același număr
# de bucăți cel mai ieftin; False = combinația cea mai ieftină, indiferent de numărul de bucăți.
TRAFO_CAT_MAI_PUTINE = True

ORE_MIN, ORE_MAX = 0.5, 8                # raport MWh/MW considerat normal
ADAOS_IMPLICIT, CURS_IMPLICIT = 1.10, 0.89   # valorile puse în fereastră la pornire
# Limitare prin EMS: când puterea skidurilor e mai mare decât cea cerută (ex. 6 MW pentru 5,85 MW cerut),
# puterea evacuată se limitează prin EMS la cea cerută; în prezentare, la [PROP_MW], apare textul de mai jos.
# {propus} = puterea skidurilor, {limitat} = puterea cerută (ambele în stil european, ex. 5,85).
TEXT_LIMITARE_EMS = "{propus} (limited by EMS to {limitat})"
ADAOS_MIN, ADAOS_MAX = 1.0, 2.0
CURS_MIN, CURS_MAX = 0.5, 2.0
EMS_PRAG_MWH = 15          # de la această capacitate EMS se completează automat
EMS_AUTO = 3500            # valoarea EMS (USD) pusă automat la capacitate >= prag
EMS_SUB_PRAG = 6500        # valoarea pusă automat la capacitate sub prag (dacă nu ai modificat EMS manual)
NR_OPTIUNI_SKID = 5

# Punct de conexiune (POC): la 20 kV se folosește foaia "MV SKID" (România), ca până acum;
# la orice altă tensiune, foaia ale cărei nume conține "POC" ("MV SKID for different POC"),
# doar skidurile care au acea tensiune în nume (ex. "SES4H-5000-EX-MV 10KV SF6 switchgear").
POC_IMPLICIT_KV = 20
# Branduri de skid disponibile doar la 20 kV (România); la alt POC nu se iau în calcul
BRANDURI_DOAR_POC_IMPLICIT = ("TBEA",)
FOAIE_SKID = "MV SKID"
FOAIE_SKID_POC_CONTINE = "poc"

# Condiția de livrare: textul pus în prezentare la [INCOTERMS] și coloana de preț baterie folosită
INCOTERMS = {
    "DDP": "DDP Incoterms 2020",   # preț din coloana I
    "DAP": "DAP Incoterms 2020",   # preț din coloana H
}
INCOTERM_IMPLICIT = "DDP"
AUTO_SKID = "Auto (cel mai ieftin)"

# Componente introduse manual (tab-ul "Componente manuale"), pentru cotații care nu sunt (încă) în Pricing.
# Se pot combina liber cu cele din baza de date. (cheie, etichetă, valoare implicită); cheie None = titlu de grup,
# cheie '..._inclus' = bifă.
OPTIUNE_MANUALA_BATERIE = "✎ Baterie manuală (tab Componente manuale)"
OPTIUNE_MANUALA_SKID = "✎ Skid manual (tab Componente manuale)"
CAMPURI_BATERIE_MANUALA = [
    ("brand", "Brand / furnizor:", ""), ("model", "Model:", ""), ("cap", "Capacitate / baterie (MWh):", ""),
    ("pret_ddp", "Preț DDP ($/MWh):", ""), ("pret_dap", "Preț DAP ($/MWh):", ""),
    ("transport", "Transport / container ($):", ""), ("transport_inclus", "Prețul include transportul", None),
    ("garantie", "Garanție standard (ani):", ""), ("ext_gar", "Extra warranty (% din total, opțional):", ""),
    ("kva_aux", "Consum aux / baterie (kVA, gol = automat):", ""),
    (None, "Specificații pentru prezentare", None),
    ("model_cod", "Cod model (rândul Model):", ""), ("config", "Configuration:", ""), ("weight", "Weight (t):", ""),
]
CAMPURI_SKID_MANUAL = [
    ("brand", "Brand / furnizor:", ""), ("model", "Model:", ""), ("putere", "Putere / skid (MW):", ""),
    ("putere_pcs", "Putere PCS (MW):", ""), ("pret", "Preț / skid ($):", ""),
    ("transport", "Transport / skid ($):", ""), ("transport_inclus", "Prețul include transportul", None),
    ("com", "Commissioning / skid ($):", ""), ("aux_kva", "Trafo aux inclus în skid (kVA, gol = fără):", ""),
    ("garantie", "Garanție standard (ani):", ""), ("ext_gar", "Extra warranty (% din total, opțional):", ""),
]

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

# Fișierul cu regulile de mai sus, editabil fără cod: stă în folderul 'Ofertare BESS' (în Drive) și se citește la
# pornire. Ce lipsește din el sau e gol rămâne cu valoarea din cod. Butonul "⚙ Setări" din fereastră îl creează.
FISIER_SETARI_GENERATOR = "Setari generator.xlsx"
FOAIE_REGULI = "Reguli"
# (variabila din cod, ce este, tipul valorii). Tipuri: numar, intreg, text, lista (texte separate prin virgulă),
# lista_intregi, perechi ("REPT = 4; ..."), perechi_text ("TBEA = rack; ..."), praguri ("0 = 55; 5,6 = 65").
SETARI_EDITABILE = [
    ("Oferta", None, None),
    ("ADAOS_IMPLICIT", "Adaos comercial (margin) pus la pornire", "numar"),
    ("ADAOS_MIN", "Adaos minim acceptat", "numar"),
    ("ADAOS_MAX", "Adaos maxim acceptat", "numar"),
    ("CURS_IMPLICIT", "Curs USD/EUR pus la pornire", "numar"),
    ("CURS_MIN", "Curs USD/EUR minim acceptat", "numar"),
    ("CURS_MAX", "Curs USD/EUR maxim acceptat", "numar"),
    ("EMS_PRAG_MWH", "EMS: pragul de capacitate (MWh)", "numar"),
    ("EMS_AUTO", "EMS (USD) la capacitate >= prag", "numar"),
    ("EMS_SUB_PRAG", "EMS (USD) la capacitate sub prag", "numar"),
    ("INCOTERM_IMPLICIT", "Incoterm pus la pornire (DDP sau DAP)", "text"),
    ("TEXT_LIMITARE_EMS", "Text în prezentare la putere limitată prin EMS ({propus}, {limitat})", "text"),
    ("ORE_MIN", "Durată proiect considerată normală: minim (ore)", "numar"),
    ("ORE_MAX", "Durată proiect considerată normală: maxim (ore)", "numar"),
    ("NR_OPTIUNI_SKID", "Câte variante de skid se arată în previzualizare", "intreg"),
    ("Baterii", None, None),
    ("BATERII_IN_LISTA", "Bateriile din listă (text conținut în model)", "lista"),
    ("ORE_BATERIE_FARA_COD", "Durata bateriilor fără -2H- / -4H- în nume (brand = ore)", "perechi"),
    ("KVA_AUX_PER_BATERIE", "Consum aux / baterie (de la MWh = kVA)", "praguri"),
    ("FARA_EXTENSIE_GARANTIE", "Branduri fără extra warranty (rândul dispare din prezentare)", "lista"),
    ("Skiduri și compatibilitate", None, None),
    ("PCS_PE_BATERIE", "Câte PCS poate alimenta o baterie (ex. 1, 2)", "lista_intregi"),
    ("KVA_CONSUM_SKID", "Consum aux propriu / skid (până la MW = kVA)", "trepte"),
    ("TRAFO_FILTRU", "Trafo aux: text obligatoriu în 'Detail' (gol = toate)", "text"),
    ("TRAFO_CAT_MAI_PUTINE", "Trafo aux: cât mai puține bucăți (DA) sau cel mai ieftin (NU)", "da_nu"),
    ("POC_IMPLICIT_KV", "POC România (kV) = foaia 'MV SKID'", "numar"),
    ("BRANDURI_DOAR_POC_IMPLICIT", "Branduri de skid doar la POC România", "lista"),
    ("PUTERE_NOMINALA_SKID", "Branduri cu puterea din coloana 'Skid power'", "lista"),
    ("SKIDURI_PENTRU_RACK", "REPT: skiduri compatibile (brand = rack sau dc)", "perechi_text"),
    ("RACKURI_PER_BATERIE", "REPT: rack-uri / baterie", "intreg"),
    ("DC_INPUTURI_PER_BATERIE", "Intrări DC / baterie (regula Min/Max)", "intreg"),
]


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


# ---------------- actualizare automată de pe GitHub ----------------
def _descarca_text(url, timeout=15):
    cerere = urllib.request.Request(url + f"?t={int(time.time())}", headers={"Cache-Control": "no-cache"})
    with urllib.request.urlopen(cerere, timeout=timeout) as r:
        return r.read().decode("utf-8")


def _normalizat(text):
    return text.replace("\r\n", "\n")


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


def incarca_date():
    """Citește toate fișierele de pricing și returnează un dicționar de DataFrame-uri."""
    sursa = cale_pricing()
    foi = citeste_copie(sursa, "*")
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
                            (BAT_COL_PRET, "Preț DDP / MWh", "numar"), (BAT_COL_TRANSPORT, "Transport", "numar")]),
    ("skid", FOAIE_SKID, [(SKID_COL_BRAND, "Brand skid", "text"), (SKID_COL_TYPE, "Model (Type)", "text"),
                          (SKID_COL_PRET, "Preț", "numar"),
                          (SKID_COL_TRANSPORT, "Transport", "numar"), (SKID_COL_COM, "Commissioning", "numar")]),
    ("skid_poc", "MV SKID for different POC",
     [(SKID_COL_BRAND, "Brand skid", "text"), (SKID_COL_TYPE, "Model (Type)", "text"),
      (SKID_COL_PRET, "Preț", "numar"), (SKID_COL_TRANSPORT, "Transport", "numar")]),
]
# Texte care trebuie să apară în antetul unor coloane (verificare sigură, independentă de istoric)
ANTETE_OBLIGATORII = {("bat", BAT_COL_TYPE): "type", ("bat", BAT_COL_PRET_DAP): "dap", ("bat", BAT_COL_PRET): "ddp",
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
    o["kva_skid"] = consum_aux_skid(o["putere_unit"])
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
    text = (f"aux: necesar {o['kva_necesar']:g} kVA (baterii {o['kva_baterii']:g} + skid {o['nr']} x "
            f"{o['kva_skid']:g} kVA), "
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
                            ["  din care consum skiduri (kVA)", skid["nr"] * skid.get("kva_skid", 0)],
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


def urmatoarea_versiune(nr_proiect, proiect, nume_existente=()):
    """Caută ofertele deja generate pentru același nr. proiect + proiect (folderele din FOLDER_OFERTE și,
    pentru ofertele mai vechi, cele de lângă program; nume_existente = folderele din Drive) și returnează
    următoarea versiune (1 dacă nu e niciuna)."""
    model = re.compile(re.escape(f"{nume_fisier_valid(nr_proiect)}_BESS_Indicative_Proposal_Taliva_"
                                 f"{nume_fisier_valid(proiect)}_") + r".*_v(\d+)(\.pptx)?$", re.I)
    maxim = 0
    foldere = [os.path.join(radacina, f) for radacina in (FOLDER_OFERTE, ".") if os.path.isdir(radacina)
               for f in os.listdir(radacina)]
    nume_gasite = list(nume_existente)
    for folder in filter(os.path.isdir, foldere):
        try:
            continut = os.listdir(folder)
        except OSError:
            continut = []
        nume_gasite += [os.path.basename(folder)] + continut
    for nume in nume_gasite:
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
# ======================================================================
# DRIVE (Microsoft Graph): fișierele direct din OneDrive/SharePoint, fără sincronizare
# ======================================================================
DRIVE = None          # DriveCloud conectat, sau None = se lucrează cu folderele sincronizate


class EroareDrive(Exception):
    pass


class NegasitDrive(EroareDrive):
    pass


def _cale_url(text):
    return urllib.parse.quote(text, safe="/")


def _mesaj_http(eroare):
    try:
        corp = json.loads(eroare.read().decode("utf-8", "replace"))
        return corp.get("error_description") or corp.get("error", {}).get("message") or str(eroare)
    except Exception:
        return str(eroare)


class DriveCloud:
    """Logare Microsoft (cod + PKCE, browser) și operații pe fișiere prin Microsoft Graph.
    Doar biblioteci standard Python, fără pachete suplimentare."""
    GRAPH = "https://graph.microsoft.com/v1.0"

    def __init__(self):
        folder = os.path.dirname(FISIER_SETARI)
        self.fisier_token = os.path.join(folder, "token_drive.json")
        self.folder_lucru = os.path.join(folder, "Drive cache")       # copiile locale = folderul de lucru
        self.fisier_stare = os.path.join(self.folder_lucru, "drive_stare.json")
        self.token = self._citeste_json(self.fisier_token)
        self.pricing = self.ofertare = None                            # (drive_id, item_id) ale celor 2 foldere
        self.cont = {}
        self.pricing_modificat_la = None

    @staticmethod
    def _citeste_json(cale):
        try:
            with open(cale, encoding="utf-8") as f:
                return json.load(f)
        except (OSError, ValueError):
            return {}

    # ---------------- logare ----------------
    def are_logare(self):
        return bool(self.token.get("refresh_token"))

    def _cere_token(self, campuri):
        date = urllib.parse.urlencode(dict(campuri, client_id=DRIVE_CLIENT_ID, scope=DRIVE_SCOPURI)).encode()
        cerere = urllib.request.Request(
            f"https://login.microsoftonline.com/{DRIVE_TENANT_ID}/oauth2/v2.0/token", data=date)
        try:
            with urllib.request.urlopen(cerere, timeout=60) as r:
                raspuns = json.load(r)
        except urllib.error.HTTPError as e:
            raise EroareDrive(f"Logarea Microsoft a eșuat: {_mesaj_http(e)}")
        raspuns["expira"] = time.time() + int(raspuns.get("expires_in", 3600)) - 120
        raspuns.setdefault("refresh_token", self.token.get("refresh_token"))
        self.token = raspuns
        os.makedirs(os.path.dirname(self.fisier_token), exist_ok=True)
        with open(self.fisier_token, "w", encoding="utf-8") as f:
            json.dump(raspuns, f)

    def logare(self):
        """Deschide browserul pentru logare și primește răspunsul pe http://localhost:<port>."""
        verificator = secrets.token_urlsafe(64)
        provocare = base64.urlsafe_b64encode(hashlib.sha256(verificator.encode()).digest()).decode().rstrip("=")
        stare = secrets.token_urlsafe(16)
        primit = {}

        class Raspuns(BaseHTTPRequestHandler):
            def do_GET(h):
                q = urllib.parse.parse_qs(urllib.parse.urlparse(h.path).query)
                if "code" in q or "error" in q:
                    primit.update({k: v[0] for k, v in q.items()})
                h.send_response(200)
                h.send_header("Content-Type", "text/html; charset=utf-8")
                h.end_headers()
                h.wfile.write("<h3>Logare terminată. Poți închide această fereastră și reveni în "
                              "Generator Oferte BESS.</h3>".encode("utf-8"))

            def log_message(h, *argumente):
                pass

        server = HTTPServer(("127.0.0.1", 0), Raspuns)
        server.timeout = 1
        redirect = f"http://localhost:{server.server_address[1]}"
        webbrowser.open(f"https://login.microsoftonline.com/{DRIVE_TENANT_ID}/oauth2/v2.0/authorize?" +
                        urllib.parse.urlencode({
                            "client_id": DRIVE_CLIENT_ID, "response_type": "code", "redirect_uri": redirect,
                            "response_mode": "query", "scope": DRIVE_SCOPURI, "state": stare,
                            "code_challenge": provocare, "code_challenge_method": "S256",
                            "prompt": "select_account"}))
        sfarsit = time.time() + 300
        try:
            while not primit and time.time() < sfarsit:
                server.handle_request()
        finally:
            server.server_close()
        if "code" not in primit:
            raise EroareDrive(primit.get("error_description") or "Logarea nu s-a terminat (anulată sau timp expirat).")
        if primit.get("state") != stare:
            raise EroareDrive("Răspuns de logare invalid.")
        self._cere_token({"grant_type": "authorization_code", "code": primit["code"],
                          "redirect_uri": redirect, "code_verifier": verificator})

    def _token_acces(self):
        if self.token.get("access_token") and time.time() < self.token.get("expira", 0):
            return self.token["access_token"]
        if self.token.get("refresh_token"):
            try:
                self._cere_token({"grant_type": "refresh_token", "refresh_token": self.token["refresh_token"]})
                return self.token["access_token"]
            except EroareDrive:
                pass                                    # logarea a expirat -> din nou în browser
        self.logare()
        return self.token["access_token"]

    # ---------------- cereri Graph ----------------
    def _cerere(self, metoda, cale, date=None, tip=None, absolut=False, autentificare=True, brut=False, antete=None):
        antete = dict(antete or {})
        if autentificare:
            antete["Authorization"] = "Bearer " + self._token_acces()
        if tip:
            antete["Content-Type"] = tip
        cerere = urllib.request.Request(cale if absolut else self.GRAPH + cale, data=date, method=metoda,
                                        headers=antete)
        try:
            with urllib.request.urlopen(cerere, timeout=300) as r:
                continut = r.read()
        except urllib.error.HTTPError as e:
            clasa = NegasitDrive if e.code == 404 else EroareDrive
            raise clasa(f"Drive ({e.code}): {_mesaj_http(e)}")
        except urllib.error.URLError as e:
            raise EroareDrive(f"Nu mă pot conecta la Microsoft (internet?): {e.reason}")
        if brut:
            return continut
        return json.loads(continut) if continut else {}

    def _folder_din_link(self, link):
        """(drive_id, item_id) pentru un link de folder din SharePoint/OneDrive."""
        cod = "u!" + base64.urlsafe_b64encode(link.encode()).decode().rstrip("=")
        try:
            j = self._cerere("GET", f"/shares/{cod}/driveItem")
            return j["parentReference"]["driveId"], j["id"]
        except EroareDrive:
            pass
        # link direct (.../:f:/r/personal/<site>/Documents/<folder>): site personal + cale
        u = urllib.parse.urlparse(link)
        parti = [urllib.parse.unquote(p) for p in u.path.split("/") if p]
        if "r" in parti:
            parti = parti[parti.index("r") + 1:]
        drive = self._cerere("GET", f"/sites/{u.netloc}:/{_cale_url('/'.join(parti[:2]))}:/drive")
        j = self._cerere("GET", f"/drives/{drive['id']}/root:/{_cale_url('/'.join(parti[3:]))}")
        return drive["id"], j["id"]

    def _copii(self, folder, subcale=""):
        drive_id, item_id = folder
        cale = (f"/drives/{drive_id}/items/{item_id}" + (f":/{_cale_url(subcale)}:" if subcale else "") +
                "/children?$top=999")
        elemente, absolut = [], False
        while cale:
            j = self._cerere("GET", cale, absolut=absolut)
            elemente += j.get("value", [])
            cale, absolut = j.get("@odata.nextLink"), True
        return elemente

    def _descarca(self, drive_id, item_id, destinatie):
        info = self._cerere("GET", f"/drives/{drive_id}/items/{item_id}")
        continut = self._cerere("GET", info["@microsoft.graph.downloadUrl"], absolut=True,
                                autentificare=False, brut=True)
        with open(destinatie + ".tmp", "wb") as f:
            f.write(continut)
        os.replace(destinatie + ".tmp", destinatie)

    def _asigura_folder(self, drive_id, parinte_id, nume):
        """ID-ul subfolderului 'nume' (îl creează dacă nu există)."""
        try:
            j = self._cerere("POST", f"/drives/{drive_id}/items/{parinte_id}/children", tip="application/json",
                             date=json.dumps({"name": nume, "folder": {},
                                              "@microsoft.graph.conflictBehavior": "fail"}).encode())
            return j["id"]
        except NegasitDrive:
            raise
        except EroareDrive:                                  # există deja (409)
            return self._cerere("GET", f"/drives/{drive_id}/items/{parinte_id}:/{_cale_url(nume)}")["id"]

    def _urca(self, drive_id, parinte_id, nume, cale_fisier):
        with open(cale_fisier, "rb") as f:
            date = f.read()
        tinta = f"/drives/{drive_id}/items/{parinte_id}:/{_cale_url(nume)}:"
        if len(date) < 4 * 1024 * 1024:
            self._cerere("PUT", tinta + "/content", date=date, tip="application/octet-stream")
            return
        sesiune = self._cerere("POST", tinta + "/createUploadSession", tip="application/json",
                               date=json.dumps({"item": {"@microsoft.graph.conflictBehavior": "replace"}}).encode())
        bucata = 320 * 1024 * 16                              # multiplu de 320 KiB, cum cere Graph
        for start in range(0, len(date), bucata):
            parte = date[start:start + bucata]
            self._cerere("PUT", sesiune["uploadUrl"], date=parte, absolut=True, autentificare=False, brut=True,
                         antete={"Content-Range": f"bytes {start}-{start + len(parte) - 1}/{len(date)}"})

    # ---------------- folosite de program ----------------
    def conecteaza(self):
        os.makedirs(self.folder_lucru, exist_ok=True)
        self.cont = self._cerere("GET", "/me?$select=displayName,mail,userPrincipalName")
        self.pricing = self._folder_din_link(LINK_FOLDER_PRICING)
        self.ofertare = self._folder_din_link(LINK_FOLDER_OFERTARE)
        self.actualizeaza()

    def actualizeaza(self):
        """Descarcă în folderul de lucru fișierele care s-au schimbat în Drive. Returnează numele lor."""
        stare = self._citeste_json(self.fisier_stare)
        de_luat = [(self.pricing, FISIER_PRICING)] + [(self.ofertare, f) for f in
                                                      (FISIER_PCS, FISIER_SPECS, FISIER_EXCEL_TEMPLATE,
                                                       FISIER_PPT_TEMPLATE, FISIER_SETARI_GENERATOR)]
        listari, actualizate = {}, []
        for folder, nume in de_luat:
            if folder not in listari:
                listari[folder] = {c["name"].lower(): c for c in self._copii(folder) if "file" in c}
            element = listari[folder].get(nume.lower())
            if element is None and nume == FISIER_SETARI_GENERATOR:      # opțional: fără el rămân regulile din cod
                if os.path.exists(os.path.join(self.folder_lucru, nume)):
                    os.remove(os.path.join(self.folder_lucru, nume))
                stare.pop(nume, None)
                continue
            if element is None and nume in (FISIER_PCS, FISIER_SPECS):      # pot fi foi în Setari generator.xlsx
                if os.path.exists(os.path.join(self.folder_lucru, nume)):
                    os.remove(os.path.join(self.folder_lucru, nume))
                stare.pop(nume, None)
                continue
            if element is None:
                unde = "Purchasing Prices Update_BESS" if folder == self.pricing else "Ofertare BESS"
                raise EroareDrive(f"Nu găsesc '{nume}' în folderul '{unde}' din Drive.")
            destinatie = os.path.join(self.folder_lucru, nume)
            if stare.get(nume) != element.get("eTag") or not os.path.exists(destinatie):
                self._descarca(folder[0], element["id"], destinatie)
                stare[nume] = element.get("eTag")
                actualizate.append(nume)
            if nume == FISIER_PRICING:
                self.pricing_modificat_la = element.get("lastModifiedDateTime")
        with open(self.fisier_stare, "w", encoding="utf-8") as f:
            json.dump(stare, f)
        return actualizate

    # ---------------- Excel Online (tabelul cu oferte) ----------------
    def fisier_din_link(self, link):
        """(drive_id, item_id) pentru un link de fișier (ex. BESS_Proposals.xlsx)."""
        cod = "u!" + base64.urlsafe_b64encode(link.encode()).decode().rstrip("=")
        j = self._cerere("GET", f"/shares/{cod}/driveItem")
        return j["parentReference"]["driveId"], j["id"]

    def _url_foaie(self, fisier, foaie):
        return f"/drives/{fisier[0]}/items/{fisier[1]}/workbook/worksheets/{urllib.parse.quote(foaie, safe='')}"

    def citeste_zona(self, fisier, foaie, adresa):
        """Valorile unei zone, ex. 'X9:X3000' -> [[v], [v], ...]."""
        return self._cerere("GET", self._url_foaie(fisier, foaie) + f"/range(address='{adresa}')?$select=values")["values"]

    def scrie_zona(self, fisier, foaie, adresa, valori=None, format_numar=None):
        corp = {}
        if valori is not None:
            corp["values"] = valori
        if format_numar is not None:
            corp["numberFormat"] = format_numar
        try:
            self._cerere("PATCH", self._url_foaie(fisier, foaie) + f"/range(address='{adresa}')",
                         date=json.dumps(corp).encode(), tip="application/json")
        except EroareDrive as e:
            if "(403)" in str(e):
                raise EroareDrive(f"{e}\nContul tău nu are drept de editare pe BESS_Proposals prin Excel Online "
                                  "(cere-i lui Levon acces de editare la fișier).") from None
            raise

    def nume_oferte(self):
        """Numele folderelor din 'Ofertare BESS/Oferte generate' (pentru numerotarea versiunilor)."""
        try:
            nume = []
            for c in self._copii(self.ofertare, FOLDER_OFERTE):
                nume.append(c["name"])
                if "folder" in c and MODEL_FOLDER_ZI.match(c["name"]):       # ofertele din folderele pe zile
                    nume += [x["name"] for x in self._copii(self.ofertare, f"{FOLDER_OFERTE}/{c['name']}")]
            return nume
        except NegasitDrive:
            return []

    def organizeaza_oferte_pe_zile(self):
        """Mută ofertele puse direct în 'Oferte generate' (cele vechi sau făcute cu o versiune veche a programului)
        în folderul zilei în care au fost create, ex. 'Oferte generate/2026-11-02/...'. Returnează câte a mutat."""
        drive_id, item_id = self.ofertare
        try:
            id_oferte = self._cerere("GET", f"/drives/{drive_id}/items/{item_id}:/{_cale_url(FOLDER_OFERTE)}")["id"]
            elemente = self._copii(self.ofertare, FOLDER_OFERTE)
        except NegasitDrive:
            return 0
        foldere_zi, mutate = {}, 0
        for el in elemente:
            if "folder" not in el or MODEL_FOLDER_ZI.match(el["name"]) or el["name"].startswith("_") \
                    or not el.get("createdDateTime"):
                continue
            creat = datetime.fromisoformat(el["createdDateTime"].replace("Z", "+00:00")).astimezone()
            zi = creat.strftime(FORMAT_FOLDER_ZI)
            if zi not in foldere_zi:
                foldere_zi[zi] = self._asigura_folder(drive_id, id_oferte, zi)
            self._cerere("PATCH", f"/drives/{drive_id}/items/{el['id']}", tip="application/json",
                         date=json.dumps({"parentReference": {"id": foldere_zi[zi]}}).encode())
            mutate += 1
        return mutate

    def folder_oferta(self, folder_local):
        """ID-ul folderului ofertei din Drive ('Oferte generate/<zi>/<nume>'), creat dacă nu există."""
        drive_id, item_id = self.ofertare
        id_oferte = self._asigura_folder(drive_id, item_id, FOLDER_OFERTE)
        zi = os.path.basename(os.path.dirname(os.path.abspath(folder_local)))
        if MODEL_FOLDER_ZI.match(zi) or zi == FOLDER_TESTE:
            id_oferte = self._asigura_folder(drive_id, id_oferte, zi)
        return self._asigura_folder(drive_id, id_oferte, os.path.basename(folder_local))

    def sterge_teste(self):
        """Șterge folderul cu calculele de probă rămase în Drive (ex. dacă programul s-a închis brusc)."""
        drive_id, item_id = self.ofertare
        try:
            el = self._cerere("GET", f"/drives/{drive_id}/items/{item_id}:/"
                                     f"{_cale_url(FOLDER_OFERTE + '/' + FOLDER_TESTE)}")
        except NegasitDrive:
            return
        self.sterge((drive_id, el["id"]))

    def urca_fisier(self, id_folder, nume, cale_locala):
        """Urcă un fișier în folderul dat; returnează (drive_id, item_id)."""
        drive_id = self.ofertare[0]
        self._urca(drive_id, id_folder, nume, cale_locala)
        j = self._cerere("GET", f"/drives/{drive_id}/items/{id_folder}:/{_cale_url(nume)}")
        return drive_id, j["id"]

    def link(self, fisier):
        return self._cerere("GET", f"/drives/{fisier[0]}/items/{fisier[1]}?$select=webUrl").get("webUrl")

    def sterge(self, element):
        self._cerere("DELETE", f"/drives/{element[0]}/items/{element[1]}")

    def cerere_excel(self, metoda, cale, corp=None, brut=False, incercari=6):
        """Cerere Graph cu reîncercare: un fișier abia urcat poate fi 'ocupat' câteva secunde pentru Excel Online."""
        for i in range(incercari):
            try:
                return self._cerere(metoda, cale, date=json.dumps(corp).encode() if corp is not None else None,
                                    tip="application/json" if corp is not None else None, brut=brut)
            except EroareDrive as e:
                if i == incercari - 1 or not re.search(r"\((409|423|429|500|502|503|504)\)", str(e)):
                    raise
                time.sleep(2 * (i + 1))

    def batch(self, cereri):
        """Mai multe cereri Graph într-una ($batch), executate în ordine (fiecare depinde de cea dinainte)."""
        corp = {"requests": [dict({"id": str(i + 1), "method": m, "url": url, "body": b,
                                   "headers": {"Content-Type": "application/json"}},
                                  **({"dependsOn": [str(i)]} if i else {}))
                             for i, (m, url, b) in enumerate(cereri)]}
        for incercare in range(5):
            raspuns = self.cerere_excel("POST", "/$batch", corp)
            erori = [r for r in raspuns.get("responses", []) if r.get("status", 200) >= 300]
            temporare = [r for r in erori if r.get("status") in (409, 423, 429, 500, 502, 503, 504, 424)]
            if not erori or len(temporare) < len(erori) or incercare == 4:
                break
            time.sleep(2 * (incercare + 1))                  # fișier ocupat: reîncercăm tot grupul
        if erori:
            r = sorted(erori, key=lambda x: int(x["id"]))[0]
            mesaj = (r.get("body") or {}).get("error", {}).get("message", "")
            raise EroareDrive(f"Excel Online ({r.get('status')}): {mesaj or 'cerere refuzată'}")

    def urca_oferta(self, folder_local, sari=()):
        """Urcă fișierele ofertei în 'Ofertare BESS/Oferte generate/<nume>'. Returnează linkul folderului."""
        drive_id, item_id = self.ofertare
        id_oferte = self._asigura_folder(drive_id, item_id, FOLDER_OFERTE)
        zi = os.path.basename(os.path.dirname(os.path.abspath(folder_local)))
        if MODEL_FOLDER_ZI.match(zi):                                       # 'Oferte generate/2026-11-06/...'
            id_oferte = self._asigura_folder(drive_id, id_oferte, zi)
        id_oferta = self._asigura_folder(drive_id, id_oferte, os.path.basename(folder_local))
        for fisier in sorted(os.listdir(folder_local)):
            cale = os.path.join(folder_local, fisier)
            if os.path.isfile(cale) and fisier not in sari:
                self._urca(drive_id, id_oferta, fisier, cale)
        return self._cerere("GET", f"/drives/{drive_id}/items/{id_oferta}").get("webUrl")

    def urca_setari(self, cale_locala):
        """Urcă Setari generator.xlsx în folderul 'Ofertare BESS'. Returnează linkul fișierului."""
        drive_id, item_id = self.ofertare
        self._urca(drive_id, item_id, FISIER_SETARI_GENERATOR, cale_locala)
        element = self._cerere("GET", f"/drives/{drive_id}/items/{item_id}:/{_cale_url(FISIER_SETARI_GENERATOR)}")
        stare = self._citeste_json(self.fisier_stare)
        stare[FISIER_SETARI_GENERATOR] = element.get("eTag")
        with open(self.fisier_stare, "w", encoding="utf-8") as f:
            json.dump(stare, f)
        return element.get("webUrl")

    def descriere_sursa(self):
        cand = ""
        if self.pricing_modificat_la:
            moment = datetime.fromisoformat(self.pricing_modificat_la.replace("Z", "+00:00")).astimezone()
            cand = f"\n(modificat ultima dată: {moment:%d/%m/%y %H:%M})"
        cont = self.cont.get("mail") or self.cont.get("userPrincipalName", "")
        return f"Prețuri din Drive: Purchasing Prices Update_BESS / {FISIER_PRICING}{cand}\nCont: {cont}"


def _index_coloana(litera):
    n = 0
    for ch in litera.upper():
        n = n * 26 + ord(ch) - 64
    return n


def _fisier_propuneri():
    global _CACHE_FISIER_PROPUNERI
    if _CACHE_FISIER_PROPUNERI is None:
        try:
            _CACHE_FISIER_PROPUNERI = DRIVE.fisier_din_link(LINK_BESS_PROPOSALS)
        except EroareDrive as e:
            raise EroareDrive(f"{e}\n(nu am acces la linkul fișierului BESS_Proposals: verifică dacă se "
                              "deschide în browser cu contul tău de firmă)") from None
    return _CACHE_FISIER_PROPUNERI


_CACHE_FISIER_PROPUNERI = None
_COPIE_PROPUNERI = {"cale": None, "cand": 0.0, "foi": None}
SECUNDE_COPIE_PROPUNERI = 120


def _citeste_din_copie(foaie, adresa):
    """Rezervă când Excel Online refuză citirea (403): descarcă fișierul și citește zona din copie."""
    from openpyxl import load_workbook
    if _COPIE_PROPUNERI["foi"] is None or time.time() - _COPIE_PROPUNERI["cand"] > SECUNDE_COPIE_PROPUNERI:
        cale = os.path.join(DRIVE.folder_lucru, "BESS_Proposals (copie citire).xlsx")
        DRIVE._descarca(*_fisier_propuneri(), cale)
        wb = load_workbook(cale, read_only=True, data_only=True)
        _COPIE_PROPUNERI.update(cale=cale, cand=time.time(), foi={})
        for nume in {cfg["foaie"] for cfg in TABELE_OFERTE.values()}:
            if nume in wb.sheetnames:
                _COPIE_PROPUNERI["foi"][nume] = [list(r) for r in wb[nume].iter_rows(values_only=True)]
        wb.close()
    randuri = _COPIE_PROPUNERI["foi"].get(foaie)
    if randuri is None:
        raise EroareDrive(f"Nu găsesc foaia '{foaie}' în BESS_Proposals.")
    m = re.fullmatch(r"([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?", adresa.upper())
    c1, r1 = _index_coloana(m.group(1)), int(m.group(2))
    c2, r2 = (_index_coloana(m.group(3)), int(m.group(4))) if m.group(3) else (c1, r1)
    zona = []
    for r in range(r1, r2 + 1):
        rand = randuri[r - 1] if r - 1 < len(randuri) else []
        zona.append([("" if (c - 1 >= len(rand) or rand[c - 1] is None) else rand[c - 1]) for c in range(c1, c2 + 1)])
    return zona


def citeste_propuneri(foaie, adresa):
    """Citește o zonă din BESS_Proposals prin Excel Online; la 403 (acces refuzat) din copia descărcată."""
    fisier = _fisier_propuneri()
    try:
        return DRIVE.citeste_zona(fisier, foaie, adresa)
    except EroareDrive as e:
        if "(403)" not in str(e):
            raise
        try:
            return _citeste_din_copie(foaie, adresa)
        except EroareDrive as e2:
            raise EroareDrive(f"Excel Online: {e}\nDescărcare: {e2}") from None


def contracte_din_tabel(tara, cu_liber=False):
    """[(rând, nr. contract, client)] din foaia țării: toate rândurile care au un nr. de contract.
    cu_liber=True: (rând, nr, client, liber), liber = fără dată, client și proiect (coloanele 'col_libere')."""
    if DRIVE is None:
        raise ValueError("Tabelul BESS_Proposals se poate folosi doar cu Drive-ul conectat.")
    cfg = TABELE_OFERTE[tara]
    start, sfarsit = cfg["rand_antet"] + 1, cfg["rand_antet"] + MAX_RANDURI_TABEL
    numere = citeste_propuneri(cfg["foaie"], f"{cfg['col_nr']}{start}:{cfg['col_nr']}{sfarsit}")
    clienti = citeste_propuneri(cfg["foaie"],
                                f"{cfg['col_client']}{start}:{cfg['col_client']}{sfarsit}")
    randuri = [(start + i, str(n[0]).strip(), str(c[0] or "").strip())
               for i, (n, c) in enumerate(zip(numere, clienti)) if str(n[0]).strip()]
    if not cu_liber:
        return randuri
    alte = [citeste_propuneri(cfg["foaie"], f"{col}{start}:{col}{sfarsit}")
            for col in cfg["col_libere"]]
    return [(r, nr, client, all(str(coloana[r - start][0] or "").strip() == "" for coloana in alte))
            for r, nr, client in randuri]


def nr_contract_liber(tara):
    """Nr. de contract pentru o ofertă nouă: primul nr. pregătit pe un rând gol (fără dată, client, proiect);
    dacă nu mai e niciunul, ultimul nr. ST din foaie + 1 (va primi un rând nou). Returnează (nr, rând_nou)."""
    randuri = contracte_din_tabel(tara, cu_liber=True)
    liber = next((nr for _, nr, _, gol in randuri if re.fullmatch(r"ST\d+", nr) and gol), None)
    if liber:
        return liber, False
    numere = [int(nr[2:]) for _, nr, _, _ in randuri if re.fullmatch(r"ST\d+", nr)]
    return (f"ST{max(numere) + 1}", True) if numere else (None, False)


def _rand_nou(cfg):
    """Primul rând de după ultimul rând folosit (nr. de contract, dată, client sau proiect)."""
    start, sfarsit = cfg["rand_antet"] + 1, cfg["rand_antet"] + MAX_RANDURI_TABEL
    ultimul = cfg["rand_antet"]
    for col in cfg["col_nr"] + cfg["col_libere"]:
        valori = citeste_propuneri(cfg["foaie"], f"{col}{start}:{col}{sfarsit}")
        folosite = [i for i, v in enumerate(valori) if str(v[0] or "").strip()]
        if folosite:
            ultimul = max(ultimul, start + folosite[-1])
    return ultimul + 1


def _valoare_tabel(inregistrare, camp):
    valoare = inregistrare.get(camp)
    if camp.startswith("data_") and valoare:
        zi = datetime.strptime(valoare, "%Y-%m-%d")
        return (zi - datetime(1899, 12, 30)).days            # dată Excel (număr de zile)
    return "" if valoare is None else valoare


def _valoare_text(inregistrare, camp):
    valoare = inregistrare.get(camp)
    if camp.startswith("data_") and valoare:
        return datetime.strptime(valoare, "%Y-%m-%d").strftime("%d.%m.%Y")
    if camp in ("pret_total", "pret_mwh", "cost") and valoare not in (None, ""):
        return f"{bani_eu(valoare)} $"
    return "" if valoare is None else str(valoare)


# Numele proiectului în BESS_Proposals, ca în restul tabelului: 'Sinesti_BESS, 12,5MWh, 4MW'
NUME_PROIECT_TABEL = "{proiect}_BESS, {mwh}MWh, {mw}MW"


def nume_proiect_tabel(proiect, mwh, mw):
    """'Sinesti', 12.5, 4 -> 'Sinesti_BESS, 12,5MWh, 4MW'. Dacă numele are deja MWh/MW, rămâne cum e."""
    proiect = str(proiect).strip()
    if re.search(r"MWh", proiect, re.I) and re.search(r"\bMW\b|\dMW", proiect):
        return proiect
    proiect = re.sub(r"[_\s]*BESS\s*$", "", proiect, flags=re.I)        # fără '_BESS_BESS'
    try:
        return NUME_PROIECT_TABEL.format(proiect=proiect, mwh=nr_eu(float(mwh)), mw=nr_eu(float(mw)))
    except (TypeError, ValueError):
        return proiect


def trece_in_tabel(inregistrare, confirma_suprascriere):
    """Completează în BESS_Proposals rândul cu nr. de contract al ofertei. Returnează rândul sau None (anulat)."""
    inregistrare = dict(inregistrare, proiect=nume_proiect_tabel(inregistrare.get("proiect", ""),
                                                                 inregistrare.get("mwh"), inregistrare.get("mw")))
    cfg = TABELE_OFERTE[inregistrare["tara"]]
    rand, client_existent = next(((r, c) for r, nr, c in contracte_din_tabel(inregistrare["tara"])
                                  if nr == inregistrare["nr"]), (None, None))
    rand_nou = rand is None
    if rand_nou:                                       # nr. nou (după ultimul din tabel) -> rând nou
        if not re.fullmatch(r"ST\d+", inregistrare["nr"]):
            raise ValueError(f"Nr. de contract '{inregistrare['nr']}' nu e de forma ST... și nu există în foaia "
                             f"'{cfg['foaie']}'. Pare o ofertă de test: șterge-o din listă.")
        rand = _rand_nou(cfg)
    elif client_existent and not confirma_suprascriere(rand, client_existent):
        return None
    celule = {col: _valoare_tabel(inregistrare, camp) for col, camp in cfg["coloane"].items()}
    if rand_nou:
        celule[cfg["col_nr"]] = inregistrare["nr"]
    coloane = sorted(celule, key=_index_coloana)
    grupuri = [[coloane[0]]]                                   # coloane alăturate -> o singură scriere
    for col in coloane[1:]:
        if _index_coloana(col) == _index_coloana(grupuri[-1][-1]) + 1:
            grupuri[-1].append(col)
        else:
            grupuri.append([col])
    fisier = _fisier_propuneri()
    for g in grupuri:
        DRIVE.scrie_zona(fisier, cfg["foaie"], f"{g[0]}{rand}:{g[-1]}{rand}", valori=[[celule[c] for c in g]])
    for col, camp in cfg["coloane"].items():
        if camp.startswith("data_") and celule[col] != "":
            DRIVE.scrie_zona(fisier, cfg["foaie"], f"{col}{rand}", format_numar=[["dd.mm.yyyy"]])
    return rand


FISIER_OFERTE_NETRECUTE = os.path.join(os.path.dirname(FISIER_SETARI), "oferte_netrecute.json")


def oferte_netrecute():
    try:
        with open(FISIER_OFERTE_NETRECUTE, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return []


def salveaza_oferte_netrecute(lista):
    os.makedirs(os.path.dirname(FISIER_OFERTE_NETRECUTE), exist_ok=True)
    with open(FISIER_OFERTE_NETRECUTE, "w", encoding="utf-8") as f:
        json.dump(lista, f, ensure_ascii=False, indent=2)


def porneste_drive():
    """Conectarea la Drive la pornire. None = utilizatorul a ales folderele sincronizate."""
    drive = DriveCloud()
    try:
        if not drive.are_logare():
            messagebox.showinfo("Logare Microsoft",
                                "Se deschide browserul pentru logarea cu contul de firmă (doar prima dată).\n\n"
                                "La 'Permissions requested' apasă Accept.")
        drive.conecteaza()
        return drive
    except Exception as e:
        if messagebox.askyesno("Drive indisponibil",
                               f"Nu m-am putut conecta la Drive:\n{e}\n\n"
                               "Continui cu folderele sincronizate pe acest calculator?"):
            return None
        sys.exit(1)


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
    return (p["nr_baterii"] * (b["cap"] * b["pret"] + b["transport"])
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
        p["auto_comparatie"] = [f"{model}: {bani_eu(cost, 0)} $" for cost, model, _ in variante]
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


def calculeaza_oferta(date, intrari, baterie, nr_baterii, skid, trafo, status=lambda text: None, nume_existente=(),
                      online=False):
    """Pasul 1: doar Excelul de calcul, în 'Oferte generate/<nume>_vN'. Prezentarea se face după verificare."""
    # Folderul, prezentarea și PDF-ul au același nume, ex. ST202606109_BESS_..._20MWh_10MW_v2
    intrari["versiune"] = urmatoarea_versiune(intrari["nr_proiect"], intrari["proiect"], nume_existente)
    nume = NUME_PREZENTARE.format(nr_proiect=nume_fisier_valid(intrari["nr_proiect"]),
                                  proiect=nume_fisier_valid(intrari["proiect"]),
                                  mwh=nr_eu(intrari["mwh"]), mw=nr_eu(intrari["mw"])) \
           + f"_v{intrari['versiune']}"
    zi = FOLDER_TESTE if e_calcul_test(intrari) else datetime.now().strftime(FORMAT_FOLDER_ZI)
    folder = os.path.abspath(os.path.join(FOLDER_OFERTE, zi, nume))
    os.makedirs(folder, exist_ok=True)

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


def finalizeaza_oferta(calcul, intrari, baterie, nr_baterii, skid, trafo, status=lambda text: None, reciteste=True):
    """Pasul 2 (după OK): prezentarea PPTX + PDF, cu valorile din Excelul de calcul (recitit, dacă l-ai modificat)."""
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
        "[SKID]": skid["tip"],
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

    r["tip"] = ("Aviz tehnic de racordare" if re.search(r"AVIZ\s+TEHNIC\s+DE\s+RACORDARE", text, re.I)
                else "Studiu de soluție" if re.search(r"studiu\w*\s+de\s+solu", plat, re.I)
                else "Schemă de stație" if re.search(r"Schema\s+normal", plat, re.I) else "document")
    if r["tip"] in ("Schemă de stație", "document"):
        av.append("Documentul nu pare a fi un ATR sau un studiu de soluție"
                  + (" (e schema normală a unei stații)" if r["tip"] == "Schemă de stație" else "")
                  + ": valorile de mai jos pot lipsi sau nu se aplică.")
    for nume, cuvant in (("Distribuție Energie Electrică Romania (Electrica)", r"Distribu\w+ Energie Electric"),
                         ("E-Distribuție", r"e-distribu"), ("Delgaz Grid", r"Delgaz"),
                         ("Rețele Electrice", r"Re[țt]ele Electrice"), ("Distribuție Oltenia", r"Distribu\w+ Oltenia"),
                         ("Transelectrica", r"Transelectrica")):
        if re.search(cuvant, original, re.I):
            r["operator"] = nume
            break
    data_re = r"(\d{2}[./]\d{2}[./]\d{4})"
    m = re.search(r"nr\.?\s*(\d{4,})\s*/\s*(?:data\s*)?" + data_re, plat, re.I) or \
        re.search(r"racordare\s+nr\.?\s*(\d{4,})\s+din\s+(?:data\s+)?" + data_re, plat, re.I) or \
        re.search(r"racordare\s+nr\s+din\s+(?:data\s+)?(\d{4,})\s+" + data_re, plat, re.I)  # câmpuri mutate la final
    if r["tip"] != "Aviz tehnic de racordare":       # studiu: ATR-ul la care se referă
        m = re.search(r"avizul\s+tehnic\s+de\s+racordare\s+nr\.?\s*(\d+)\s+din\s+(\d{2}\.\d{2}\.\d{4})", plat, re.I)
    if m:
        r["nr"], r["data"] = m.group(1), m.group(2).replace("/", ".")
    r["consumator"] = bool(re.search(r"PENTRU\s+CONSUMATOR", plat, re.I))   # rafinat după puterea aprobată
    r["beneficiar"] = _cauta(r"adresat[ăa] de\s+(.+?),\s+pentru", plat) or \
        _cauta(r"utilizatorului\s+(.{3,80}?),\s+cu\s+(?:domiciliul|sediul)", plat) or \
        _cauta(r"Beneficiar\s*:?\s*([A-ZĂÂÎȘȚŞŢ0-9 .&-]+?S\.?R\.?L\.?)", plat)
    if not r["beneficiar"]:      # Rețele Electrice (consumator): valorile apar după etichete, la capăt de rând
        zona = _cauta(r"utilizatorului(.{0,250})", plat) or ""
        r["beneficiar"] = _cauta(r"((?:S\.?C\.?\s+)?[A-ZĂÂÎȘȚŞŢ][A-ZĂÂÎȘȚŞŢ0-9 .&-]{2,60}?\s(?:S\.?R\.?L\.?|S\.?A\.?))(?![A-Z])",
                                 zona, re.S)
    r["obiectiv"] = _cauta(r"adresat[ăa] de .{0,120}?,\s+pentru\s+(.{0,200}?)\s+ce\s+a?par[țţt]ine", plat) or \
        _cauta(r"se\s+aprob\w*\s+racordarea.{0,90}?(?:consum\s+[sș]i\s+de\s+producere|producere|consum)\s+(?:permanent/temporar\s+)?"
               r"([^()]{2,80}?)\s*\(denumirea\)", plat) or \
        _cauta(r"\(denumirea[^)]*\)\s*([A-ZĂÂÎȘȚ0-9][A-ZĂÂÎȘȚ0-9 .,+-]{2,60}?)\s*(?:\n|amplasat)", text, re.S)
    r["amplasament"] = _cauta(r"amplasat\(?[ăa]?\)?\s+[îi]n\s+(jude[țţt]ul\s+[^,]+(?:,\s*-?\s*[^,]+)?)", plat)
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
    if kv:
        r["poc_kv"] = _nr_ro(kv)

    # Puterea aprobată (secțiunea 2, etapa finală = ultima valoare de pe rând). Formatele diferă de la un operator la
    # altul: eticheta e înainte sau după valori, kW sau kVA primul, '-' pentru etapele goale -> fiecare rând de valori
    # ia eticheta cea mai apropiată.
    sec = re.search(r"2\.\s*Puterea\s+aprobat\w*(.+?)3\.\s*Descrierea", text, re.I | re.S) or \
        re.search(r"Puterea\s+aprobat\w*\s*:\s*1\.(.+?)2\.\s*Descrierea", text, re.I | re.S) or \
        re.search(r"2\.\s*Puterea\s+aprobat\w*(.{0,3000})", text, re.I | re.S)
    if sec:
        corp = sec.group(1)
        model_rand = r"\((kW|kVA)\)((?:\s+(?:" + _NR + r"|-))+)"
        randuri = list(re.finditer(model_rand, corp))
        # etichetele se caută pe textul fără rândurile de valori (uneori un rând taie eticheta în două)
        fara_valori = re.sub(model_rand, lambda m: " " * len(m.group(0)), corp)
        etichete = [(m.start(), m.end(), re.sub(r"\s+", " ", m.group(0)).lower()) for m in re.finditer(
            r"Puterea\s+maxim\w*\s+simultan\w*\s+ce\s+poate\s+fi\s+(?:evacuat\w*(?:\s+f[aă]r[aă]\s+realiz\w*|"
            r"\s+[îi]n\s+situa\w*)?|absorbit\w*)", fara_valori, re.I)]
        # Un rând (kVA) ține de eticheta dinaintea lui. Un rând (kW) lipit de eticheta următoare ('4100,000Puterea...',
        # cum iese din unele PDF-uri Electrica) ține de aceea; altfel tot de eticheta dinaintea lui.
        for m in randuri:
            valori = re.findall(_NR, m.group(2))
            if not valori or not etichete:
                continue
            urmatoare = [e for e in etichete if e[0] == m.end()]
            inainte = [e for e in etichete if e[0] <= m.start()]
            if m.group(1) == "kW" and urmatoare:
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
            r.setdefault(f"{tip}_{m.group(1).lower()}", valoare)
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
    if "mwh" not in r:
        v = _cauta(r"(" + _NR + r")\s*MWh", plat)
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
        av.append(f"Pi stocare din tabel ({nr_eu(pi)} kW) e de peste 2,5 ori evacuarea: pare trecută capacitatea "
                  "(kWh) în coloana de kW. Nu am folosit-o ca limită.")
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
    if "mwh" not in r:
        av.append("Capacitatea (MWh) nu apare în ATR (doar Ah): ia-o din cererea clientului.")
    if not evac and not r["consumator"]:
        av.append("Nu am găsit puterea aprobată pentru evacuare.")
    return r


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
        "curs": numar_cerere(c.get("curs"), 0.89),
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


class AplicatieOferte:
    def __init__(self, root):
        self.root = root
        self.root.title("Generator Oferte BESS - Taliva")
        # Fereastra se potrivește pe ecran (laptop); ce nu încape se vede cu bara de derulare / rotița mouse-ului
        latime = min(1180, root.winfo_screenwidth() - 40)
        inaltime = min(940, root.winfo_screenheight() - 90)
        self.root.geometry(f"{latime}x{inaltime}+10+10")
        # butonul principal și starea stau fix jos (mereu vizibile); restul se derulează
        self._jos = ttk.Frame(root)
        self._jos.pack(side="bottom", fill="x")
        ttk.Separator(self._jos).pack(fill="x")
        root = self._zona_derulare(root)
        self.continut = root

        ttk.Label(root, text="Sistem Ofertare BESS", font=("Helvetica", 16, "bold")).pack(pady=10)
        # Tab-uri: oferta, componentele introduse manual și verificarea calculului (înainte de PowerPoint / PDF)
        taburi = ttk.Notebook(root)
        taburi.pack(padx=10, fill="x")
        self.taburi = taburi
        tab_oferta, tab_manual, tab_verificare = ttk.Frame(taburi), ttk.Frame(taburi), ttk.Frame(taburi)
        tab_atr = ttk.Frame(taburi)
        taburi.add(tab_oferta, text="  Ofertă  ")
        taburi.add(tab_manual, text="  Componente manuale  ")
        taburi.add(tab_verificare, text="  Verificare calcul  ")
        taburi.add(tab_atr, text="  📄 Citire ATR  ")
        self.tab_oferta = tab_oferta
        self.tab_verificare = tab_verificare
        self._calcul_in_asteptare = None
        self._construieste_tab_verificare(tab_verificare)
        self._construieste_tab_atr(tab_atr)
        frame = ttk.Frame(tab_oferta)
        frame.pack(side="left", padx=20, pady=10, anchor="n")
        self._construieste_panou_tabel(tab_oferta)
        self._construieste_tab_manual(tab_manual)

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
            baterii = baterii_din_lista(self.baterii_map) or baterii
            branduri = ["Oricare"] + self._branduri_skid()
        except Exception as e:
            messagebox.showwarning("Atenție", f"Nu am putut citi fișierele Excel la pornire:\n{e}")

        self.intrari = {}
        campuri = [
            ("Client:", "GEP"), ("Nume Proiect:", "Proiect X"),
            # cu Drive-ul conectat rămâne gol: nr. de contract se ia automat din BESS_Proposals
            ("Nr. proiect:", "" if DRIVE is not None else NR_PROIECT_IMPLICIT.format(an_luna=datetime.now().strftime("%Y%m"))),
            ("Locație:", "București, România"),
            ("Capacitate (MWh):", "20"), ("Putere (MW):", "10"), ("Putere max. admisă (MW, opțional):", ""),
            ("Valoare EMS (USD):", f"{EMS_AUTO:g}"),
            ("Valoare ENG (USD):", "0"), ("Adaos Comercial (Margin):", f"{ADAOS_IMPLICIT:.2f}"),
            ("Curs USD/EUR:", f"{CURS_IMPLICIT:g}"),
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
        self.combo_baterie = ttk.Combobox(frame, values=[OPTIUNE_AUTO_BATERIE] + baterii + [OPTIUNE_MANUALA_BATERIE],
                                          state="readonly",
                                          width=40)
        self.combo_baterie.set(OPTIUNE_AUTO_BATERIE)
        self.combo_baterie.grid(row=r, column=1, sticky="e", pady=5)

        ttk.Label(frame, text="Brand Skid preferat:").grid(row=r + 1, column=0, sticky="w", pady=5)
        self.combo_brand = ttk.Combobox(frame, values=branduri + [OPTIUNE_MANUALA_SKID], state="readonly", width=40)
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

        # Trafo aux extern: nicio bifă = automat (doar unde skidul nu acoperă consumul); cele două bife se exclud
        self.foloseste_aux, self.fara_aux = tk.BooleanVar(), tk.BooleanVar()
        ttk.Checkbutton(frame, text="Forțează trafo aux extern (la toate skidurile)", variable=self.foloseste_aux,
                        command=lambda: self.foloseste_aux.get() and self.fara_aux.set(False))\
            .grid(row=r + 5, column=0, columnspan=2, pady=(10, 2), sticky="w")
        ttk.Checkbutton(frame, text="Fără trafo aux extern (nu se adaugă deloc)", variable=self.fara_aux,
                        command=lambda: self.fara_aux.get() and self.foloseste_aux.set(False))\
            .grid(row=r + 6, column=0, columnspan=2, pady=(2, 10), sticky="w")
        ttk.Label(frame, text="Fără nicio bifă: trafo aux extern adăugat automat doar unde skidul nu acoperă consumul.",
                  foreground="gray").grid(row=r + 7, column=0, columnspan=2, sticky="w")

        # Excel + PDF prin Microsoft 365 (Excel Online + conversie PDF în Drive): merge fără Excel / PowerPoint
        # instalate (Mac). Fără Excel instalat e singura variantă; altfel se ține minte alegerea.
        fara_office = xw is None or sys.platform != "win32"
        self.mod_online = tk.BooleanVar(value=DRIVE is not None and
                                        (fara_office or bool(citeste_setari().get("mod_m365"))))
        bifa = ttk.Checkbutton(frame, text="Calcul Excel și PDF prin Microsoft 365 (fără Excel / PowerPoint instalate)",
                               variable=self.mod_online,
                               command=lambda: salveaza_setari(dict(citeste_setari(), mod_m365=self.mod_online.get())))
        bifa.grid(row=r + 8, column=0, columnspan=2, pady=(10, 2), sticky="w")
        if DRIVE is None or fara_office:
            bifa.state(["disabled"])                  # fără Drive nu se poate; fără Office e obligatoriu

        self.btn = tk.Button(self._jos, text="Previzualizare și calcul în Excel", font=("Helvetica", 12, "bold"),
                             command=self.previzualizeaza, **STIL_BUTON)
        self.btn.pack(pady=(8, 4), fill="x", padx=20)

        self.lbl_status = ttk.Label(self._jos, text="Aștept date...", foreground="gray")
        self.lbl_status.pack(pady=(0, 6))

        self.lbl_sursa = ttk.Label(root, text="", foreground="gray", wraplength=860)
        self.lbl_sursa.pack(pady=5)
        self._afiseaza_sursa()
        self._actualizeaza_liste_preia()
        ttk.Button(root, text="⚙ Setări generator", command=self.dialog_setari).pack(pady=(0, 8))
        self._actualizare = None
        threading.Thread(target=lambda: setattr(self, "_actualizare", cauta_actualizare()), daemon=True).start()
        if DRIVE is not None:                     # ofertele vechi din 'Oferte generate' -> foldere pe zile (în fundal)
            threading.Thread(target=self._organizeaza_pe_zile, daemon=True).start()
        self.root.after(1000, self._arata_actualizare)
        self.probleme_pricing = None
        if self.date:
            self.root.after(200, self._verifica_pricing)
        shutil.rmtree(os.path.join(FOLDER_OFERTE, FOLDER_TESTE), ignore_errors=True)   # resturi locale de probă
        self.root.protocol("WM_DELETE_WINDOW", self.la_inchidere)
        if DRIVE is not None:
            self.root.after(300, self.ia_nr_contract)

    def _branduri_skid(self):
        return branduri_skid(self.date)

    # ---------------- fereastră cu derulare ----------------
    def _zona_derulare(self, parinte):
        """Un cadru cu bară de derulare verticală care ocupă toată fereastra; întoarce cadrul interior."""
        canvas = tk.Canvas(parinte, highlightthickness=0)
        bara = ttk.Scrollbar(parinte, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=bara.set)
        bara.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        interior = ttk.Frame(canvas)
        fereastra = canvas.create_window((0, 0), window=interior, anchor="nw")
        interior.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(fereastra, width=e.width))

        def roti(e):
            if canvas.winfo_height() < interior.winfo_reqheight():
                pas = -1 if (getattr(e, "delta", 0) > 0 or getattr(e, "num", 0) == 4) else 1
                canvas.yview_scroll(pas * 3, "units")
        # rotița mouse-ului derulează doar cât timp cursorul e deasupra ferestrei principale
        canvas.bind("<Enter>", lambda e: (canvas.bind_all("<MouseWheel>", roti), canvas.bind_all("<Button-4>", roti),
                                          canvas.bind_all("<Button-5>", roti)))
        canvas.bind("<Leave>", lambda e: (canvas.unbind_all("<MouseWheel>"), canvas.unbind_all("<Button-4>"),
                                          canvas.unbind_all("<Button-5>")))
        self._canvas = canvas
        return interior

    # ---------------- verificarea calculului (Excel) înainte de PowerPoint / PDF ----------------
    def _construieste_tab_verificare(self, tab):
        self.lbl_verificare = ttk.Label(tab, foreground="gray", wraplength=1000, justify="left",
                                        text="Aici apare calculul din Excel după „Previzualizare și calcul în Excel”. "
                                             "Verifici prețurile (poți modifica și salva Excelul), apoi dai OK și abia "
                                             "atunci se fac PowerPoint-ul și PDF-ul.")
        self.lbl_verificare.pack(anchor="w", padx=10, pady=8)
        coloane = ("produs", "cant", "unit", "achizitie", "vanzare")
        self.tabel_verificare = ttk.Treeview(tab, columns=coloane, show="headings", height=12)
        for col, titlu, latime, aliniere in (("produs", "Produs", 380, "w"), ("cant", "Capacitate / putere", 140, "e"),
                                             ("unit", "Buc.", 60, "e"), ("achizitie", "Valoare achiziție DAP ($)", 190, "e"),
                                             ("vanzare", "Preț vânzare ($)", 170, "e")):
            self.tabel_verificare.heading(col, text=titlu)
            self.tabel_verificare.column(col, width=latime, anchor=aliniere)
        self.tabel_verificare.pack(fill="x", padx=10)
        self.lbl_rezumat = ttk.Label(tab, font=("Helvetica", 11, "bold"), justify="left")
        self.lbl_rezumat.pack(anchor="w", padx=10, pady=8)
        butoane = ttk.Frame(tab)
        butoane.pack(anchor="w", padx=10, pady=(0, 10))
        self.btn_deschide_excel = ttk.Button(butoane, text="📊 Deschide Excelul de calcul", state="disabled",
                                             command=self._deschide_excel_calcul)
        self.btn_deschide_excel.pack(side="left", padx=(0, 8))
        self.btn_reciteste = ttk.Button(butoane, text="↻ Recitește Excelul (după ce l-ai salvat)", state="disabled",
                                        command=self._arata_verificare)
        self.btn_reciteste.pack(side="left", padx=(0, 8))
        self.btn_ok = tk.Button(butoane, text="✅ OK – fă PowerPoint și PDF", font=("Helvetica", 11, "bold"),
                                state="disabled", command=self._ok_genereaza, **STIL_BUTON)
        self.btn_ok.pack(side="left", padx=(0, 8))
        self.btn_renunta = ttk.Button(butoane, text="🗑 Renunță și șterge calculul", state="disabled",
                                      command=self._renunta_calcul)
        self.btn_renunta.pack(side="left")

    def _stare_verificare(self, activ):
        for b in (self.btn_deschide_excel, self.btn_reciteste, self.btn_ok, self.btn_renunta):
            b.config(state="normal" if activ else "disabled")

    def _arata_verificare(self):
        """Tabelul din Excelul de calcul (recitit din fișierul salvat) + totalul și $/MWh."""
        c = self._calcul_in_asteptare
        if not c:
            return
        try:
            rez = citeste_rezultate(c["calcul"])
        except Exception as e:
            messagebox.showerror("Verificare", f"Nu pot citi Excelul de calcul:\n{e}\n\n"
                                 "Dacă e deschis și modificat, salvează-l (Ctrl+S).")
            return
        self.tabel_verificare.delete(*self.tabel_verificare.get_children())
        for rand in rez["tabel"][1:]:                     # primul rând = antetul din Excel
            produs = rand[0]
            if produs in (None, "") or not any(v not in (None, "") for v in (rand[6], rand[7])):
                continue
            fmt = lambda v: bani_eu(v, 0) if isinstance(v, (int, float)) else ("" if v is None else str(v))
            self.tabel_verificare.insert("", "end", values=(str(produs).strip(), fmt(rand[1]), fmt(rand[2]),
                                                             fmt(rand[6]), fmt(rand[7])))
        i = c["intrari"]
        self.lbl_rezumat.config(text=(
            f"{c['calcul']['nume']}\n"
            f"Preț total ({c['baterie']['incoterm']}): {bani_eu(rez['pret_ddp'])} $   |   "
            f"{bani_eu(rez['pret_mwh'])} $/MWh   |   cost achiziție: {bani_eu(rez['cost'])} $   |   "
            f"adaos {i['adaos']:g}\nEMS / an: {bani_eu(rez['ems_yearly'])} $   |   extra warranty baterie: "
            f"{bani_eu(rez['ext_gar_bat'])} $/MWh/an, PCS: {bani_eu(rez['ext_gar_pcs'])} $/MWh/an"))
        if c.get("test"):
            self.lbl_verificare.config(foreground="#b35c00", text=(
                "🧪 CALCUL DE PROBĂ (Client și Proiect = „test”): nu se păstrează nimic în Drive și nimic în "
                "BESS_Proposals. Se șterge la „Renunță”, la următorul calcul sau la închiderea programului; "
                "OK face doar un PDF de probă, pe care îl deschide local."))
            return
        self.lbl_verificare.config(foreground="black", text=(
            "Verifică prețurile. Dacă vrei să schimbi ceva, deschide Excelul (se deschide în browser, în Excel Online, "
            "și se salvează singur), modifică și apasă „Recitește”. La OK, PowerPoint-ul și PDF-ul iau valorile din "
            "Excelul din Drive." if c["calcul"].get("online") else
            "Verifică prețurile. Dacă vrei să schimbi ceva, deschide Excelul, modifică, SALVEAZĂ (Ctrl+S) și apasă "
            "„Recitește”. La OK, PowerPoint-ul și PDF-ul iau valorile din Excelul salvat."))

    def _deschide_excel_calcul(self):
        calcul = self._calcul_in_asteptare["calcul"]
        if calcul["rez"].get("link_excel"):            # Excelul din Drive, în Excel Online
            webbrowser.open(calcul["rez"]["link_excel"])
            return
        cale = calcul["cale_excel"]
        try:
            if sys.platform == "win32":
                os.startfile(cale)
            else:
                subprocess.Popen(["open", cale])
        except Exception as e:
            messagebox.showerror("Excel", f"Nu pot deschide fișierul:\n{cale}\n\n{e}")

    def _renunta_calcul(self):
        c = self._calcul_in_asteptare
        if not c or not messagebox.askyesno("Renunț", "Renunți la acest calcul? (nu se face PowerPoint / PDF și "
                                                       "nu se urcă nimic în Drive)"):
            return
        self._sterge_calcul(c)
        self._calcul_in_asteptare = None
        self._stare_verificare(False)
        self.tabel_verificare.delete(*self.tabel_verificare.get_children())
        self.lbl_rezumat.config(text="")
        self.status("Calcul anulat.", "gray")
        self.taburi.select(0)

    def _sterge_calcul(self, c, arata_erori=True):
        """Șterge un calcul: folderul local și, în modul Microsoft 365, folderul din Drive."""
        shutil.rmtree(c["calcul"]["folder"], ignore_errors=True)    # dacă Excelul e deschis, rămâne local
        if c["calcul"]["rez"].get("folder_drive"):                 # mod Microsoft 365: calculul e și în Drive
            try:
                DRIVE.sterge(c["calcul"]["rez"]["folder_drive"])
            except Exception as e:
                if arata_erori:
                    messagebox.showwarning("Ștergere", f"Nu am putut șterge calculul din Drive:\n{e}")

    def la_inchidere(self):
        """La închiderea programului: calculele de probă se șterg (local și din Drive)."""
        c = self._calcul_in_asteptare
        if c and c.get("test"):
            self._sterge_calcul(c, arata_erori=False)
        shutil.rmtree(os.path.join(FOLDER_OFERTE, FOLDER_TESTE), ignore_errors=True)
        self.root.destroy()

    def _ok_genereaza(self):
        c = self._calcul_in_asteptare
        if not c:
            return
        if not messagebox.askyesno("OK", "Ai verificat calculul? Fac PowerPoint-ul și PDF-ul cu valorile din "
                                         "Excelul salvat?"):
            return
        self._stare_verificare(False)
        try:
            self.genereaza(c)
            self._calcul_in_asteptare = None
            self.taburi.select(0)
        except Exception as e:
            self._stare_verificare(True)
            self.status("Eroare la generare.", "red")
            messagebox.showerror("Eroare", str(e))

    def _organizeaza_pe_zile(self):
        try:
            DRIVE.sterge_teste()                   # calcule de probă rămase (ex. program închis brusc)
        except Exception:
            pass
        try:
            DRIVE.organizeaza_oferte_pe_zile()
        except Exception:
            pass                                   # se reîncearcă la următoarea pornire

    # ---------------- actualizare automată ----------------
    def _arata_actualizare(self, incercari=60):
        if self._actualizare is None:                       # verificarea încă rulează (max. ~1 minut)
            if incercari:
                self.root.after(1000, lambda: self._arata_actualizare(incercari - 1))
            return
        if self._actualizare:
            tk.Button(self.continut, text="⬆ Versiune nouă disponibilă – Actualizează", bg="#1e7e34", fg="white",
                      font=("Helvetica", 10, "bold"), command=self.actualizeaza_programul)\
                .pack(pady=(0, 8), fill="x", padx=20)

    def actualizeaza_programul(self):
        if not messagebox.askyesno("Actualizare", "Instalez versiunea nouă și repornesc programul?\n\n"
                                   "(Ce ai completat în fereastră se pierde.)"):
            return
        try:
            aplica_actualizare(self._actualizare)
        except Exception as e:
            messagebox.showerror("Actualizare", f"Nu am putut actualiza (programul a rămas cum era):\n{e}")
            return
        reporneste_programul()
        self.root.destroy()

    # ---------------- verificarea Pricing și setările ----------------
    def _verifica_pricing(self, blocheaza=False):
        """Verifică coloanele din Pricing. self.probleme_pricing: [] = în regulă, listă = nu se generează."""
        if not self.date:
            return
        probleme, schimbate, actuale = verifica_coloane_pricing(self.date)
        if probleme:
            self.probleme_pricing = probleme
            if not blocheaza:
                messagebox.showerror("Pricing.xlsx – coloane schimbate",
                                     "Coloanele din Pricing nu mai sunt unde le caută programul, deci prețurile "
                                     "s-ar citi greșit. Nu se pot face oferte până nu se rezolvă:\n\n- "
                                     + "\n- ".join(probleme) + "\n\nVerifică Pricing.xlsx (sau anunță-mă să "
                                     "actualizez programul).")
            return
        if schimbate and not messagebox.askyesno(
                "Pricing.xlsx – antete schimbate",
                "Unele coloane din Pricing au alt titlu decât la ultima verificare:\n\n- " + "\n- ".join(schimbate)
                + "\n\nValorile arată în continuare corect (numere unde trebuie). Dacă e doar o redenumire, "
                  "apasă Da și noile titluri devin cele normale.\n\nNu = nu se fac oferte până nu verifici."):
            self.probleme_pricing = schimbate
            return
        accepta_antete_pricing(actuale)
        self.probleme_pricing = []

    def _aplica_setari(self):
        schimbate, probleme = aplica_setari_generator()
        if probleme:
            messagebox.showwarning(FISIER_SETARI_GENERATOR, "Probleme în fișierul de setări (rămân valorile din "
                                   "program):\n\n- " + "\n- ".join(probleme))
        if schimbate and self.date:
            self.combo_baterie["values"] = ([OPTIUNE_AUTO_BATERIE] + baterii_din_lista(self.baterii_map)
                                            + [OPTIUNE_MANUALA_BATERIE])
        return schimbate

    def dialog_setari(self):
        """Creează / deschide Setari generator.xlsx (în 'Ofertare BESS' din Drive) și reîncarcă setările."""
        exista = os.path.exists(FISIER_SETARI_GENERATOR)
        if DRIVE is not None:
            if exista:
                actualizate = DRIVE.actualizeaza()
                if FISIER_SETARI_GENERATOR in actualizate:
                    n = self._aplica_setari()
                    messagebox.showinfo("Setări", f"Am reîncărcat setările din Drive ({n} schimbate).")
                if messagebox.askyesno("Setări generator",
                                       f"Regulile sunt în '{FISIER_SETARI_GENERATOR}' din folderul Ofertare BESS "
                                       "(Drive).\n\nÎl deschid în browser ca să-l modifici?\n\nDupă salvare, "
                                       "setările noi se aplică la următoarea previzualizare."):
                    drive_id, item_id = DRIVE.ofertare
                    try:
                        el = DRIVE._cerere("GET", f"/drives/{drive_id}/items/{item_id}:/"
                                                  f"{_cale_url(FISIER_SETARI_GENERATOR)}")
                        webbrowser.open(el["webUrl"])
                    except Exception as e:
                        messagebox.showerror("Setări", f"Nu am putut deschide fișierul:\n{e}")
                return
            if not messagebox.askyesno("Setări generator",
                                       f"Nu există încă '{FISIER_SETARI_GENERATOR}' în Ofertare BESS.\n\n"
                                       "Îl creez acum, cu regulile folosite azi de program, și îl urc în Drive?"):
                return
            try:
                scrie_setari_generator(FISIER_SETARI_GENERATOR)
                link = DRIVE.urca_setari(FISIER_SETARI_GENERATOR)
            except Exception as e:
                messagebox.showerror("Setări", f"Nu am putut crea fișierul:\n{e}")
                return
            if messagebox.askyesno("Setări generator", "Gata, fișierul e în Ofertare BESS. Îl deschid în browser?"):
                webbrowser.open(link)
            return
        # fără Drive: fișierul stă în folderul sincronizat 'Ofertare BESS'
        if not exista:
            if not messagebox.askyesno("Setări generator", f"Creez '{FISIER_SETARI_GENERATOR}' în folderul "
                                       f"Ofertare BESS, cu regulile folosite azi de program?"):
                return
            scrie_setari_generator(FISIER_SETARI_GENERATOR)
        else:
            n = self._aplica_setari()
            messagebox.showinfo("Setări", f"Am reîncărcat setările ({n} schimbate).")
        if sys.platform == "win32":
            os.startfile(os.path.abspath(FISIER_SETARI_GENERATOR))

    def _construieste_tab_manual(self, tab):
        """Câmpurile pentru bateria și skidul introduse manual (două coloane)."""
        ttk.Label(tab, foreground="gray", wraplength=850, justify="left",
                  text="Completează doar componentele care nu sunt în Pricing (cotații primite de la furnizori). "
                       "Apoi, în tab-ul Ofertă, alege „✎ Baterie manuală” la Model baterie și/sau „✎ Skid manual” "
                       "la Brand skid. Se pot combina cu componentele din baza de date.")\
            .grid(row=0, column=0, columnspan=2, padx=10, pady=8, sticky="w")
        self.manual = {"bat": {}, "skid": {}}
        self.combo_preia, self._skiduri_preia = {}, {}
        self._construieste_rezerva(tab)
        for coloana, (cheie, titlu, campuri) in enumerate([("bat", "Baterie manuală", CAMPURI_BATERIE_MANUALA),
                                                           ("skid", "Skid manual", CAMPURI_SKID_MANUAL)]):
            grup = ttk.LabelFrame(tab, text=titlu)
            grup.grid(row=2, column=coloana, padx=10, pady=5, sticky="n")
            # completare rapidă din Pricing: se aleg brandul / modelul, apoi se modifică doar ce diferă (ex. prețul)
            ttk.Label(grup, text="Preia din Pricing:", font=("Helvetica", 10, "bold"))\
                .grid(row=0, column=0, sticky="w", padx=5, pady=(6, 2))
            combo = ttk.Combobox(grup, state="readonly", width=30)
            combo.grid(row=0, column=1, padx=5, pady=(6, 2))
            combo.bind("<<ComboboxSelected>>", lambda e, k=cheie: self._preia_din_pricing(k))
            self.combo_preia[cheie] = combo
            ttk.Separator(grup).grid(row=1, column=0, columnspan=2, sticky="we", padx=5, pady=6)
            for i, (camp, eticheta, implicit) in enumerate(campuri, start=2):
                if camp is None:
                    ttk.Label(grup, text=eticheta, font=("Helvetica", 10, "bold"))\
                        .grid(row=i, column=0, columnspan=2, sticky="w", padx=5, pady=(10, 2))
                elif camp.endswith("_inclus"):
                    var = tk.BooleanVar()
                    ttk.Checkbutton(grup, text=eticheta, variable=var)\
                        .grid(row=i, column=0, columnspan=2, sticky="w", padx=5, pady=3)
                    self.manual[cheie][camp] = var
                else:
                    ttk.Label(grup, text=eticheta).grid(row=i, column=0, sticky="w", padx=5, pady=3)
                    e = ttk.Entry(grup, width=18)
                    e.insert(0, implicit)
                    e.grid(row=i, column=1, padx=5, pady=3)
                    self.manual[cheie][camp] = e

    def _construieste_rezerva(self, tab):
        """Rezervă de preț (%) peste prețul bateriei / skidului, pentru componentele manuale și cele din Pricing."""
        grup = ttk.LabelFrame(tab, text="Rezervă de preț (în caz că se schimbă prețurile) – se aplică și "
                                         "componentelor din Pricing")
        grup.grid(row=1, column=0, columnspan=2, padx=10, pady=(0, 6), sticky="w")
        self.rezerva = {}
        for i, (cheie, eticheta) in enumerate((("bat", "Baterie (+%):"), ("skid", "Skid (+%):"))):
            ttk.Label(grup, text=eticheta).grid(row=0, column=2 * i, padx=(8, 4), pady=6)
            e = ttk.Entry(grup, width=8)
            e.grid(row=0, column=2 * i + 1, padx=(0, 12), pady=6)
            self.rezerva[cheie] = e

        def seteaza(valoare):
            for e in self.rezerva.values():
                e.delete(0, "end")
                e.insert(0, valoare)
        ttk.Button(grup, text="+10% la ambele", command=lambda: seteaza("10")).grid(row=0, column=4, padx=4)
        ttk.Button(grup, text="Fără rezervă", command=lambda: seteaza("")).grid(row=0, column=5, padx=4)
        ttk.Label(grup, foreground="gray", text="Prețul cu rezervă apare în Excel ca formulă (ex. =85000*(1+10%)).")\
            .grid(row=1, column=0, columnspan=6, sticky="w", padx=8, pady=(0, 6))

    # ---------------- citire ATR ----------------
    def _construieste_tab_atr(self, tab):
        ttk.Label(tab, foreground="gray", wraplength=1000, justify="left",
                  text="Alege un ATR (aviz tehnic de racordare) sau un studiu de soluție în PDF. Programul citește "
                       "puterea aprobată, instalația de stocare, punctul de racordare și trafo-urile, apoi poți "
                       "completa oferta cu aceste valori. Documentele scanate (poze) nu se pot citi.")\
            .pack(anchor="w", padx=10, pady=8)
        butoane = ttk.Frame(tab)
        butoane.pack(anchor="w", padx=10)
        ttk.Button(butoane, text="📄 Alege ATR (PDF)…", command=self.citeste_atr).pack(side="left", padx=(0, 8))
        self.btn_completeaza_atr = tk.Button(butoane, text="➡ Completează oferta cu aceste valori", state="disabled",
                                             command=self._completeaza_din_atr, **STIL_BUTON)
        self.btn_completeaza_atr.pack(side="left")
        cadru = ttk.Frame(tab)
        cadru.pack(fill="both", expand=True, padx=10, pady=8)
        self.txt_atr = tk.Text(cadru, height=24, wrap="word", font=("Consolas", 10))
        bara = ttk.Scrollbar(cadru, command=self.txt_atr.yview)
        self.txt_atr.configure(yscrollcommand=bara.set, state="disabled")
        bara.pack(side="right", fill="y")
        self.txt_atr.pack(side="left", fill="both", expand=True)
        self._atr = None

    def citeste_atr(self):
        cale = filedialog.askopenfilename(title="Alege ATR-ul (PDF)", filetypes=[("PDF", "*.pdf")])
        if not cale:
            return
        try:
            text = text_pdf(cale)
        except RuntimeError as e:
            if "pypdf" in str(e) and messagebox.askyesno(
                    "Citire ATR", "Pentru citirea PDF-urilor trebuie instalată o dată biblioteca pypdf.\n\n"
                                  "O instalez acum? (durează câteva secunde, e nevoie de internet)"):
                rez = subprocess.run([sys.executable.replace("pythonw", "python"), "-m", "pip", "install", "--user",
                                      "pypdf"], capture_output=True, text=True)
                if rez.returncode != 0:
                    messagebox.showerror("Citire ATR", f"Instalarea nu a mers:\n{rez.stderr[-800:]}")
                    return
                import importlib
                import site
                site.addsitedir(site.getusersitepackages())       # folderul --user poate fi nou
                importlib.invalidate_caches()
                text = text_pdf(cale)
            else:
                return
        except Exception as e:
            messagebox.showerror("Citire ATR", f"Nu pot citi PDF-ul:\n{e}")
            return
        r = citeste_atr_text(text)
        r["fisier"] = os.path.basename(cale)
        self._atr = r
        self._arata_atr(r)
        self.btn_completeaza_atr.config(state="disabled" if r.get("scanat") or not r.get("mw") else "normal")

    def _arata_atr(self, r):
        kw = lambda v: f"{nr_eu(v)} kW" if v is not None else "-"
        linii = [f"📄 {r.get('fisier', '')}", ""]
        if not r.get("scanat"):
            linii += [
                f"Tip document:          {r.get('tip', '-')}",
                f"Operator:              {r.get('operator', '-')}",
                f"Nr. / data:            {r.get('nr', '-')} / {r.get('data', '-')}",
                f"Beneficiar:            {r.get('beneficiar') or '-'}",
                f"Obiectiv:              {r.get('obiectiv') or '-'}",
                f"Amplasament:           {r.get('amplasament') or '-'}",
                f"Punct de racordare:    {nr_eu(r['poc_kv']) + ' kV' if r.get('poc_kv') else '-'}",
                "",
                "PUTERE APROBATĂ",
                f"  evacuare în rețea:   {kw(r.get('evac_kw'))}" + (f" / {nr_eu(r['evac_kva'])} kVA" if r.get("evac_kva") else ""),
                f"  absorbție din rețea: {kw(r.get('abs_kw'))}",
                "",
                "INSTALAȚIA DE STOCARE" + ("   ⚠ NU EXISTĂ ÎN ACEST ATR" if r.get("fara_stocare") else
                                           "   ⚠ ATR DE CONSUMATOR (fără evacuare)" if r.get("consumator") else ""),
                f"  Pi stocare (tabel):  {kw(r.get('pi_is_kw'))}" + (f"   (în text: {kw(r.get('pi_text_kw'))})" if r.get("pi_text_kw") else ""),
                f"  Pmax evacuare IS:    {kw(r.get('evac_is_kw'))}",
                f"  Pmax absorbție IS:   {kw(r.get('abs_is_kw'))}",
                f"  capacitate:          " + (f"{nr_eu(r['mwh'])} MWh" if r.get("mwh") else "-")
                + (f"   ({nr_eu(r['capacitate_ah'])} Ah)" if r.get("capacitate_ah") else ""),
                f"  cuplare pe DC:       {'DA' if r.get('cuplare_dc') else 'nu'}",
                f"  legată pe JT (TDRI): {'DA (prin trafo-urile existente)' if r.get('cuplare_jt') else 'nu'}",
                f"  fără absorbție:      {'DA' if r.get('fara_absorbtie') else 'nu'}",
                "",
                "TRAFO-URI ÎN INSTALAȚIE: " + (", ".join(f"{nr_eu(k)} kVA ({rap})" for k, rap in r["trafo_kva"])
                                               or "-"),
                "",
                "PROPUNERE PENTRU OFERTĂ",
                f"  Putere (MW) = evacuarea aprobată:        {nr_eu(r['mw']) if r.get('mw') else '-'}",
                f"  Putere max. admisă (MW) = Pi stocare:    {nr_eu(r['mw_max']) if r.get('mw_max') else '-'}",
                f"  Capacitate (MWh):                        {nr_eu(r['mwh']) if r.get('mwh') else 'din cererea clientului'}",
                "  Skidul poate fi mai mare decât evacuarea (limitat prin EMS), dar nu peste puterea max. admisă.",
            ]
        if r["avertismente"]:
            linii += ["", "DE VERIFICAT"] + [f"  ⚠ {a}" for a in r["avertismente"]]
        self.txt_atr.configure(state="normal")
        self.txt_atr.delete("1.0", "end")
        self.txt_atr.insert("1.0", "\n".join(linii))
        self.txt_atr.configure(state="disabled")

    def _completeaza_din_atr(self):
        r = self._atr
        if not r:
            return

        def pune(eticheta, valoare):
            if valoare not in (None, ""):
                self.intrari[eticheta].delete(0, "end")
                self.intrari[eticheta].insert(0, valoare)
        pune("Putere (MW):", nr_eu(r["mw"]).replace(".", "") if r.get("mw") else None)
        pune("Putere max. admisă (MW, opțional):", nr_eu(r["mw_max"]).replace(".", "") if r.get("mw_max") else "")
        if r.get("mwh"):
            pune("Capacitate (MWh):", nr_eu(r["mwh"]).replace(".", ""))
        if r.get("beneficiar"):
            pune("Client:", r["beneficiar"])
        if r.get("amplasament"):
            pune("Locație:", re.sub(r"(?i)^jude[țţt]ul\s+", "", r["amplasament"]).replace(", -", ",").strip()
                 + ", Romania")
        if r.get("poc_kv") and r["poc_kv"] <= 35:
            valori = list(self.combo_poc["values"])
            potrivit = next((v for v in valori if abs(float(re.match(r"\s*(\d+(?:[.,]\d+)?)", v).group(1)
                                                                 .replace(",", ".")) - r["poc_kv"]) < 1e-6), None)
            if potrivit:
                self.combo_poc.set(potrivit)
        self.taburi.select(self.tab_oferta)
        messagebox.showinfo("Citire ATR", "Am completat: Putere, Putere max. admisă"
                            + (", Capacitate" if r.get("mwh") else "") + ", Client, Locație și POC.\n\n"
                            "Verifică valorile (mai ales cele cu ⚠ din tab-ul ATR) și completează ce lipsește "
                            "(ex. capacitatea din cererea clientului).")

    def _actualizeaza_liste_preia(self):
        """Listele „Preia din Pricing” (toate bateriile și skidurile din Pricing)."""
        if not self.date or not hasattr(self, "combo_preia"):
            return
        self.combo_preia["bat"]["values"] = list(self.baterii_map)
        self._skiduri_preia = skiduri_din_pricing(self.date)
        self.combo_preia["skid"]["values"] = list(self._skiduri_preia)

    def _preia_din_pricing(self, cheie):
        """Completează câmpurile bateriei / skidului manual cu datele din Pricing."""
        alegere = self.combo_preia[cheie].get()
        try:
            if cheie == "bat":
                valori = baterie_din_pricing(self.date, self.baterii_map, alegere)
            else:
                valori = skid_din_pricing(self.date, *self._skiduri_preia[alegere])
        except Exception as e:
            messagebox.showerror("Preia din Pricing", f"Nu am putut prelua datele:\n{e}")
            return
        for camp, widget in self.manual[cheie].items():
            if isinstance(widget, tk.BooleanVar):
                widget.set(False)                                  # transportul e separat în Pricing
                continue
            v = valori.get(camp, "")
            text = (f"{v:.4f}".rstrip("0").rstrip(".").replace(".", ",") if isinstance(v, float) else str(v)) \
                if v not in (None, "") else ""          # virgulă zecimală: "75000,5" nu se confundă cu mii
            widget.delete(0, "end")
            widget.insert(0, "" if text in ("0", "nan") and camp in ("ext_gar", "kva_aux", "aux_kva") else text)

    # ---------------- tabelul BESS_Proposals ----------------
    def _construieste_panou_tabel(self, parinte):
        grup = ttk.LabelFrame(parinte, text="Date pentru tabelul BESS_Proposals")
        grup.pack(side="left", padx=10, pady=10, anchor="n")
        ttk.Label(grup, text="Țară (foaia din tabel):").grid(row=0, column=0, sticky="w", padx=5, pady=4)
        self.combo_tara = ttk.Combobox(grup, values=list(TABELE_OFERTE), state="readonly", width=18)
        self.combo_tara.set("Romania")
        self.combo_tara.grid(row=0, column=1, padx=5, pady=4)
        self.combo_tara.bind("<<ComboboxSelected>>", lambda e: self.ia_nr_contract())
        ttk.Label(grup, text="Tip proiect:").grid(row=1, column=0, sticky="w", padx=5, pady=4)
        self.combo_tip = ttk.Combobox(grup, values=TIPURI_PROIECT, state="readonly", width=18)
        self.combo_tip.set(TIPURI_PROIECT[0])
        self.combo_tip.grid(row=1, column=1, padx=5, pady=4)
        ttk.Label(grup, text="Responsabil:").grid(row=2, column=0, sticky="w", padx=5, pady=4)
        self.txt_responsabil = ttk.Entry(grup, width=21)
        self.txt_responsabil.insert(0, citeste_setari().get("responsabil", ""))
        self.txt_responsabil.grid(row=2, column=1, padx=5, pady=4)
        ttk.Label(grup, text="Data (zz.ll.aaaa):").grid(row=3, column=0, sticky="w", padx=5, pady=4)
        self.txt_data_cerere = ttk.Entry(grup, width=21)
        self.txt_data_cerere.insert(0, datetime.now().strftime("%d.%m.%Y"))
        self.txt_data_cerere.grid(row=3, column=1, padx=5, pady=4)
        tk.Button(grup, text="↻ Ia nr. de contract liber din tabel", command=self.ia_nr_contract)\
            .grid(row=4, column=0, columnspan=2, sticky="we", padx=5, pady=(8, 4))
        self.lbl_tabel = ttk.Label(grup, text="", foreground="gray", wraplength=300, justify="left")
        self.lbl_tabel.grid(row=5, column=0, columnspan=2, sticky="w", padx=5, pady=4)
        ttk.Label(grup, foreground="gray", wraplength=300, justify="left",
                  text="Dacă „Nr. proiect” e gol, nr. de contract se ia automat din tabel. Generarea NU scrie "
                       "nimic în tabel: după ce verifici oferta, apasă butonul de mai jos.")\
            .grid(row=6, column=0, columnspan=2, sticky="w", padx=5, pady=4)
        tk.Button(grup, text="📋 Trece o ofertă verificată în BESS_Proposals…", command=self.dialog_tabel,
                  **STIL_BUTON).grid(row=7, column=0, columnspan=2, sticky="we", padx=5, pady=(4, 8))

    def ia_nr_contract(self):
        """Pune în 'Nr. proiect' primul nr. de contract liber din foaia țării alese."""
        if DRIVE is None:
            self.lbl_tabel.config(text="Drive-ul nu e conectat: nr. de contract se scrie manual.")
            return
        try:
            self.status("Caut nr. de contract liber în BESS_Proposals...")
            nr, rand_nou = nr_contract_liber(self.combo_tara.get())
        except Exception as e:
            self.lbl_tabel.config(text=f"Nu am putut citi BESS_Proposals: {e}", foreground="red")
            self.status("Aștept date...", "gray")
            return
        self.status("Aștept date...", "gray")
        foaie = TABELE_OFERTE[self.combo_tara.get()]["foaie"]
        if not nr:
            self.lbl_tabel.config(text=f"Nu găsesc niciun nr. de contract ST în foaia '{foaie}'.", foreground="red")
            return
        camp = self.intrari["Nr. proiect:"]
        camp.delete(0, "end")
        camp.insert(0, nr)
        self.lbl_tabel.config(text=f"Nr. de contract: {nr} (foaia '{foaie}'" +
                                   (", rând nou după ultimul din tabel)" if rand_nou else ", rând pregătit în tabel)"),
                              foreground="green")

    def _verifica_nr_contract(self, intrari):
        """Înainte de generare: nr. de contract există în tabel și e liber (sau al aceluiași client)?"""
        if DRIVE is None or e_calcul_test(intrari):
            return True
        tara = self.combo_tara.get()
        try:
            gasit = next(((r, c) for r, nr, c in contracte_din_tabel(tara) if nr == intrari["nr_proiect"]), None)
        except Exception as e:
            return messagebox.askyesno("Atenție", f"Nu am putut verifica nr. de contract în BESS_Proposals:\n{e}"
                                                  "\n\nContinui oricum?")
        foaie = TABELE_OFERTE[tara]["foaie"]
        if gasit is None:                    # nr. nou (va primi un rând nou la trecerea în tabel)
            return True
        rand, client = gasit
        if client and client.strip().lower() != intrari["client"].strip().lower():
            return messagebox.askyesno(
                "Nr. de contract", f"Nr. {intrari['nr_proiect']} e deja folosit în foaia '{foaie}' (rândul {rand}) "
                f"pentru clientul „{client}”.\n\nContinui cu acest număr?")
        return True

    def _data_cerere(self):
        text = self.txt_data_cerere.get().strip()
        for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(text, fmt).strftime("%Y-%m-%d")
            except ValueError:
                pass
        raise ValueError(f"Data '{text}' nu e validă. Scrie-o ca zz.ll.aaaa (ex. 02.11.2026).")

    def _inregistrare_tabel(self, intrari, baterie, skid, r):
        """Datele ofertei pentru BESS_Proposals (salvate în lista 'de trecut în tabel')."""
        salveaza_setari(dict(citeste_setari(), responsabil=self.txt_responsabil.get().strip()))
        return {
            "nume": r["nume"], "creat": datetime.now().strftime("%d.%m.%Y %H:%M"),
            "tara": self.combo_tara.get(), "nr": intrari["nr_proiect"],
            "data_cerere": self._data_cerere(), "data_oferta": datetime.now().strftime("%Y-%m-%d"),
            "client": intrari["client"], "proiect": intrari["proiect"], "locatie": intrari["locatie"],
            "tip": self.combo_tip.get(), "responsabil": self.txt_responsabil.get().strip(),
            "brand_skid": skid["brand"], "mw": intrari["mw"], "mwh": intrari["mwh"],
            "brand_baterie": baterie["brand"], "mwh_container": baterie["cap"],
            "c_rate": nr_eu(round(intrari["mw"] / intrari["mwh"], 2)) + "C",
            "pret_total": round(r["rez"]["pret_ddp"], 2), "pret_mwh": round(r["rez"]["pret_mwh"], 2),
            "cost": round(r["rez"].get("cost", 0), 2), "po": "W",
        }

    def dialog_tabel(self):
        """Lista ofertelor generate și netrecute; cea aleasă se trece în BESS_Proposals."""
        if DRIVE is None:
            messagebox.showwarning("BESS_Proposals", "Tabelul se poate completa doar cu Drive-ul conectat.")
            return
        win = tk.Toplevel(self.root)
        win.title("Trece în BESS_Proposals")
        win.transient(self.root)
        win.grab_set()
        ttk.Label(win, text="Oferte generate și încă netrecute în tabel (alege una, după ce ai verificat-o):",
                  font=("Helvetica", 10, "bold")).pack(padx=10, pady=8, anchor="w")
        lista = tk.Listbox(win, width=120, height=12, font=("Consolas", 9))
        lista.pack(padx=10, fill="both", expand=True)

        def reincarca():
            lista.delete(0, "end")
            for o in oferte_netrecute():
                lista.insert("end", f"{o['nr']:<14} {o['tara']:<8} {o['client'][:18]:<18} "
                                    f"{bani_eu(o['pret_total'])} $   {o['nume']}   (generată {o['creat']})")

        def aleasa():
            if not lista.curselection():
                messagebox.showinfo("BESS_Proposals", "Alege o ofertă din listă.", parent=win)
                return None
            return lista.curselection()[0]

        def trece():
            i = aleasa()
            if i is None:
                return
            o = oferte_netrecute()[i]
            cfg = TABELE_OFERTE[o["tara"]]
            rezumat = "\n".join(f"{col}: {_valoare_text(o, camp)}" for col, camp in cfg["coloane"].items())
            if not messagebox.askyesno("Confirmare", f"Trec în foaia '{cfg['foaie']}', pe rândul cu nr. de contract "
                                                     f"{o['nr']}:\n\n{rezumat}\n\nContinui?", parent=win):
                return
            try:
                rand = trece_in_tabel(o, lambda r, c: messagebox.askyesno(
                    "Rând deja completat", f"Rândul {r} (nr. {o['nr']}) are deja clientul „{c}”.\n\n"
                    "Îl suprascriu cu oferta aceasta (ex. o versiune nouă)?", parent=win))
            except Exception as e:
                messagebox.showerror("Eroare", str(e), parent=win)
                return
            if rand is None:
                return
            ramase = oferte_netrecute()
            del ramase[i]
            salveaza_oferte_netrecute(ramase)
            reincarca()
            messagebox.showinfo("Gata", f"Oferta {o['nr']} a fost trecută în '{cfg['foaie']}', rândul {rand}.",
                                parent=win)
            self.ia_nr_contract()

        def sterge():
            i = aleasa()
            if i is None:
                return
            ramase = oferte_netrecute()
            if messagebox.askyesno("Șterge", f"Scot din listă oferta {ramase[i]['nume']}?\n(Tabelul nu se "
                                             "modifică; fișierele ofertei rămân în Drive.)", parent=win):
                del ramase[i]
                salveaza_oferte_netrecute(ramase)
                reincarca()

        bf = ttk.Frame(win)
        bf.pack(pady=10)
        tk.Button(bf, text="Trece în tabel", command=trece, **STIL_BUTON).pack(side="left", padx=6)
        tk.Button(bf, text="Șterge din listă (test)", command=sterge).pack(side="left", padx=6)
        tk.Button(bf, text="Închide", command=win.destroy).pack(side="left", padx=6)
        reincarca()

    def _mod_trafo(self):
        """'fara' / True (forțat) / False (automat), după cele două bife."""
        return "fara" if self.fara_aux.get() else bool(self.foloseste_aux.get())

    def citeste_manual(self, cheie):
        return {camp: (w.get() if isinstance(w, tk.BooleanVar) else w.get().strip())
                for camp, w in self.manual[cheie].items()}

    def _valori_poc(self):
        valori = [f"{POC_IMPLICIT_KV} kV (România)"]
        if self.date:
            valori += [f"{kv:g} kV" for kv in tensiuni_poc(self.date.get("skid_poc")) if kv != POC_IMPLICIT_KV]
        return valori

    def _afiseaza_sursa(self):
        if DRIVE is not None:
            self.lbl_sursa.config(text=DRIVE.descriere_sursa())
        elif self.date:
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
            self.ems_var.set(f"{ems_pentru(mwh):g}")
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
                "ems_auto": not self.ems_manual,
                "eng": float(g("Valoare ENG (USD):").replace(",", ".")),
                "adaos": float(g("Adaos Comercial (Margin):").replace(",", ".")),
                "curs": float(g("Curs USD/EUR:").replace(",", ".")),
                "poc_kv": float(re.match(r"\s*(\d+(?:[.,]\d+)?)", self.combo_poc.get()).group(1).replace(",", ".")),
                "mw_max": float(g("Putere max. admisă (MW, opțional):").replace(",", ".") or 0),
                "rezerva_bat": float(self.rezerva["bat"].get().replace(",", ".").replace("%", "") or 0),
                "rezerva_skid": float(self.rezerva["skid"].get().replace(",", ".").replace("%", "") or 0),
            }
        except (ValueError, AttributeError):
            raise ValueError("Unul dintre câmpurile numerice (sau punctul de conexiune) nu conține un număr valid.")

    def previzualizeaza(self):
        try:
            self.status("Verific datele...")
            nr_scris = self.intrari["Nr. proiect:"].get().strip()
            test = all(self.intrari[k].get().strip().lower() == TEXT_TEST for k in ("Client:", "Nume Proiect:"))
            if test and (not nr_scris or "XXX" in nr_scris.upper()):
                self.intrari["Nr. proiect:"].delete(0, "end")          # calcul de probă: fără nr. din tabel
                self.intrari["Nr. proiect:"].insert(0, NR_PROIECT_TEST)
            elif DRIVE is not None and (not nr_scris or "XXX" in nr_scris.upper()):
                self.ia_nr_contract()             # nr. de contract automat din BESS_Proposals
            intrari = self.citeste_intrari()
            erori, avert = valideaza(intrari)
            if erori:
                raise ValueError("\n".join(erori))
            if avert and not messagebox.askyesno("Atenție", "\n".join(avert) + "\n\nContinui oricum?"):
                self.status("Anulat.", "gray")
                return

            # Cu Drive-ul conectat: descărcăm fișierele schimbate (și atunci recitim tot)
            if DRIVE is not None:
                self.status("Verific fișierele din Drive...")
                actualizate = DRIVE.actualizeaza()
                if actualizate:
                    self.date = None
                if FISIER_SETARI_GENERATOR in actualizate:
                    self._aplica_setari()
            if self.probleme_pricing is not None and not pricing_modificat(self.date):
                self._verifica_pricing(blocheaza=True)
            # Recitim Pricing.xlsx doar dacă s-a modificat în Drive de la ultima citire
            if pricing_modificat(self.date):
                self.status("Pricing.xlsx s-a modificat, îl recitesc...")
                self.date = incarca_date()
                self.baterii_map = lista_baterii(self.date["bat"])
                self.combo_baterie["values"] = ([OPTIUNE_AUTO_BATERIE] + baterii_din_lista(self.baterii_map)
                                                + [OPTIUNE_MANUALA_BATERIE])
                self.combo_brand["values"] = ["Oricare"] + self._branduri_skid() + [OPTIUNE_MANUALA_SKID]
                self.combo_poc["values"] = self._valori_poc()
                self._afiseaza_sursa()
                self._actualizeaza_liste_preia()
                self.probleme_pricing = None
                self._verifica_pricing()
            if self.probleme_pricing:
                raise ValueError("Pricing.xlsx are coloane mutate sau schimbate, prețurile s-ar citi greșit:\n\n- "
                                 + "\n- ".join(self.probleme_pricing))
            date = self.date

            nr_cerut = self.combo_nr_skid.get()
            p = pregateste_oferta(date, self.baterii_map, intrari, self.combo_baterie.get(),
                                  self.combo_incoterm.get(), self.combo_brand.get(),
                                  None if nr_cerut == AUTO_SKID else int(nr_cerut), self._mod_trafo(),
                                  manual_bat=self.citeste_manual("bat"), manual_skid=self.citeste_manual("skid"))
            baterie, nr_baterii = p["baterie"], p["nr_baterii"]
            optiuni, motive, lista, exact = p["optiuni"], p["motive"], p["lista"], p["exact"]
            if p["specs_lipsa"] and not messagebox.askyesno(
                    "Atenție", f"{p['specs_lipsa']}, deci Configuration și Weight vor rămâne goale "
                               "în prezentare.\n\nContinui oricum?"):
                self.status("Anulat.", "gray")
                return

            ems_nou = ajusteaza_ems(intrari, p)
            if ems_nou is not None:                    # EMS după capacitatea ofertată; câmpul arată valoarea nouă
                self._setare_auto = True
                try:
                    self.ems_var.set(f"{ems_nou:g}")
                finally:
                    self._setare_auto = False
            self._auto_info = (p.get("auto_comparatie"), p.get("auto_excluse"))
            alegere = self.dialog_previzualizare(intrari, baterie, nr_baterii, lista, exact,
                                                 raport_skiduri(optiuni, motive))
            if alegere is None:
                self.status("Anulat.", "gray")
                return
            if not self._verifica_nr_contract(intrari):
                self.status("Anulat.", "gray")
                return
            if lista[alegere].get("manual") and not messagebox.askyesno(
                    "Verificare compatibilitate",
                    "Skidul e introdus manual, iar programul NU a verificat compatibilitatea lui cu bateriile "
                    f"({nr_baterii} x {baterie['tip']}): intrări DC, tensiune, putere.\n\n"
                    "Ai verificat compatibilitatea componentelor? Generez oferta?"):
                self.status("Anulat.", "gray")
                return
            self.calculeaza(intrari, baterie, nr_baterii, lista[alegere], lista[alegere]["trafo"])

        except Exception as e:
            self.status("Eroare!", "red")
            messagebox.showerror("Eroare", str(e))

    def dialog_previzualizare(self, intrari, baterie, nr_baterii, optiuni, exact=True, raport=""):
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
               + (f"\nPutere max. admisă: {nr_eu(intrari['mw_max'])} MW (skidurile peste ea nu se propun)"
                  if intrari.get("mw_max") else "")
               + (f"\nRezervă de preț: baterie +{intrari.get('rezerva_bat') or 0:g}%, skid +{intrari.get('rezerva_skid') or 0:g}%"
                  if intrari.get("rezerva_bat") or intrari.get("rezerva_skid") else "")
               + f"\nEMS: {bani_eu(intrari['ems'], 0)} $"
               + (f" (automat după capacitatea ofertată {cap_prop:.3f} MWh, prag {EMS_PRAG_MWH:g} MWh)"
                  if intrari.get("ems_auto") else " (scris de mână)")
               + (f"\nPreț {baterie['incoterm']}: {baterie['pret']:,.2f} $/MWh (✎ baterie introdusă manual, "
                  f"{baterie['brand']}) -> baterii: {cost_bat:,.0f} $" if baterie.get("manual") else
                  f"\nPreț {baterie['incoterm']} (coloana {baterie['coloana_pret']}): "
                  f"{baterie['pret']:,.2f} $/MWh (cotație {baterie['data']}, "
                  f"rând Excel {baterie['rand_excel']}) -> baterii: {cost_bat:,.0f} $"))
        kva_bat = baterie.get("kva_aux") or kva_aux_baterie(baterie["cap"])
        txt += (f"\nConsum auxiliar: {nr_baterii} baterii x {kva_bat:g} kVA + "
                "consum skid după putere (vezi la fiecare skid); "
                + {"fara": "trafo aux extern EXCLUS (bifa „Fără trafo aux extern”)",
                   True: "trafo aux extern forțat la toate opțiunile"}.get(
                    self._mod_trafo(), "trafo-ul aux extern se adaugă automat unde skidul nu acoperă"))
        comparate, excluse = getattr(self, "_auto_info", (None, None))
        if comparate:
            txt += ("\n\nAuto - baterii comparate (cost achiziție, cea mai ieftină prima): " + "; ".join(comparate)
                    + ("\nExcluse: " + "; ".join(excluse) if excluse else ""))
        ttk.Label(win, text=txt, justify="left", wraplength=900).pack(padx=15, pady=10, anchor="w")

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
        tk.Button(bf, text="Generează", command=ok, **STIL_BUTON).pack(side="left", padx=8)
        tk.Button(bf, text="Vezi toate skidurile", command=arata_toate).pack(side="left", padx=8)
        tk.Button(bf, text="Anulează", command=win.destroy).pack(side="left", padx=8)
        self.root.wait_window(win)
        return rezultat["v"]

    def calculeaza(self, intrari, baterie, nr_baterii, skid, trafo):
        """Pasul 1: Excelul de calcul; apoi tab-ul „Verificare calcul” (PowerPoint / PDF abia după OK)."""
        vechi = self._calcul_in_asteptare
        if vechi and not vechi.get("test") and not messagebox.askyesno(
                "Calcul în așteptare", f"Calculul „{vechi['calcul']['nume']}” nu a primit încă "
                                       "OK. Îl înlocuiesc cu cel nou?"):
            self.status("Anulat.", "gray")
            return
        if vechi and vechi.get("test"):
            self._sterge_calcul(vechi)                 # calculul de probă anterior nu se păstrează
        calcul = calculeaza_oferta(self.date, intrari, baterie, nr_baterii, skid, trafo, status=self.status,
                                   online=self.mod_online.get(),
                                   nume_existente=DRIVE.nume_oferte() if DRIVE is not None else ())
        self._calcul_in_asteptare = {"calcul": calcul, "intrari": intrari, "baterie": baterie,
                                     "nr_baterii": nr_baterii, "skid": skid, "trafo": trafo,
                                     "test": e_calcul_test(intrari)}
        self._stare_verificare(True)
        self._arata_verificare()
        self.taburi.select(self.tab_verificare)
        self._canvas.yview_moveto(0)
        self.status("Calcul gata: verifică prețurile în tab-ul „Verificare calcul”, apoi dă OK.", "green")

    def _termina_test(self, c, r):
        """Calcul de probă: PDF-ul se deschide local; din Drive se șterge tot, nimic nu ajunge în BESS_Proposals.
        Fișierele locale se șterg la închiderea programului."""
        if c["calcul"]["rez"].get("folder_drive"):
            try:
                DRIVE.sterge(c["calcul"]["rez"]["folder_drive"])
            except Exception:
                pass
        self.status("Calcul de probă gata (nu s-a salvat nimic în Drive).", "green")
        deschide = r["cale_pdf"] if r.get("cale_pdf") and os.path.exists(r["cale_pdf"]) else r["cale_pptx"]
        try:
            if sys.platform == "win32":
                os.startfile(deschide)
            else:
                subprocess.Popen(["open", deschide])
        except Exception:
            pass
        messagebox.showinfo("Calcul de probă", "Am deschis PDF-ul de probă.\n\nNu s-a salvat nimic în Drive și "
                            "nimic în BESS_Proposals; fișierele locale se șterg la închiderea programului.")

    def genereaza(self, c):
        """Pasul 2 (după OK): PowerPoint + PDF cu valorile din Excelul salvat, apoi urcare în Drive."""
        intrari, baterie, skid = c["intrari"], c["baterie"], c["skid"]
        r = finalizeaza_oferta(c["calcul"], intrari, baterie, c["nr_baterii"], skid, c["trafo"], status=self.status)
        if c.get("test"):
            self._termina_test(c, r)
            return
        link_drive, eroare_drive = None, None
        if DRIVE is not None:
            self.status("Urc oferta în Drive...")
            try:
                # în modul Microsoft 365 Excelul de calcul e deja în Drive (poate modificat în browser): nu-l suprascriem
                link_drive = DRIVE.urca_oferta(r["folder"], sari={os.path.basename(r["cale_excel"])}
                                               if c["calcul"].get("online") else ())
            except Exception as e:
                eroare_drive = str(e)
        if DRIVE is not None:
            try:
                salveaza_oferte_netrecute(oferte_netrecute() + [self._inregistrare_tabel(intrari, baterie, skid, r)])
            except Exception as e:
                messagebox.showwarning("BESS_Proposals", f"Oferta nu a putut fi pusă în lista de trecut în tabel:\n{e}")
        self.status("Finalizat cu succes!", "green")
        if DRIVE is not None and eroare_drive is None:
            zi = os.path.basename(os.path.dirname(r["folder"]))
            mesaj = f"Oferta a fost urcată în Drive:\n\n📂 Ofertare BESS / {FOLDER_OFERTE} / {zi} / {r['nume']}\n\n"
        else:
            mesaj = f"Fișierele au fost salvate în:\n\n📂 {r['folder']}\n\n"
            if eroare_drive:
                mesaj = f"⚠ Oferta NU a putut fi urcată în Drive:\n{eroare_drive}\n\n" + mesaj
        mesaj += f"{os.path.basename(r['cale_excel'])}\n{r['nume']}.pptx"
        if DRIVE is not None:
            mesaj = ("ℹ Oferta NU e trecută încă în BESS_Proposals. După ce o verifici, folosește butonul "
                     "„Trece o ofertă verificată în BESS_Proposals”.\n\n" + mesaj)
        mesaj += f"\n{r['nume']}.pdf" if r["eroare_pdf"] is None else \
                 f"\n\n⚠ PDF-ul nu s-a putut crea (e nevoie de PowerPoint instalat):\n{r['eroare_pdf']}"
        if r["ramase"]:
            mesaj += "\n\n⚠ Placeholdere rămase neînlocuite în prezentare:\n" + ", ".join(r["ramase"])
        if link_drive:
            if messagebox.askyesno("Gata!", mesaj + "\n\nDeschid folderul ofertei în browser?"):
                webbrowser.open(link_drive)
        elif r["eroare_pdf"] or r["ramase"] or eroare_drive:
            messagebox.showwarning("Gata, dar...", mesaj)
        else:
            messagebox.showinfo("Gata!", mesaj)


if __name__ == "__main__":
    root = tk.Tk()
    if DRIVE_ACTIV and DRIVE_CLIENT_ID and "--motor" not in sys.argv:
        DRIVE = porneste_drive()      # fișierele direct din Drive (copiile locale stau în DRIVE.folder_lucru)
    # PCS, specificații, template-uri și 'Oferte generate': din Drive, altfel din folderul sincronizat 'Ofertare BESS'
    os.chdir(DRIVE.folder_lucru if DRIVE is not None else alege_folder_lucru())
    _nr_setari, _probleme_setari = aplica_setari_generator()
    if _probleme_setari:
        messagebox.showwarning(FISIER_SETARI_GENERATOR, "Probleme în fișierul de setări (rămân valorile din program):"
                               "\n\n- " + "\n- ".join(_probleme_setari))
    # 'Generator Oferte BESS.exe --motor' (sau fișierul .py cu --motor) pornește modul motor pentru Power Apps
    gui = Motor(root) if "--motor" in sys.argv else AplicatieOferte(root)
    root.mainloop()
