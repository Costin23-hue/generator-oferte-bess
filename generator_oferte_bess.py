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
FISIERE_PROGRAM = ("generator_oferte_bess.py",
                   "bess_setari.py",
                   "bess_logica.py",
                   "bess_excel.py",
                   "bess_prezentare.py",
                   "bess_drive.py",
                   "bess_generare.py",
                   "bess_atr.py",
                   "bess_motor.py",
                   "bess_fereastra.py")
# Codul e împărțit pe teme în fișierele de mai sus (bess_*.py); toate se încarcă aici, în același spațiu de nume.
MODULE_PROGRAM = FISIERE_PROGRAM[1:]

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
# 5 MWh (Jinko, REPT) -> ~50 kVA (informație nouă, 11.2026; înainte 55); 6,25 MWh -> 65 kVA.
# Tabelul se scrie și în Excelul de calcul (A24:B25).
KVA_AUX_PER_BATERIE = [(0, 50), (5.6, 65)]
# Consumul auxiliar propriu al unui skid (PCS + MV skid), kVA, după puterea skidului: (până la MW, kVA), crescător.
# Valori acoperitoare (cele mai mari), pentru orice skid. Între trepte: treapta următoare; peste ultima: proporțional.
KVA_CONSUM_SKID = [(2, 30), (5, 40), (7.5, 50), (10, 60)]
# Skidul are trafo propriu, separat, pentru consumul lui (altul decât trafo-ul aux din numele lui, care rămâne pentru
# baterii) -> consumul skidului NU intră în necesarul aux. True = se adună iar (regula veche).
CONSUM_SKID_IN_NECESAR = False
# ATR-uri cu capacitatea doar în Ah: MWh = Ah x tensiunea nominală a bateriei (V) / 1.000.000. Tensiunea depinde de
# baterie: {text din tipul bateriei din ATR: V}, ex. {"LUNA2000": 1331.2}. Se completează în Setari generator.xlsx.
TENSIUNE_BATERIE_V = {}
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
# Commissioning baterie / unitate, după câte containere sunt în proiect: coloanele K..O = 1, 2, 3, 4, 5+ containere
BAT_COL_COM = (10, 11, 12, 13, 14)
SKID_COL_BRAND, SKID_COL_TYPE, SKID_COL_POWER = 2, 3, 4
SKID_COL_PRET, SKID_COL_TRANSPORT, SKID_COL_COM = 5, 6, 7
SKID_COL_GARANTIE = 12                   # coloana M = ani garanție skid / PCS
SKID_COL_DATA = 0                        # coloana A = data cotației skidului
ZILE_COTATIE_VECHE = 60                  # cotațiile mai vechi de atâtea zile sunt semnalate în previzualizare

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
    ("TENSIUNE_BATERIE_V", "ATR: tensiunea nominală a bateriei pentru Ah -> MWh (tip din ATR = V)", "perechi"),
    ("ZILE_COTATIE_VECHE", "Previzualizare: cotații mai vechi de atâtea zile sunt semnalate ⚠", "intreg"),
    ("CONSUM_SKID_IN_NECESAR", "Consumul skidului intră în necesarul aux (NU = skidul are trafo propriu separat)",
     "da_nu"),
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
# ÎNCĂRCAREA PĂRȚILOR PROGRAMULUI (bess_*.py)
# ======================================================================
def _incarca_module():
    """Rulează fișierele bess_*.py în spațiul de nume al acestui fișier (ca și cum ar fi un singur fișier).
    Dacă unul lipsește (ex. instalare veche, cu un singur fișier), îl descarcă o dată de pe GitHub."""
    baza = getattr(sys, "_MEIPASS", FOLDER_PROGRAM)
    for nume in MODULE_PROGRAM:
        cale = os.path.join(baza, nume)
        if not os.path.exists(cale):
            cerere = urllib.request.Request(URL_ACTUALIZARE + nume + f"?t={int(time.time())}",
                                            headers={"Cache-Control": "no-cache"})
            with urllib.request.urlopen(cerere, timeout=30) as r:
                text = r.read().decode("utf-8")
            compile(text, nume, "exec")
            with open(cale, "w", encoding="utf-8", newline="\n") as f:
                f.write(text)
        with open(cale, encoding="utf-8") as f:
            exec(compile(f.read(), cale, "exec"), globals())


_incarca_module()


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
