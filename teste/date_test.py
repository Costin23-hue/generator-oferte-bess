"""Date INVENTATE pentru teste (nu sunt prețurile reale din Pricing.xlsx).
Au aceeași structură de coloane ca foile din Pricing și 'PCS - MV skid.xlsx'."""
import pandas as pd

COLOANE_BATERIE = (["Date", "Supplier", "Brand", "Type", "Capacity MWh", "FOB", "Tax",
                    "Price / MWh with taxes (DAP)", "Price / MWh FOB with taxes and DDP from DAP Tax",
                    "Transport"] + [f"Col{i}" for i in range(10, 20)] + ["Warranty"])
COLOANE_SKID = (["Date", "Supplier", "Brand", "Type", "Power", "Price", "Transport", "Commissioning"]
                + [f"Col{i}" for i in range(8, 12)] + ["Warranty"])


def _baterie(data, brand, tip, cap, dap, ddp, transport, garantie=15):
    return [data, "-", brand, tip, cap, 0, 0, dap, ddp, transport] + [None] * 10 + [garantie]


def _skid(brand, tip, putere, pret, transport, com, garantie=5):
    return ["2026-09-01", "-", brand, tip, putere, pret, transport, com] + [None] * 4 + [garantie]


def date_test():
    bat = pd.DataFrame([
        _baterie("2026-01-01", "Jinko", "JKE-5015K-2H-LAA", 5.015, 90000, 95000, 5000),   # cotație veche
        _baterie("2026-09-01", "Jinko", "JKE-5015K-2H-LAA", 5.015, 80000, 85000, 4000),
        _baterie("2026-09-01", "Jinko", "JKE-6250K-2H-LAA", 6.25, 78000, 83000, 4500),
        _baterie("2026-09-01", "Jinko", "JKE-5015K-4H-LAA", 5.015, 70000, 75000, 4000),
        _baterie("2026-09-01", "Jinko", "JKE-4597K-2H-LAA", 4.597, 84000, 89000, 4000),   # 11 rack-uri
        _baterie("2026-09-01", "REPT", "REPT 5,015", 5.015, 60000, 65000, 4000, 10),
    ], columns=COLOANE_BATERIE)
    skid = pd.DataFrame([
        _skid("Kehua", "BCS5000K-B-HUD MV", 5, 200000, 10000, 8000),
        _skid("Kehua", "BCS7500K-B-HUD MV", 7.5, 280000, 12000, 9000),
        _skid("Sinexcel", "PWS2-5000 MV", 5, 210000, 10000, 8000),
        _skid("TBEA", "TBEA 10MW + AUX TRAFO 200kVA", 10, 380000, 15000, 10000),
    ], columns=COLOANE_SKID)
    skid_poc = pd.DataFrame([
        _skid("Kehua", "BCS5000K-B-HUD 33kV", 5, 230000, 10000, 8000),
        _skid("TBEA", "TBEA 10MW 33kV", 10, 400000, 15000, 10000),
    ], columns=COLOANE_SKID)
    aux = pd.DataFrame({"Detail": ["450kVA 690V/400V Aux TRAFO", "1250 KVA Aux transformer"],
                        "Price USD DAP": [12000, 25000], "Price USD EXW Ankara": [10000, 21000]})
    pcs = pd.DataFrame({
        "SKID": ["BCS5000K-B-HUD", "BCS7500K-B-HUD", "PWS2-5000", "TBEA 10MW + AUX TRAFO 200kVA"],
        "PCS power": [1.25, 1.25, 1.25, 0.215], "no / frame": [4, 6, 4, 48],
        "Skid power": [5, 7.5, 5, 10], "Min no of imputs/ skid": [2, 3, 2, 12],
        "Max no of imputs/ skid": [4, 6, 4, 24], "Number of DC branch": [4, 6, 4, 48],
        "Aux trafo kVA": [None, None, 400, 200],          # Sinexcel: trafo aux pentru tot sistemul
    })
    return {"bat": bat, "skid": skid, "aux": aux, "pcs": pcs, "specs": None, "skid_poc": skid_poc,
            "foaie_poc": "MV SKID for different POC", "sursa": "test", "mtime": 0, "mtime_specs": None}


def intrari_test(mwh=20, mw=10, poc_kv=20):
    return {"client": "Test", "proiect": "P", "locatie": "-", "nr_proiect": "ST202611001", "mwh": mwh, "mw": mw,
            "ems": 3500, "eng": 0, "adaos": 1.1, "curs": 0.89, "poc_kv": poc_kv}
