# -*- coding: utf-8 -*-
"""Drive (Microsoft Graph) și BESS_Proposals.
Parte din generator_oferte_bess.py: se încarcă în același spațiu de nume (toate funcțiile
și setările sunt comune), deci aici NU se fac importuri proprii și nu se rulează direct."""
# pylint: disable=undefined-variable
# flake8: noqa

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
        self._oferte_zi = {}                                            # nume folder ofertă -> zi
        try:
            nume = []
            for c in self._copii(self.ofertare, FOLDER_OFERTE):
                nume.append(c["name"])
                if "folder" in c and MODEL_FOLDER_ZI.match(c["name"]):       # ofertele din folderele pe zile
                    for x in self._copii(self.ofertare, f"{FOLDER_OFERTE}/{c['name']}"):
                        nume.append(x["name"])
                        self._oferte_zi.setdefault(x["name"], c["name"])
            return nume
        except NegasitDrive:
            return []

    def lista_oferte(self):
        """Toate ofertele din 'Oferte generate' (pe zile și vechi): [{'zi', 'nume', 'url', 'id'}], cele noi primele."""
        rezultat = []
        for c in self._copii(self.ofertare, FOLDER_OFERTE):
            if "folder" not in c or c["name"].startswith("_"):
                continue
            if MODEL_FOLDER_ZI.match(c["name"]):
                for x in self._copii(self.ofertare, f"{FOLDER_OFERTE}/{c['name']}"):
                    if "folder" in x:
                        rezultat.append({"zi": c["name"], "nume": x["name"], "url": x.get("webUrl"), "id": x["id"]})
            else:
                rezultat.append({"zi": (c.get("createdDateTime") or "")[:10], "nume": c["name"],
                                 "url": c.get("webUrl"), "id": c["id"]})
        return sorted(rezultat, key=lambda o: (o["zi"], o["nume"]), reverse=True)

    def date_oferta(self, id_folder):
        """Conținutul date_oferta.json din folderul unei oferte (None dacă oferta e veche și nu-l are)."""
        drive_id = self.ofertare[0]
        for f in self._copii((drive_id, id_folder)):
            if f["name"] == FISIER_DATE_OFERTA:
                temp = os.path.join(self.folder_lucru, FISIER_DATE_OFERTA)
                self._descarca(drive_id, f["id"], temp)
                with open(temp, encoding="utf-8") as fh:
                    return json.load(fh)
        return None

    def info_oferta(self, nume):
        """(zi, are_prezentare) pentru folderul unei oferte din 'Oferte generate/<zi>/'; None dacă nu-l găsim."""
        zi = getattr(self, "_oferte_zi", {}).get(nume)
        if zi is None:
            return None
        try:
            fisiere = self._copii(self.ofertare, f"{FOLDER_OFERTE}/{zi}/{nume}")
        except NegasitDrive:
            return None
        return zi, any(f["name"].lower().endswith(".pptx") for f in fisiere)

    def redenumeste_oferta(self, zi, vechi, nou):
        """Redenumește folderul unei oferte din Drive (dacă există)."""
        drive_id, item_id = self.ofertare
        cale = f"{FOLDER_OFERTE}/{zi}/{vechi}" if MODEL_FOLDER_ZI.match(zi) else f"{FOLDER_OFERTE}/{vechi}"
        try:
            el = self._cerere("GET", f"/drives/{drive_id}/items/{item_id}:/{_cale_url(cale)}")
        except NegasitDrive:
            return
        self._cerere("PATCH", f"/drives/{drive_id}/items/{el['id']}", tip="application/json",
                     date=json.dumps({"name": nou}).encode())

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
