# -*- coding: utf-8 -*-
"""Fereastra programului (Tkinter).
Parte din generator_oferte_bess.py: se încarcă în același spațiu de nume (toate funcțiile
și setările sunt comune), deci aici NU se fac importuri proprii și nu se rulează direct."""
# pylint: disable=undefined-variable
# flake8: noqa

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
            ("Client:", ""), ("Nume Proiect:", ""),          # goale la fiecare ofertă nouă
            # cu Drive-ul conectat rămâne gol: nr. de contract se ia automat din BESS_Proposals
            ("Nr. proiect:", "" if DRIVE is not None else NR_PROIECT_IMPLICIT.format(an_luna=datetime.now().strftime("%Y%m"))),
            ("Locație:", "Romania"),                        # țara se schimbă din „Țară (foaia din tabel)”
            ("Capacitate (MWh):", "20"), ("Putere (MW):", "10"), ("Putere max. admisă (MW, opțional):", ""),
            ("Valoare EMS (USD):", f"{EMS_AUTO:g}"),
            ("Valoare ENG (USD):", "0"), ("Adaos Comercial (Margin):", f"{ADAOS_IMPLICIT:.2f}"),
            ("Curs USD/EUR:", f"{curs_salvat()[0]:g}"),     # ultimul curs BCE; se actualizează la pornire
        ]
        self.etichete = {}
        for i, (eticheta, implicit) in enumerate(campuri):
            self.etichete[eticheta] = ttk.Label(frame, text=eticheta)
            self.etichete[eticheta].grid(row=i, column=0, sticky="w", pady=5)
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
        self._curs = None
        self._arata_data_curs(curs_salvat()[1], salvat=True)
        threading.Thread(target=self._descarca_curs, daemon=True).start()
        self.root.after(1000, self._arata_curs)
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
                "BESS_Proposals. Se șterge la „Renunță”, la următorul calcul sau la închiderea programului. "
                "La OK te întreb dacă îl salvezi ca ofertă (client, proiect...) sau faci doar un PDF de probă."))
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
        if not c or not messagebox.askyesno("Renunț", "Renunți la acest calcul? Se șterge Excelul lui (nu se face "
                                                       "PowerPoint / PDF); calculele anterioare ale versiunii rămân."):
            return
        self._sterge_calcul(c)
        self._calcul_in_asteptare = None
        self._stare_verificare(False)
        self.tabel_verificare.delete(*self.tabel_verificare.get_children())
        self.lbl_rezumat.config(text="")
        self.status("Calcul anulat.", "gray")
        self.taburi.select(0)

    def _sterge_calcul(self, c, arata_erori=True):
        """Șterge un calcul. Proba: tot folderul (local și din Drive). Ofertă: doar Excelul acestui calcul; folderul
        versiunii se șterge doar dacă a rămas gol (celelalte calcule ale versiunii rămân)."""
        calcul = c["calcul"]
        try:
            if c.get("test"):
                shutil.rmtree(calcul["folder"], ignore_errors=True)   # dacă Excelul e deschis, rămâne local
                if calcul["rez"].get("folder_drive"):                # mod Microsoft 365: calculul e și în Drive
                    DRIVE.sterge(calcul["rez"]["folder_drive"])
                return
            try:
                os.remove(calcul["cale_excel"])
                if os.path.isdir(calcul["folder"]) and not os.listdir(calcul["folder"]):
                    os.rmdir(calcul["folder"])
            except OSError:
                pass
            if calcul["rez"].get("excel_drive"):
                DRIVE.sterge(calcul["rez"]["excel_drive"])
                if not DRIVE._copii(calcul["rez"]["folder_drive"]):
                    DRIVE.sterge(calcul["rez"]["folder_drive"])
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
        if c.get("test"):
            date_oferta = self._dialog_salvare_test(c)
            if date_oferta is None:
                return
        elif not messagebox.askyesno("OK", "Ai verificat calculul? Fac PowerPoint-ul și PDF-ul cu valorile din "
                                           "Excelul salvat?"):
            return
        self._stare_verificare(False)
        try:
            if c.get("test") and date_oferta:               # proba devine ofertă reală, salvată în Drive
                c = self._salveaza_test_ca_oferta(c, date_oferta)
            self.genereaza(c)
            self._calcul_in_asteptare = None
            self.taburi.select(0)
        except Exception as e:
            self._stare_verificare(True)
            self.status("Eroare la generare.", "red")
            messagebox.showerror("Eroare", str(e))

    def _dialog_salvare_test(self, c):
        """La OK pe un calcul de probă: datele ofertei (client, proiect...) ca să fie salvată în Drive, sau doar
        PDF-ul de probă. Returnează dict (salvează), {} (doar probă) sau None (anulat)."""
        win = tk.Toplevel(self.root)
        win.title("Calcul de probă – salvez oferta?")
        win.transient(self.root)
        win.grab_set()
        ttk.Label(win, justify="left", wraplength=460, text=(
            "Calculul de probă e gata. Dacă vrei să-l păstrezi ca ofertă (prezentare, PDF și Excel în Drive), "
            "completează datele de mai jos. Excelul de probă (cu ce ai modificat în el) se copiază în oferta nouă.")
                  ).grid(row=0, column=0, columnspan=2, padx=12, pady=(12, 8), sticky="w")
        campuri = (("client", "Client:"), ("proiect", "Nume Proiect:"), ("locatie", "Locație:"),
                   ("nr_proiect", "Nr. proiect:"))
        intrari, valori = c["intrari"], {}
        for i, (cheie, eticheta) in enumerate(campuri, start=1):
            ttk.Label(win, text=eticheta).grid(row=i, column=0, padx=12, pady=4, sticky="w")
            e = ttk.Entry(win, width=40)
            v = str(intrari.get(cheie, "")).strip()
            if v.lower() not in (TEXT_TEST, NR_PROIECT_TEST.lower(), "-"):
                e.insert(0, v)
            e.grid(row=i, column=1, padx=12, pady=4, sticky="w")
            valori[cheie] = e
        rezultat = {"v": None}

        def salveaza():
            date = {k: e.get().strip() for k, e in valori.items()}
            if not date["client"] or not date["proiect"] or TEXT_TEST in (date["client"].lower(),
                                                                          date["proiect"].lower()):
                messagebox.showwarning("Date lipsă", "Completează Client și Nume Proiect (altele decât „test”).",
                                       parent=win)
                return
            date["locatie"] = date["locatie"] or "-"
            date["nr_proiect"] = date["nr_proiect"] or NR_PROIECT_IMPLICIT.format(
                an_luna=datetime.now().strftime("%Y%m"))
            rezultat["v"] = date
            win.destroy()

        def doar_proba():
            rezultat["v"] = {}
            win.destroy()

        bf = ttk.Frame(win)
        bf.grid(row=len(campuri) + 1, column=0, columnspan=2, pady=12)
        tk.Button(bf, text="💾 Salvează ca ofertă (Drive)", command=salveaza, **STIL_BUTON).pack(side="left", padx=6)
        tk.Button(bf, text="Doar PDF de probă", command=doar_proba).pack(side="left", padx=6)
        tk.Button(bf, text="Anulează", command=win.destroy).pack(side="left", padx=6)
        valori["client"].focus_set()
        self.root.wait_window(win)
        return rezultat["v"]

    def _salveaza_test_ca_oferta(self, c, date_oferta):
        """Copiază calculul de probă într-o ofertă reală; formularul primește datele noi (pentru calculele următoare)."""
        self.status("Salvez calculul de probă ca ofertă...")
        intrari = dict(c["intrari"], **date_oferta)
        calcul = transforma_calcul_test(c["calcul"], intrari,
                                        nume_existente=DRIVE.nume_oferte() if DRIVE is not None else (),
                                        info_drive=DRIVE.info_oferta if DRIVE is not None else None)
        self._sterge_calcul(c, arata_erori=False)
        for cheie, eticheta in (("client", "Client:"), ("proiect", "Nume Proiect:"), ("locatie", "Locație:"),
                                ("nr_proiect", "Nr. proiect:")):
            self.intrari[eticheta].delete(0, "end")
            self.intrari[eticheta].insert(0, intrari[cheie])
        c = dict(c, intrari=intrari, calcul=calcul, test=False)
        self._calcul_in_asteptare = c
        return c

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
    def _descarca_curs(self):
        try:
            self._curs = actualizeaza_curs()
        except Exception as e:
            self._curs = e

    def _arata_data_curs(self, data, salvat=False, eroare=False):
        text = "Curs USD/EUR:"
        if data:
            text = f"Curs USD/EUR (BCE {data}" + (", ultimul salvat" if salvat and eroare else "") + "):"
        elif eroare:
            text = "Curs USD/EUR (nu s-a putut actualiza):"
        self.etichete["Curs USD/EUR:"].config(text=text)

    def _arata_curs(self, incercari=30):
        """Pune cursul BCE în fereastră (doar dacă nu l-ai modificat tu între timp)."""
        if self._curs is None:
            if incercari > 0:
                self.root.after(1000, lambda: self._arata_curs(incercari - 1))
            return
        if isinstance(self._curs, Exception):
            self._arata_data_curs(curs_salvat()[1], salvat=True, eroare=True)
            return
        curs, data = self._curs
        camp = self.intrari["Curs USD/EUR:"]
        if camp.get().strip().replace(",", ".") in (f"{curs_salvat()[0]:g}", f"{CURS_IMPLICIT:g}", ""):
            camp.delete(0, "end")
            camp.insert(0, f"{curs:g}")
        self._arata_data_curs(data)

    def _arata_actualizare(self, incercari=60):
        if self._actualizare is None:                       # verificarea încă rulează (max. ~1 minut)
            if incercari:
                self.root.after(1000, lambda: self._arata_actualizare(incercari - 1))
            else:
                self._programeaza_verificare_actualizare()
            return
        if self._actualizare and getattr(self, "btn_actualizare", None) is None:
            self.btn_actualizare = tk.Button(
                self.continut, text="⬆ Versiune nouă disponibilă – Actualizează", bg="#1e7e34", fg="white",
                font=("Helvetica", 10, "bold"), command=self.actualizeaza_programul)
            self.btn_actualizare.pack(pady=(0, 8), fill="x", padx=20)
            self.status("A apărut o versiune nouă a programului: butonul verde „Actualizează”.", "green")
        self._programeaza_verificare_actualizare()

    def _programeaza_verificare_actualizare(self):
        """Cu programul deschis, verifică din nou pe GitHub la fiecare MINUTE_VERIFICARE_ACTUALIZARE minute."""
        if getattr(self, "_verificare_programata", False):
            return
        self._verificare_programata = True

        def verifica():
            self._verificare_programata = False
            self._actualizare = None
            threading.Thread(target=lambda: setattr(self, "_actualizare", cauta_actualizare()), daemon=True).start()
            self.root.after(1000, self._arata_actualizare)
        self.root.after(MINUTE_VERIFICARE_ACTUALIZARE * 60 * 1000, verifica)

    def actualizeaza_programul(self):
        in_lucru = self._calcul_in_asteptare and not self._calcul_in_asteptare.get("test")
        if not messagebox.askyesno("Actualizare", "Instalez versiunea nouă și repornesc programul?\n\n"
                                   "(Ce ai completat în fereastră se pierde."
                                   + (" Calculul care așteaptă OK rămâne în folderul ofertei, dar va trebui "
                                      "refăcut pentru prezentare." if in_lucru else "") + ")"):
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
        ttk.Label(butoane, text="   Tensiune baterie (V):").pack(side="left")
        self.txt_tensiune_atr = ttk.Entry(butoane, width=8)
        self.txt_tensiune_atr.pack(side="left", padx=4)
        ttk.Button(butoane, text="Ah → MWh", command=self._mwh_din_ah).pack(side="left")
        cadru = ttk.Frame(tab)
        cadru.pack(fill="both", expand=True, padx=10, pady=8)
        self.txt_atr = tk.Text(cadru, height=24, wrap="word", font=("Consolas", 10))
        bara = ttk.Scrollbar(cadru, command=self.txt_atr.yview)
        self.txt_atr.configure(yscrollcommand=bara.set, state="disabled")
        bara.pack(side="right", fill="y")
        self.txt_atr.pack(side="left", fill="both", expand=True)
        self._atr = None

    def _mwh_din_ah(self):
        """MWh = Ah (din ATR) x tensiunea nominală a bateriei / 1.000.000."""
        r = self._atr
        if not r or not r.get("capacitate_ah"):
            messagebox.showinfo("Ah → MWh", "Citește întâi un ATR care are capacitatea în Ah.")
            return
        try:
            v = float(self.txt_tensiune_atr.get().replace(",", "."))
        except ValueError:
            messagebox.showwarning("Ah → MWh", "Scrie tensiunea nominală a bateriei în V (ex. 1331,2).")
            return
        r["mwh"] = round(r["capacitate_ah"] * v / 1e6, 3)
        r["avertismente"] = [a for a in r["avertismente"] if not a.startswith(("Capacitatea", "Capacitate din Ah"))]
        r["avertismente"].append(f"Capacitate din Ah: {nr_eu(r['capacitate_ah'])} Ah x {nr_eu(v)} V = "
                                 f"{nr_eu(r['mwh'])} MWh.")
        self._arata_atr(r)

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
                + (f"   ({nr_eu(r['capacitate_ah'])} Ah)" if r.get("capacitate_ah") else "")
                + (f"   (bateriile din tabel: {nr_eu(r['mwh_baterii'])} MWh)" if r.get("mwh_baterii") else ""),
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
                *([f"  Putere instalată minimă cerută (MW):     {nr_eu(r['mw_min'])}"] if r.get("mw_min") else []),
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
                 + ", " + self.combo_tara.get())
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
        self.combo_tara.bind("<<ComboboxSelected>>", lambda e: (self._schimba_tara_locatie(), self.ia_nr_contract()))
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
                  **STIL_BUTON).grid(row=7, column=0, columnspan=2, sticky="we", padx=5, pady=(4, 4))
        tk.Button(grup, text="📂 Istoric oferte / preia o ofertă veche…", command=self.dialog_istoric)\
            .grid(row=8, column=0, columnspan=2, sticky="we", padx=5, pady=(4, 8))

    def _schimba_tara_locatie(self):
        """Țara aleasă pentru BESS_Proposals trece și în locație (doar partea de țară)."""
        camp = self.intrari["Locație:"]
        nou = locatie_cu_tara(camp.get(), self.combo_tara.get(), tuple(TABELE_OFERTE))
        camp.delete(0, "end")
        camp.insert(0, nou)

    def _stare_formular(self):
        """Tot ce e completat în tab-ul Ofertă (texte + meniuri), ca să se poată relua într-o versiune nouă."""
        d = {k: e.get() for k, e in self.intrari.items()}
        for cheie, w in (("_baterie", self.combo_baterie), ("_brand", self.combo_brand),
                         ("_nr_skid", self.combo_nr_skid), ("_incoterm", self.combo_incoterm),
                         ("_poc", self.combo_poc), ("_tara", self.combo_tara), ("_tip", self.combo_tip)):
            d[cheie] = w.get()
        d["_rezerva"] = {k: e.get() for k, e in getattr(self, "rezerva", {}).items()}
        d["_trafo"] = self._mod_trafo()
        return d

    def _aplica_formular(self, d):
        for k, e in self.intrari.items():
            if k in d and k != "Curs USD/EUR:":           # cursul rămâne cel de azi
                e.delete(0, "end")
                e.insert(0, d[k])
        for cheie, w in (("_baterie", self.combo_baterie), ("_brand", self.combo_brand),
                         ("_nr_skid", self.combo_nr_skid), ("_incoterm", self.combo_incoterm),
                         ("_poc", self.combo_poc), ("_tara", self.combo_tara), ("_tip", self.combo_tip)):
            if d.get(cheie) and (d[cheie] in w["values"] or cheie == "_poc"):
                w.set(d[cheie])
        for k, v in (d.get("_rezerva") or {}).items():
            if k in getattr(self, "rezerva", {}):
                self.rezerva[k].delete(0, "end")
                self.rezerva[k].insert(0, v)
        if "_trafo" in d:
            self.fara_aux.set(d["_trafo"] == "fara")
            self.foloseste_aux.set(d["_trafo"] is True)

    def dialog_istoric(self):
        """Ofertele din Drive: căutare (nr., client, proiect), deschidere în browser, preluare în formular."""
        if DRIVE is None:
            messagebox.showwarning("Istoric", "Istoricul merge doar cu Drive-ul conectat.")
            return
        try:
            self.status("Citesc ofertele din Drive...")
            oferte = DRIVE.lista_oferte()
        except Exception as e:
            messagebox.showerror("Istoric", f"Nu pot citi ofertele din Drive:\n{e}")
            return
        finally:
            self.status("Aștept date...", "gray")
        win = tk.Toplevel(self.root)
        win.title("Istoric oferte")
        win.transient(self.root)
        sus = ttk.Frame(win)
        sus.pack(fill="x", padx=10, pady=8)
        ttk.Label(sus, text="Caută (nr. contract, proiect, MWh...):").pack(side="left")
        cautare = ttk.Entry(sus, width=40)
        cautare.pack(side="left", padx=6)
        cautare.insert(0, self.intrari["Nr. proiect:"].get().strip())
        lista = tk.Listbox(win, width=130, height=20, font=("Consolas", 9))
        lista.pack(padx=10, fill="both", expand=True)
        vizibile = []

        def filtreaza(*_):
            text = cautare.get().strip().lower()
            vizibile[:] = [o for o in oferte if not text or text in (o["zi"] + " " + o["nume"]).lower()]
            lista.delete(0, "end")
            for o in vizibile:
                lista.insert("end", f"{o['zi']}   {o['nume']}")

        def aleasa():
            if not lista.curselection():
                messagebox.showinfo("Istoric", "Alege o ofertă din listă.", parent=win)
                return None
            return vizibile[lista.curselection()[0]]

        def deschide(*_):
            o = aleasa()
            if o and o.get("url"):
                webbrowser.open(o["url"])

        def preia():
            o = aleasa()
            if not o:
                return
            try:
                d = DRIVE.date_oferta(o["id"])
            except Exception:
                d = None
            nota = ""
            if d is None:                                # ofertă veche: din nume + clientul din BESS_Proposals
                d = date_din_nume_oferta(o["nume"]) or {}
                nota = "\n\nOferta e veche (fără datele salvate): am luat doar ce se vede în nume"
                try:
                    client = next((c for _, nr, c in contracte_din_tabel(self.combo_tara.get())
                                   if nr == d.get("Nr. proiect:")), None)
                    if client:
                        d["Client:"] = client
                        nota += " și clientul din BESS_Proposals"
                except Exception:
                    pass
                nota += ". Restul verifică-l."
            self._aplica_formular(d)
            win.destroy()
            self.taburi.select(self.tab_oferta)
            messagebox.showinfo("Istoric", f"Am preluat oferta în formular (v{d.get('_versiune', '?')}). La calcul "
                                           "se face automat versiunea următoare." + nota)

        cautare.bind("<KeyRelease>", filtreaza)
        lista.bind("<Double-Button-1>", deschide)
        bf = ttk.Frame(win)
        bf.pack(pady=10)
        tk.Button(bf, text="↺ Preia în formular (versiune nouă)", command=preia, **STIL_BUTON).pack(side="left", padx=6)
        tk.Button(bf, text="🌐 Deschide în browser", command=deschide).pack(side="left", padx=6)
        tk.Button(bf, text="Închide", command=win.destroy).pack(side="left", padx=6)
        filtreaza()

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

    def dialog_tabel(self, selecteaza_ultima=False):
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
        if selecteaza_ultima and lista.size():
            lista.selection_set("end")
            lista.see("end")

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
                  f"rând Excel {baterie['rand_excel']}) -> baterii: {cost_bat:,.0f} $"
                  + avertisment_cotatie(baterie.get("data_cotatie"), "bateriei")))
        com_bat = commissioning_baterie(baterie, nr_baterii)
        if com_bat is not None:
            txt += f"\nCommissioning baterii: {nr_baterii} x {bani_eu(com_bat, 0)} $ = {bani_eu(nr_baterii * com_bat, 0)} $"
        kva_bat = baterie.get("kva_aux") or kva_aux_baterie(baterie["cap"])
        txt += (f"\nConsum auxiliar: {nr_baterii} baterii x {kva_bat:g} kVA"
                + (" + consum skid după putere (vezi la fiecare skid); " if CONSUM_SKID_IN_NECESAR else
                   " (skidul are trafo propriu pentru consumul lui); ")
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
        minim = min((cost_bat + o["cost_total"] for o in optiuni), default=0)
        for i, o in enumerate(optiuni):
            text = descriere_skid(o, intrari["mw"])
            prima, rest = (text.split("\n", 1) + [""])[:2]
            prima += (f"   [total achiziție {bani_eu(cost_bat + o['cost_total'], 0)} $, "
                      f"{procent_fata_de(cost_bat + o['cost_total'], minim)}]"
                      + avertisment_cotatie(o.get("data_cotatie"), "skidului"))
            ttk.Radiobutton(win, variable=var, value=i, text=prima + ("\n" + rest if rest else ""))\
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
        if vechi and vechi.get("test"):
            self._sterge_calcul(vechi)                 # calculul de probă anterior nu se păstrează
        # calculul anterior fără OK rămâne în folderul versiunii (istoric); versiunea crește doar după prezentare
        calcul = calculeaza_oferta(self.date, intrari, baterie, nr_baterii, skid, trafo, status=self.status,
                                   online=self.mod_online.get(),
                                   nume_existente=DRIVE.nume_oferte() if DRIVE is not None else (),
                                   info_drive=DRIVE.info_oferta if DRIVE is not None else None)
        self._calcul_in_asteptare = {"calcul": calcul, "intrari": intrari, "baterie": baterie,
                                     "formular": self._stare_formular(),
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
        try:                                           # datele formularului, pentru „Preia în formular” (v2, v3...)
            with open(os.path.join(r["folder"], FISIER_DATE_OFERTA), "w", encoding="utf-8") as f:
                json.dump(dict(c.get("formular") or {}, _versiune=intrari.get("versiune")), f, ensure_ascii=False,
                          indent=1)
        except OSError:
            pass
        if DRIVE is not None:
            self.status("Urc oferta în Drive...")
            try:
                # în modul Microsoft 365 Excelul de calcul e deja în Drive (poate modificat în browser): nu-l suprascriem
                link_drive = DRIVE.urca_oferta(r["folder"], sari={
                    f for f in os.listdir(r["folder"]) if f.startswith("Calcul_") and f.endswith(".xlsx")}
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
        if DRIVE is not None and messagebox.askyesno(
                "BESS_Proposals", "Treci acum oferta în BESS_Proposals?\n\n(Se deschide lista cu oferta selectată; "
                                  "vezi datele înainte de confirmare.)"):
            self.dialog_tabel(selecteaza_ultima=True)
