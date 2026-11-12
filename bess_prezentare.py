# -*- coding: utf-8 -*-
"""Prezentarea PowerPoint și PDF-ul.
Parte din generator_oferte_bess.py: se încarcă în același spațiu de nume (toate funcțiile
și setările sunt comune), deci aici NU se fac importuri proprii și nu se rulează direct."""
# pylint: disable=undefined-variable
# flake8: noqa

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
    model = _model_oferta(nr_proiect, proiect)
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


def _model_oferta(nr_proiect, proiect):
    return re.compile(re.escape(f"{nume_fisier_valid(nr_proiect)}_BESS_Indicative_Proposal_Taliva_"
                                f"{nume_fisier_valid(proiect)}_") + r".*_v(\d+)(\.pptx)?$", re.I)


def versiune_de_lucru(nr_proiect, proiect, nume_existente=(), info_drive=None):
    """Versiunea pentru un calcul nou și folderul în care intră. Versiunea crește doar după ce s-a făcut
    prezentarea: cât timp ultima versiune nu are PPTX (doar calcule), calculele noi intră în același folder.
    info_drive(nume_folder) -> (zi, are_pptx) sau None (pentru folderele din Drive).
    Returnează (versiune, (zi, nume_folder) sau None = folder nou în ziua de azi)."""
    ultima = urmatoarea_versiune(nr_proiect, proiect, nume_existente) - 1
    if ultima < 1:
        return 1, None
    model = _model_oferta(nr_proiect, proiect)
    potrivite = lambda nume: (m := model.match(nume)) and not m.group(2) and int(m.group(1)) == ultima
    gasit = None
    for zi in ([""] + sorted(os.listdir(FOLDER_OFERTE))) if os.path.isdir(FOLDER_OFERTE) else []:
        baza = os.path.join(FOLDER_OFERTE, zi)
        if zi and not (os.path.isdir(baza) and (MODEL_FOLDER_ZI.match(zi))):
            continue
        for nume in os.listdir(baza):
            cale = os.path.join(baza, nume)
            if os.path.isdir(cale) and potrivite(nume):
                if any(f.lower().endswith(".pptx") for f in os.listdir(cale)):
                    return ultima + 1, None
                gasit = gasit or (zi, nume)
    for nume in nume_existente:
        if model.match(nume) and model.match(nume).group(2) and int(model.match(nume).group(1)) == ultima:
            return ultima + 1, None                       # prezentare veche, direct în 'Oferte generate'
        if potrivite(nume) and info_drive is not None:
            info = info_drive(nume)
            if info is None:
                continue
            if info[1]:
                return ultima + 1, None
            gasit = gasit or (info[0], nume)
    if gasit is None:                                     # versiunea există doar ca nume (fără folder de verificat)
        return ultima + 1, None
    return ultima, gasit


def nume_oferta(intrari):
    """Numele folderului / prezentării / PDF-ului: ST..._BESS_Indicative_Proposal_Taliva_<proiect>_20MWh_10MW_v2."""
    return NUME_PREZENTARE.format(nr_proiect=nume_fisier_valid(intrari["nr_proiect"]),
                                  proiect=nume_fisier_valid(intrari["proiect"]),
                                  mwh=nr_eu(intrari["mwh"]), mw=nr_eu(intrari["mw"])) + f"_v{intrari['versiune']}"


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
