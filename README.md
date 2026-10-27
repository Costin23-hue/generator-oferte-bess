# generator-oferte-bess

Generator oferte BESS - Taliva. Din `Pricing.xlsx` (OneDrive, folderul Purchasing) și template-urile din
folderul comun `Ofertare BESS`, generează prezentarea PPTX, PDF-ul și Excelul de calcul pentru o ofertă.

## Aplicațiile (Windows .exe și Mac .app)

Se construiesc automat pe GitHub la fiecare modificare a codului:

1. Tab-ul **Actions** → **Construieste aplicatiile** (pornește singur la fiecare modificare; manual: **Run workflow**).
2. Când rularea are bifă verde, deschide-o și descarcă din **Artifacts**:
   - `Generator Oferte BESS - Windows` → `Generator Oferte BESS.exe`
   - `Generator Oferte BESS - macOS` → `Generator Oferte BESS - Mac.zip`
3. Pune-le în folderul comun `Ofertare BESS` din OneDrive.

**Windows:** `.exe`-ul se pornește direct din `Ofertare BESS` (scurtătură pe Desktop).

**Mac:** dezarhivează zip-ul, trage aplicația în **Applications**; prima dată **click dreapta → Open → Open**.
La prima pornire alege folderul `Ofertare BESS` din OneDrive; macOS va cere permisiunea de a controla Excel și PowerPoint.

## Ce trebuie să fie în `Ofertare BESS`

- `PCS - MV skid.xlsx`
- `Specificatii_baterii_BESS.xlsx`
- `Project Proposal TAMPLATE.xlsx`
- `TEMPLATE BESS.pptx`

Ofertele se salvează în `Ofertare BESS/Oferte generate/`. `Pricing.xlsx` se găsește automat în folderul
sincronizat care se termină cu `Purchasing Prices Update_BESS`.

## Fișiere

| Fișier | Rol |
|---|---|
| `generator_oferte_bess.py` | programul (Windows și Mac) |
| `construieste_exe.py` | construiește `.exe` (Windows) / `.app` (Mac) cu PyInstaller |
| `construieste_exe.bat` | construire locală pe Windows (Python 3.14) |
| `construieste_app.command` | construire locală pe Mac |
| `.github/workflows/construieste.yml` | construirea automată pe GitHub |

Fișierele cu prețuri (`.xlsx`, `.pptx`, `.pdf`) nu se încarcă aici (vezi `.gitignore`).
