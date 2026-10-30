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

## Lucrul direct cu Drive-ul (fără foldere sincronizate)

La pornire, programul se loghează cu contul Microsoft de firmă (browser, doar prima dată; la
„Permissions requested” → Accept) și lucrează direct cu Drive-ul, prin aplicația „Generator Oferte BESS”
înregistrată în Entra ID (permisiuni delegate: vede doar ce vede utilizatorul logat):

- descarcă `Pricing.xlsx` din folderul Purchasing și restul fișierelor din `Ofertare BESS`
  (doar dacă s-au schimbat) într-un folder local de lucru (`Drive cache`);
- generează oferta local (Excel + PowerPoint) și o urcă în `Ofertare BESS/Oferte generate/<nume>`.

Setările sunt în `CONFIGURARE` (`DRIVE_CLIENT_ID`, `DRIVE_TENANT_ID`, `LINK_FOLDER_PRICING`, `LINK_FOLDER_OFERTARE`).
Dacă Drive-ul nu e disponibil, programul poate continua cu folderele sincronizate (ca înainte).

## Componente manuale

Tab-ul „Componente manuale” permite o baterie și/sau un skid care nu sunt în Pricing (cotații primite de la
furnizori). În tab-ul „Ofertă” se aleg cu „✎ Baterie manuală” / „✎ Skid manual” și se pot combina cu
componentele din baza de date. Bateria: preț DDP/DAP în $/MWh, transport pe container (sau inclus în preț),
garanție, extra warranty opțional, specificații pentru prezentare. Skidul: putere, preț, transport (sau inclus),
commissioning, trafo aux inclus; compatibilitatea nu e verificată (programul cere confirmare înainte de generare).
Fără extra warranty completat, rândul componentei manuale dispare din tabelul „Warranty Extension”.
În câmpurile de bani, punctul e separator de mii („75.000” = 75000).

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

## Modul motor (Power Apps / Power Automate)

`Generator Oferte BESS.exe --motor` (sau `Porneste motor.bat`) pornește o fereastră mică ce verifică la 10 secunde
folderul `Ofertare BESS/Cereri`. Fiecare cerere `.json` (scrisă de Power Automate, vezi `exemple/cerere_exemplu.json`)
e generată ca o ofertă normală în `Oferte generate`; rezultatul se scrie în `Cereri/Rezultate/<id>.json`
(cu căile fișierelor în OneDrive), iar cererea e mutată în `Cereri/Procesate`.
O cerere cu `"tip": "previzualizare"` nu generează fișiere: răspunsul conține bateria, trafo-ul și skidurile
compatibile (`optiuni`), din care colegul alege unul în Power Apps; cererea de generare trimite apoi `"skid"`.
`Cereri/Liste.xlsx` conține tabelele `TabelBaterii`, `TabelBranduri`, `TabelPOC` pentru meniurile din Power Apps.
Jurnal: `Cereri/motor_log.txt`.
