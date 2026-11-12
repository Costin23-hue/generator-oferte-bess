# Generator Oferte BESS – note pentru Claude

Program Python/Tkinter (`generator_oferte_bess.py`) care face oferte BESS: Excel de calcul (template), prezentare
PPTX + PDF, din Pricing.xlsx (al lui Levon, doar citire) și „Setari generator.xlsx” (Drive, Ofertare BESS).
Utilizatorul (Costin) vorbește română; răspunsurile în română, scurte, cu tabele unde ajută.
- Răspunsuri FOARTE scurte: ce s-a schimbat + doar ce trebuie să facă el de mână (fișiere pe care nu le pot modifica).
  Fără explicații lungi/liste de verificări (testează el). Nu afișa în chat output mare (base64, fișiere întregi):
  redirecționează în fișier și citește doar ce trebuie.

## Structura codului
- `generator_oferte_bess.py` = importuri + CONFIGURARE (constante/setări) + pornire; încarcă cu exec, în ACELAȘI
  spațiu de nume, părțile din `MODULE_PROGRAM`: bess_setari (setări, curs, auto-update), bess_logica (Pricing, baterii,
  skiduri, aux), bess_excel, bess_prezentare (PPTX), bess_drive (Graph, BESS_Proposals), bess_generare (calcul/finalizare
  ofertă), bess_atr (cititor ATR), bess_motor (Power Apps), bess_fereastra (Tkinter). Fișier nou -> și în FISIERE_PROGRAM.
- În bess_*.py nu se fac importuri (sunt în fișierul principal); testele folosesc `g = generator_oferte_bess`.

## Reguli de lucru
- Nu se trimite nicio ofertă pe e-mail. Nimic nu se scrie în BESS_Proposals fără OK-ul utilizatorului.
- Repository-ul e PUBLIC: fără prețuri reale, clienți, ATR-uri sau date din Pricing în cod/teste (teste = date inventate).
- După orice modificare: `python3 -m unittest discover -s teste` (trebuie să treacă), commit + push pe `main`.
  Utilizatorii iau versiunea nouă din butonul verde „Versiune nouă – Actualizează” (auto-update din GitHub, doar .py).
- Commit-urile se termină cu trailer-ul Co-Authored-By / Claude-Session din instrucțiunile sesiunii.
- Mac: doar Python (instalare_mac.sh); fără Excel/PowerPoint se folosește modul Microsoft 365 (Excel Online + PDF din Drive).

## Reguli de ofertare implementate
- Baterii în listă: toate Jinko „2H-” (inclusiv variantele cu rack-uri scoase) + REPT; REPT = 4 ore; „Auto” = cea mai ieftină.
- PCS pe baterie 1:1 sau 1:2; REPT pe TBEA (rack) sau Kehua (dc); TBEA doar la POC 20 kV.
- Consum aux: doar bateriile, 50 kVA (<5,6 MWh) / 65 kVA. Skidul are trafo propriu separat pentru consumul lui
  (CONSUM_SKID_IN_NECESAR = NU; tabelul KVA_CONSUM_SKID 30/40/50/60 kVA rămâne pentru DA). Trafo-ul aux din numele
  skidului acoperă bateriile.
- Trafo aux extern: cât mai puține bucăți (1 x 1250 în loc de 3 x 450), apoi cel mai ieftin; trafo aux inclus în skid se scade.
- EMS după capacitatea OFERTATĂ (prag 15 MWh: 3500 / 6500 USD), nu cea cerută.
- Limitare prin EMS: skidul poate fi peste evacuarea aprobată, în prezentare „6 (limited by EMS to 5,85)”.
- Commissioning baterie: din Pricing (Battery 2027 delivery, K..O = 1,2,3,4,5+ containere) / unitate x nr. baterii -> G11.
- Rezervă de preț (% baterie / skid) în tab-ul Componente manuale; în Excel ca formulă.
- „Putere max. admisă (MW)”: skidurile peste ea nu se propun (din ATR = Pi stocare).
- Flux: calcul Excel -> tab „Verificare calcul” -> OK -> PPTX + PDF; „test/test” = calcul de probă, se șterge
  (la OK pe probă: dialog client/proiect/locație/nr. -> se salvează ca ofertă reală, cu Excelul de probă copiat).
- Versiunea (_vN) crește doar după o prezentare: calculele refăcute înainte de OK stau în folderul aceleiași versiuni
  (mai multe Excel-uri); la OK folderul se redenumește dacă s-a schimbat MWh / MW.
- Oferte în Drive pe zile: Oferte generate/AAAA-LL-ZZ/<oferta>.
- BESS_Proposals: numele proiectului „<Proiect>_BESS, <MWh>MWh, <MW>MW”.

## Cititorul de ATR (tab „Citire ATR”, funcția `citeste_atr_text`)
- Citește: operator, nr./data, beneficiar, POC kV, putere aprobată (evacuare/absorbție, kW/kVA), tabelul stocării
  (Pi IS, Pmax evac/abs IS, Ah, MWh), cuplare DC, legare pe JT (TDRI), fără absorbție, trafo-uri, ATR fără stocare.
- Formate testate: Electrica 2023 și 2026 (pypdf lipește valorile kW de eticheta următoare), avize CTE de studiu de
  soluție, Rețele Electrice (kVA înaintea kW, „-” pentru etape goale; consumator; tabel stocare în MW), studiu de
  soluție (tabel normă tehnică MW), Distribuție Energie Oltenia (antet „ATR nr. X / data” pe fiecare pagină, unități
  fără paranteze, POC „0/20000/0 V”, uneori lipsește rândul „evacuată”), Transelectrica (draft), E-Distribuție,
  caiet de sarcini (putere min. instalată, string), aviz de amplasament (nu e ATR).
- Capacitate doar în Ah: MWh = Ah x tensiunea bateriei (TENSIUNE_BATERIE_V în setări sau câmpul din tab-ul ATR).
- Buton „🤖 Citește cu AI”: PDF-ul trimis la Claude (MODEL_AI, implicit claude-sonnet-5-5; cheie API locală în setari.json
  „cheie_claude” sau ANTHROPIC_API_KEY), structured output -> același dicționar (citeste_atr_ai / rezultat_din_ai).
- Scanate (fără text) -> doar avertisment (fără OCR deocamdată). Lipsesc exemple E-Distribuție și Delgaz în format text.
- Regula utilizatorului: puterea skidului poate depăși evacuarea (limitare EMS), dar nu trafo-urile din parc / Pi stocare.

## Întrebări deschise (utilizatorul testează întâi)
- „Putere max. admisă” strictă, doar avertisment, sau max(Pi stocare, evacuare)? (caz: ofertat 2,5 MW la Pi 2,4 MW)
- La neconcordanțe Pi (tabel vs text): acum se ia valoarea din tabel + ⚠.
- Cuplare DC/DC în ATR, dar ofertat MV skid AC: ce mesaj vrea utilizatorul.
- De făcut mai târziu: fișa de prețuri editabilă în aplicație, toate versiunile în același Excel, împărțirea codului.
