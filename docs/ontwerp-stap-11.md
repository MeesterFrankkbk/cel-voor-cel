# Ontwerp stap 11 – Klaar voor een gebruiker

**Kern:** een werkblad dat je aan iemand anders geeft, moet **tegen een stootje kunnen**. De gebruiker ziet zijn score, mag een spatie te veel typen, ziet de oplossingen niet en kan niets kapotmaken. De leerling denkt hier voor het eerst echt na over **iemand anders** die zijn werkblad gebruikt.

## Lessen

| Les | Wat de leerling ontdekt | Oefenvorm |
|---|---|---|
| De score tellen | `=AANTAL.ALS(D4:D9;"juist")` telt hoe vaak "juist" voorkomt | mini-rekenblad |
| Spaties weghalen | `SPATIES.WISSEN` haalt spaties weg; **valkuil**: het maakt van een getal tekst, dus "42" ≠ 42. Daarom rond **beide** kanten: `SPATIES.WISSEN(C4)=SPATIES.WISSEN(B4)` | mini-rekenblad met "Rome " en 42 |
| Een afbeelding | Invoegen → Afbeeldingen; in Excel 365 ook "Plaatsen in cel" | stappen |
| Oplossingen verbergen | kolom verbergen; rasterlijnen uit (Beeld) | quiz |
| Beveiligen | eerst de antwoordcellen **ontgrendelen** (Celeigenschappen → Bescherming), dan het blad **beveiligen** (Controleren). De volgorde is belangrijk! | beveiligings-simulator: ontgrendelen, beveiligen, uitproberen |

## Opdrachten in Excel (startbestand: werkblad met vragen, oplossingen en feedbackformules)

1. B11: tel het aantal "juist" met AANTAL.ALS.
2. B12: de score als procent (aantal juist / aantal vragen).
3. D4:D9: de formules verbeteren met SPATIES.WISSEN (aan beide kanten) en doortrekken.
4. Een afbeelding toevoegen.
5. Kolom B (de oplossingen) verbergen.
6. De rasterlijnen verbergen.
7. De antwoordcellen C4:C9 ontgrendelen (en de rest geblokkeerd laten).
8. Het blad beveiligen.

## Controle

- AANTAL.ALS en SPATIES.WISSEN: de app probeert de formules uit met andere waarden (alles juist, niets juist; een antwoord met een spatie erachter; een getal met een spatie erachter).
- Afbeelding: een afbeelding op het blad of in een cel.
- Verborgen kolom, rasterlijnen, geblokkeerde cellen en beveiliging worden rechtstreeks uit het bestand gelezen.

Foutmeldingen onder meer voor: SPATIES.WISSEN alleen rond het antwoord (getallen worden dan "jammer"), het hele blad ontgrendeld (dan kan de gebruiker alles wijzigen), beveiligd maar de antwoordcellen niet ontgrendeld (dan kan de gebruiker niets invullen), de verkeerde kolom verborgen.
