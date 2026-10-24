# Ontwerp stap 10 – Keuzelijsten en kleuren die reageren

**Kern:** een werkblad wordt gebruiksvriendelijk als het de gebruiker **helpt** (een keuzelijst, een foutmelding bij een onmogelijk antwoord) en **toont** wat er gebeurt (groen bij juist, rood bij jammer). Hetzelfde geldt voor het rapport: in één oogopslag zie je wat goed ging en wat minder.

## Lessen (met twee nieuwe simulators)

| Les | Wat de leerling ontdekt | Oefenvorm |
|---|---|---|
| Een keuzelijst | Gegevens → Gegevensvalidatie → Toestaan: Lijst; bron als losse woorden met ; of als bereik | nagebootst venster Gegevensvalidatie + de keuzelijst zelf uitproberen |
| Alleen goede antwoorden toelaten | Geheel getal tussen … en …; tabblad Foutmelding | hetzelfde venster; daarna een fout getal typen en de eigen foutmelding zien |
| Kleuren die reageren | Start → Voorwaardelijke opmaak → Markeringsregels: Kleiner dan, Groter dan, Gelijk aan | regelbouwer: de kleuren verschijnen meteen |
| Groen voor juist, rood voor jammer | Tekst die bevat; twee regels op hetzelfde bereik | regelbouwer met juist/jammer/vul in |
| Kleuren in je rapport | procenten: 50% of 0,5 typen in de regel | regelbouwer met procenten + tip: geef ook je eigen rapport uit stap 8 kleuren |

## Opdrachten in Excel

Startbestand met twee bladen: **Werkblad** (zes vragen, de juist/jammer-formules staan er al in) en **Rapport** (een kant-en-klaar rapport met procenten).

1. C5: keuzelijst met losse woorden (bv. word, wordt, werd, worden).
2. C6: keuzelijst met als bron het bereik K6:N6.
3. C4: alleen gehele getallen tussen 0 en 1000, met een eigen foutmelding.
4. D4:D9: "juist" kleurt groen.
5. D4:D9: "jammer" kleurt rood.
6. Rapport: procenten lager dan 50% kleuren rood.
7. Rapport: procenten hoger dan 80% kleuren groen.
8. Het werkblad testen: de drie eerste vragen juist beantwoorden met de keuzelijsten.

## Controle

De app leest de regels uit het bestand en **probeert ze uit**: ze kijkt welke kleur een cel zou krijgen bij "juist", bij "jammer", bij 45% of bij 85%. Ze herkent zowel de kant-en-klare opmaak ("Lichtrode opvulling met donkerrode tekst") als zelf gekozen kleuren, als opvulkleur of als tekstkleur.

Foutmeldingen voor: keuzelijst op de verkeerde cel, ontbrekende of extra keuzes, losse woorden in plaats van het bereik, geen foutmelding, verkeerde grenzen, regel op het verkeerde bereik, "Groter dan" in plaats van "Kleiner dan", 50 in plaats van 50% of 0,5, de verkeerde kleur.
